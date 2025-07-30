from PIL import Image as Im
import pytesseract
from pytesseract import image_to_data
import io

from pathlib import Path
from pdf2image import convert_from_path
import copy

from img2table.document import Image
from img2table.document import PDF
from img2table.ocr import TesseractOCR
import fitz
import re

import pandas as pd
import openpyxl

import subprocess
import ollama
from ollama import chat
from ollama import ChatResponse
#https://ollama.com/blog/structured-outputs
from pydantic import BaseModel, create_model
import json
from typing import get_type_hints
from enum import Enum

import os
import shutil
import time
import tkinter as tk
from tkinter import filedialog

class Attendance(str, Enum):
    Yes = "Yes"
    No = "No"
    WS = "Written submission received"
    WS0 = "Written hearing but no written submission received"

class Order(BaseModel):
    Address: str
    FileNumber: str
    LandlordName: str
    DateOfHearing: str
    City: str
    PostalCode: str
    FirstEffectiveDate: str
    WeightedUsefulLife: list[float]
    LandlordRate: float
    AGIRateTotal: float
    CapitalExpenditure: list[str] | None
    LandlordAttendance: Attendance
    LandlordLegalAttendance: Attendance
    TenantAttendance: Attendance
    TenantLegalAttendance: Attendance
    InterestingNotes: str | None


system_message = {
    "role": 'system',
    "content": "You extract rental increase data from strings, with no extra text."}

splitColumns = ["TotalIncreaseexcludesguideline",
                "%incforCapExp"]

totalColumns = ["WeightedUsefulLifeforCapitalExp*",
                 "Total%forCapExp"]

models = ["phi3:mini", "qwen3:8b", "deepseek-r1:7b", "llama3.2", "deepseek-r1:1.5b", "qwen3:1.7b"]
hyperparameters = {'temperature': 0.0,
                   'top_p': 0.0,
                   'repeat_penalty': 1.0,
                   'num_predict': 4096}

possibleColumns = {"WeightedUsefulLifeforCapitalExp*": "WeightedUsefulLife",
                   "Total%forCapExp": "AGIRateTotal",
                   "TotalIncreaseexcludesguideline1": "AGIRate1",
                   "%incforCapExp1": "AGIRate1",
                   "TotalIncreaseexcludesguideline2": "AGIRate2",
                   "%incforCapExp2": "AGIRate2",
                   "TotalIncreaseexcludesguideline3": "AGIRate3",
                   "%incforCapExp3": "AGIRate3",
                   "TotalIncreaseexcludesguideline4": "AGIRate4",
                   "%incforCapExp4": "AGIRate4",
                   "TotalIncreaseexcludesguideline5": "AGIRate5",
                   "%incforCapExp5": "AGIRate5"}

regexDict = {r'(?<=In the matter of:\s)(.*?)(?=\s*Between:)': "Address",
             r'Between:\s*(.*?)\s*Landlord\s*(.*?)\s*and\s*Refer': "LandlordName",
             r'File Number:\s*([A-Z]{3}-\d{5}-\d{2})' : "FileNumber",
             #r'\b([A-Z]\d[A-Z]\d[A-Z]\d)\b(?=[\s\n]*Between:)': "PostalCode",
             r'([A-Z]\d[A-Z]\d[A-Z]\d)': "PostalCode",
             r'This application was heard in (.*?) on': "City",
             r'This application was heard in .+? on (.+?)\.': "DateOfHearing"
             }

otherDict = {r'First Effective Date of Rent Increase in this Order is\s+((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4})': "FirstEffectiveDate",
             r'\b\d+(?:\.\d+)?%': "LandlordRate"}
#in the case anyone needs them and doesn't want to comb through the mess below

parsingDict = {"LandlordAttendance": "The following parties",
               "LandlordLegalAttendance": "The following parties",
               "TenantAttendance": "The following parties",
               "TenantLegalAttendance": "The following parties",
               "FirstEffectiveDate": "effective date",
               "LandlordRate": "Landlord justified",
               "InterestingNotes": "withdraw",
               "CapitalExpenditure": "apital expenditure"
}

leftoverRegex = {"LandlordAttendance": r"\(The following parties.*?\)(?=\s*It is determined that:)",
                 "LandlordLegalAttendance": r"\(The following parties.*?\)(?=\s*It is determined that:)",
                 "TenantAttendance" : r"\(The following parties.*?\)(?=\s*It is determined that:)",
                 "TenantLegalAttendance": r"\(The following parties.*?\)(?=\s*It is determined that:)"
                 }

fieldsRegex = {"Address", "FileNumber", "LandlordName", "DateOfHearing", "City", "PostalCode"} #regex
miscellaneousFields = {"FirstEffectiveDate", "LandlordRate", "FileName"}
tableFields = {"AGIRateTotal", "AGIRate1", "AGIRate2", "AGIRate3", "AGIRate4", "AGIRate5", "WeightedUsefulLife"} #tables
ollamaFields = {"CapitalExpenditure", "InterestingNotes",
                "LandlordAttendance", "LandlordLegalAttendance","TenantLegalAttendance", "TenantAttendance"} #ollama
#totalFields = fieldsRegex.union(ollamaFields).union(tableFields)
totalFields = ["FileName", "Address", "LandlordName", "FileNumber", "DateOfHearing", "City", "PostalCode", "FirstEffectiveDate", "LandlordRate",
               "AGIRateTotal", "AGIRate1", "AGIRate2", "AGIRate3", "AGIRate4", "AGIRate5", "WeightedUsefulLife", "CapitalExpenditure", "InterestingNotes",
                "LandlordAttendance", "LandlordLegalAttendance","TenantLegalAttendance", "TenantAttendance"]

questionsContext = "You are an assistant that extracts data from long strings of rent increase court orders." +\
"You keep answers as short as possible (a couple words per field maximum, (one word if possible, your output tokens are limited to 1024) and don't repeat entries for the same field." +\
"The address consists of a number and the street name. It typically ends in 'Road' or 'Street' or 'Drive'.\n" +\
                   "The city is the city of the address which the hearing is concerning." +\
                       "The file number is in the format XXX-DDDDD-DD, where X is a letter and D is a digit." +\
                       "The landlord name immediately precedes '('the Landlord')', and is fully capitalized in the text." +\
                       "The date of hearing is when the hearing was heard. It should be in YYYY-MM-DD format. The city is usually listed with it, but it is not always capitalized." +\
                       "The postal code is usually listed in the address, and it is 6 characters long, in the format XDXDXD or XDX DXD" +\
                       "The first effective date should be in YYYY-MM-DD format, and it is when the rent increase goes into effect."+\
                       "The weighted useful life of capital expenditures should be listed. If there are multiple for different repairs or expenses, include all of them." +\
                       "The landlord rate is the rent increase initially justified by the Landlord." +\
                       "The AGIRateTotal is the total percent increase in rent over all years specified in the text. If it doesn't mention different years, this quantity is just the increase" +\
                       "The AGIRateByYear is an ordered list of the maximum rent increases each year, in percents. If there are n years specified, this should be n elements long." +\
                       "Capital expenditures are the repairs or improvements to the property. If not explicitly stated, use 'Not stated'. If not at all mentioned, use None." +\
                       "The capital expenditures strings should be a couple words maximum." +\
                       "The attendance fields are as such: Yes = 1; No =0; WS = written submission received; WS0 = written hearing but no written submission received. Keep in mind that names may have been censored, so if Tenants or Units are mentioned, that is very likely them being in attendance."+\
                       "For fields with legal attendance, assume ANY representation for that party is legal representation." +\
                       "The interesting notes is a list of the following: Were units/tenants removed from the application? Anything else of note? Any pro forma language?"

def FinalApproach(file, trial):
    #setup
    pages = getPagesBySchedule(file)
    fileContents = getOCR(pages, 0, file)
    #print(fileContents)
    BaseAnswers = {}
    BaseAnswers["FileName"] = file.stem
    
    #Regex
    BaseAnswers = Regex(fileContents, BaseAnswers)

    x = getListOfStringSeparatedThings(BaseAnswers["Address"])
    ConstructedAnswers = []
    firstPass = True
    for address in x:
        answers = copy.deepcopy(BaseAnswers)
        answers["Address"] = address
        answers = Table(file, pages, answers, len(x)!=1)
        
        answers = Miscellaneous(pages, file, fileContents, answers)

        print("Starting LLMCall")
        answers = LLMOneshot(fileContents, answers)
        print("Finished Oneshot")
        for key, value in answers.items():
            if value is None or value == []:
                answers[key] = "FLAGGED FOR REVIEW"
            elif isinstance(value, list):
                answers[key] = ', '.join(map(str, value))  # Turn list into readable string
            elif isinstance(value, dict):
                answers[key] = str(value)
            else:
                answers[key] = value

        ConstructedAnswers.append(answers)
        displayDict(answers)
        
    return ConstructedAnswers

def getListOfStringSeparatedThings(addressString: str):
    address_pattern = re.compile(
        r'(\d+\s+[^,]+,\s+[^,]+,\s+[A-Z]{2},\s+[A-Z0-9]{6})', re.IGNORECASE)
    
    # Find all complete addresses
    matches = list(address_pattern.finditer(addressString))
    
    if not matches:
        return [addressString.strip()] if addressString.strip() else []
    
    result = []
    
    for match in matches:
        address = match.group(1).strip()
        if address:
            result.append(address)
    
    print(result)
    return result

def getPagesBySchedule(file):
    pageArray = {0: [], 3: [], 4: [], 1:[], 2:[]}
    fileImages = convert_from_path(file, dpi=400)
    last = 0;

    i = 0
    for image in fileImages:
        
        #top strip (from top to -, full width, maybe 200px tall)
        width, height = image.size
        top_strip = image.crop((0, 0, width, 500))
        
        text = pytesseract.image_to_string(top_strip)
        text = text.strip()  # cleanup whitespace/newlines


        for j in range(1, 5):
            if "Schedule " + str(last + j) in text:
                last = last + j

        if last in (0,1,2,3,4):
            pageArray[last].append(i)
        i = i + 1

    print(pageArray)
    return pageArray
    

def create_filtered_model(base_model, exclude_fields):
    annotations = get_type_hints(base_model)
    
    filtered_annotations = {
        field_name: field_type 
        for field_name, field_type in annotations.items() 
        if field_name not in exclude_fields
    }
    
    return create_model(
        f"{base_model.__name__}Filtered",
        **filtered_annotations
    )

def LLMJSON(text, excludeFields, additional):
    FilteredOrder = create_filtered_model(Order, excludeFields)
    print(questionsContext + text)
    print("Filtered")
    subprocess.run(["ollama", "stop", singleModel])
    response = chat(
        messages = [
            {'role': 'user',
             'content': text + f"Fill out the JSON according to the address being {additional}" + 
             "Return a JSON object. Be concise and add no extra words.",
            }],
        model = singleModel,
        format=FilteredOrder.model_json_schema(),  # Use filtered model schema
        options = hyperparameters,
        stream=False
        )
    return response['message']['content']

def LLMOneshot(text, answers):
    fields_to_exclude = set(answers.keys())
    try:
        json_response = LLMJSON(text, fields_to_exclude, answers["Address"])
    except Exception:
        return answers
    parsed_response = json.loads(json_response)

    for key, value in parsed_response.items():
        answers[key] = value
    return answers

def Miscellaneous(pages, file, fileContents, answers):
    for method in [FirstEffectiveDate, LandlordRate]:
        answers = method(pages, answers, file, fileContents)
    return answers

def FirstEffectiveDate(pages, answers, file, fileContents):
    context = getOCR(pages, 3, file)
    x = regexSmall(context, r'First Effective Date of Rent Increase in this Order is\s+((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4})')
    if x == None:
        x = regexSmall(context, r'((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4})')#just match for ANY date     
    answers["FirstEffectiveDate"] = x
    return answers

def getMax(answers):
    m = 0
    for i in range(1, 6):
        x = (answers["AGIRate"+str(i)]).keys()
        if x:
            m = m + max(x)
    return m

def LandlordRate(pages, answers, file, fileContents):
    context = getOCR(pages, 0, file)
    matches = re.findall(r'\b\d+(?:\.\d+)?%', context)
    numbers = [float(m.strip('%')) for m in matches]
    numbers.append(getMax(answers))
    if numbers:
        answers["LandlordRate"] = max(numbers)
    else:
        answers["LandlordRate"] = 0
    return answers

def AddAGIFields(answers):
    for i in range(5):
        if "AGIRate" + str(i+1) not in answers.keys():
            answers["AGIRate" + str(i+1)] = {}
    return answers
    

def Table(file, pageNo, answers, multiple = True): #https://github.com/xavctn/img2table, pretty much exactly what I did and packaged so I'll use this
    # Validate inputs
    if file is None:
        raise ValueError("File parameter is None")
    
    if not pageNo[3]:
        print("Warning: No pages found for schedule 3")
        return AddAGIFields(answers)
        
    # Convert Path to string and validate file exists
    file_path = str(file)
    if not Path(file_path).exists():
        raise FileNotFoundError(f"File does not exist: {file_path}")
    
    start_page = pageNo[3][0]
    end_page = pageNo[3][-1]
    
    pages = list(range(start_page, end_page + 1))
    pdf = PDF(file_path, pages=pages, detect_rotation=False, pdf_text_extraction=True)
    
    ocr = TesseractOCR(lang="eng")
    tables = pdf.extract_tables(ocr=ocr)
    pdf.to_xlsx('tables.xlsx', ocr=ocr)
    table = tableJoiner(tables, pages)
    table.to_excel("output_table.xlsx", index=False)

    if multiple:
        answers = tableToFields(table, answers)
    else:
        answers = tableToField(table, answers)
    return AddAGIFields(answers)

def tableToFields(df, answers, columnToReference="unit"):
    # VERY IMPORTANT!!! column names to lowercase
    df.columns = df.columns.str.lower()
    
    # Convert possibleColumns keys to lowercase for matching
    possibleColumns_lower = {k.lower(): v for k, v in possibleColumns.items()}
    
    match = re.search(r'(\d+\s[\w\s]+),', answers["Address"])
    if match:
        cleaned_address = match.group(1).strip()
    else:
        cleaned_address = ""
    print(cleaned_address)
    
    if columnToReference in df.columns and cleaned_address:
        df_filtered = df[df[columnToReference].astype(str).str.contains(re.escape(cleaned_address), case=False, na=False)]
        # If no matches found after filtering, use the full dataframe
        if df_filtered.empty:
            print(f"Warning: No matches found for address '{cleaned_address}' in unit column. Using full table.")
            df_filtered = df
    else:
        df_filtered = df
    df_filtered.to_excel("output_table.xlsx", index=False)
    
    for column in df_filtered.columns:
        if column in possibleColumns_lower and possibleColumns_lower[column] not in answers.keys():
            numeric_series = pd.to_numeric(df_filtered[column], errors='coerce').dropna()
            if not numeric_series.empty:
                answers[possibleColumns_lower[column]] = numeric_series.value_counts().to_dict()
                
    return answers


def tableToField(df, answers):
    for column in df.columns:
        if column in possibleColumns and possibleColumns[column] not in answers.keys():
            numeric_series = pd.to_numeric(df[column], errors='coerce').dropna()
            answers[possibleColumns[column]] = numeric_series.value_counts().to_dict()
    return answers

def tableJoiner(tables, pages):
    FullDF = pd.DataFrame()
    for i in pages:
        df = pd.DataFrame()

        
        year = {"TotalIncreaseexcludesguideline":0,
                    "%incforCapExp":0}
        for j in tables[i]:
            removed = 0
            frame = j.df
            frame.columns = frame.columns.astype(str).str.replace(r'[()\d\s.]+', '', regex=True)
            while (not any(col in frame.columns for col in splitColumns + totalColumns) and removed < 5 and len(frame) > 0):
                frame.columns = frame.iloc[0]
                frame = frame[1:]
                frame.columns = frame.columns.astype(str).str.replace(r'[()\d\s.]+', '', regex=True)
                removed += 1

            #this regex simply removes digits, whitespace, and .'s.
            frame.columns = frame.columns.astype(str).str.replace(r'[()\d\s.]+', '', regex=True)

            selected_cols = []
            column_map = {}

            for col in frame.columns:
                col_lower = col.lower()
                if col in totalColumns:
                    selected_cols.append(col)

                if col in splitColumns:
                    year[col] = year[col] + 1
                    column_map[col] = f"{col}{year[col]}"
                    selected_cols.append(col)

                if "unit" in col_lower:
                    selected_cols.append(col)
                    
            frame = frame[selected_cols]
            frame = frame.rename(columns = column_map)

            for col in frame.columns:
                try:
                    frame[col] = pd.to_numeric(frame[col])
                except Exception:
                    pass
            
            df = pd.concat([df, frame], axis=1)
            df = df.loc[:, ~df.columns.duplicated()]
        FullDF = pd.concat([FullDF, df], axis=0, ignore_index=True)
    return FullDF

def getOCR(pages, scheduleNo, file):
    doc = fitz.open(file)
    all_text = ""

    for page_index in pages[scheduleNo]:
        page = doc[page_index]
        selectable_text = page.get_text().strip()

        ocr_text = ""
        for img in page.get_images(full=True):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image = Im.open(io.BytesIO(image_bytes))
            ocr_text_part = pytesseract.image_to_string(image).strip()
            if ocr_text_part:
                ocr_text += ocr_text_part + "\n"
            image.close()

        combined = selectable_text
        if ocr_text:
            if ocr_text not in selectable_text:
                combined = (selectable_text + "\n" + ocr_text).strip()
        all_text = all_text + combined + "\n\n"

    doc.close()
    if scheduleNo != 3:
        print(all_text)
    return all_text

def Regex(text, answers):
    for regex in regexDict.keys():
        answer = regexSmall(text, regex)
        if answer:
            answers[regexDict[regex]] = answer
            
        if set(regexDict.values()).issubset(answers.keys()):
            break
            
        #check if dictionary contains all fields, break if so
        #in the worst case when the regexDict ever gets large (unlikely)
        #significant room for optimization if regexDict ever does get large
        #fortunately, this is extremely unlikely for now
    return answers

def regexSmall(text, pattern):
    match = re.search(pattern, text, re.DOTALL)
    if match:
        if len(match.groups()) > 1:
            result = ' '.join(match.groups())
        else:
            result = match.group(1)
        result = re.sub(r'\s+', ' ', result.replace('\n', ' '))
        return result.strip()
    return None

def addRow(JSON, databasePath, sheetName, allowRepeats = True):
    wb = openpyxl.load_workbook(databasePath)

    if sheetName not in wb.sheetnames or not allowRepeats:
        ws = wb.create_sheet(title=sheetName)
        ws.append(list(totalFields))
    else:
        ws = wb[sheetName]

    if ws.max_row == 0 or all(cell.value is None for cell in ws[1]):
        ws.append(list(totalFields))

    for col in totalFields:
        if col not in JSON:
            JSON[col] = ""

    newRow = []
    for col in totalFields:
        val = JSON[col]
        if isinstance(val, dict):
            val = str(val)
        newRow.append(val)

    ws.append(newRow)
    wb.save(databasePath)
    print(f"Appended {JSON.get('FileNumber', '[no FileNumber]')} to {databasePath}")

def checkExcel(dbPath, file_name):
    df = pd.read_excel(dbPath)
    if df.empty or 'FileName' not in df.columns:
        return False
    match_df = df[df['FileName'] == file_name]
    return not match_df.empty

def displayDict(dct):
    for k,v in dct.items():
        print(f"{k}: {v}")
    print("---------------------------------------------------------------")

def FinalApproachWrapper(file, folderPath, newFolderPath, localDatabasePath, newDatabasePath, errorFolderPath, trial = None):
    sheetName = "sample"
    try:
        print(file.stem)
        if trial or not checkExcel(localDatabasePath, file.stem):
            start = time.perf_counter()
            answers = FinalApproach(file, trial)
            print("\n")
            for answer in answers:
                displayDict(answer)
            for x in answers:
                if "toAdd" not in x.keys():
                    n = addRow(x, localDatabasePath, sheetName)
                    m = addRow(x, newDatabasePath, sheetName)

            end = time.perf_counter()
            elapsed = end - start
            print(f"Total execution time: {elapsed:.2f} seconds")
        shutil.move(str(file), newFolderPath / file.name)
    except Exception as e:
        print(f"Error processing {file.name}: {e}")
        shutil.move(str(file), errorFolderPath / file.name)
        

def main(trial):
    print("Entering main")
    sheetName = 'sample'
    if True:
        root = tk.Tk()
        root.withdraw()
        root.update()
        folderPath = Path(filedialog.askdirectory())
        f = filedialog.askopenfilenames()
        localDatabasePath, newDatabasePath = Path(f[0]), Path(f[1])
        newFolderPath = Path(filedialog.askdirectory())

    parentDir = newFolderPath.parent
    errorFolderPath = parentDir / "Excessively long"
    errorFolderPath.mkdir(parents=True, exist_ok=True)

    start_time = time.time()
    time_array = []
    print("Process started.")
    folder = [file for file in folderPath.glob("*.pdf") if (("Order" in file.name) or ("Hearing" in file.name) or ("Redacted" in file.name))]

    for file in folder[2:]:
        FinalApproachWrapper(file, folderPath, newFolderPath, localDatabasePath, newDatabasePath, errorFolderPath)

    end_time = time.time()
    elapsed_time = end_time - start_time
    print(time_array)
    print(f"Total execution time: {elapsed_time:.2f} seconds")    
    
if __name__ == "__main__":
    singleModel = models[0]
    main(0)

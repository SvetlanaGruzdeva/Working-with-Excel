# vlookup - program performs lookup in the same manner as Excel formula 'VLOOKUP' does.
# Working range consists of 3 columns - key, English value and Russian value.

import fnmatch, os, pandas as pd
from pprint import pprint
from tkinter.filedialog import askopenfilename

myFile = os.path.abspath(askopenfilename()) # Selected by user from browser

def get_lookup_area(excelFile):
    ws = pd.read_excel(excelFile, sheet_name='Map', header=2)
    wb = pd.ExcelFile(excelFile)
    df = wb.parse('Map')
    # Find needed column by header
    header = df.iloc[1].tolist()
    header = [str(i) for i in header]
    pattern = '*cost center*'
    matching = fnmatch.filter(header, pattern)
    columnIndex = header.index(matching[0])
    # Transfer needed column into list
    ccList = ws[matching[0]].unique().tolist()
    # Set up lookup area as dataframe
    df = pd.read_excel(myFile, sheet_name='Map', skiprows = 2, usecols = (columnIndex, columnIndex+1, columnIndex+2))
    # Create dictionaries for future lookup
    ccDictEng = {}
    ccDictRus = {}
    row = 0
    for costCenter in ccList:
        ccDictEng[costCenter] = str(df.iat[row, 1]).replace('  ', ' ').rstrip()
        ccDictRus[costCenter] = str(df.iat[row, 2]).replace('  ', ' ').rstrip()
        row +=1
    return(ccDictEng, ccDictRus)

def get_rcc_instructions(excelFile):
    ws = pd.read_excel(excelFile, sheet_name='RCC', header = 5, usecols = (1, 2))
    rccCode = list(ws.iloc[0:, 0])
    rccDescr = list(ws.iloc[0:, 1])
    # Split RCC description to English and Russian descriptions
    rccDescrEng = []
    rccDescrRus = []
    for rcc in rccDescr:
        rccDescrEng.append(rcc.split(';\n')[0].replace('  ', ' ').rstrip())
        rccDescrRus.append(rcc.split(';\n')[1].replace('  ', ' ').rstrip())
    # Create dictionaries for future lookup
    rccDictEng = {}
    rccDictRus = {}
    n = 0
    for rcc in rccCode:
        rccDictEng[rcc] = rccDescrEng[n]
        rccDictRus[rcc] = rccDescrRus[n]
        n +=1
    return(rccDictEng, rccDictRus)

correctNamesEng = get_rcc_instructions(myFile)[0]
correctNamesRus = get_rcc_instructions(myFile)[1]

toCheckEng = get_lookup_area(myFile)[0]
toChecRus = get_lookup_area(myFile)[1]

open('P:\\Documents Svetlana\\Python\\result.txt', 'w').close() # Clean result of previous run
f = open('P:\\Documents Svetlana\\Python\\result.txt', 'w')
for key in correctNamesEng:
    if correctNamesEng[key] != toCheckEng[key]:
        print(key + ': "' + toCheckEng[key] + '" to be updated to: "' + correctNamesEng[key] + '"', file=f)
    if correctNamesRus[key] != toChecRus[key]:
        print(key + ': "' + toChecRus[key] + '" to be updated to: "' + correctNamesRus[key] + '"', file=f)
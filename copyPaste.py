# CopyPaste - copying and pasting specified ranges in template

import os, logging, openpyxl as xl, win32com.client as win32
from pathlib import Path
from tkinter.filedialog import askopenfilename, askdirectory
from pprint import pprint

def copyPaste(userFolder, templateFile):
    logging.basicConfig(filename=(str(userFolder) + 'logs.txt'), level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')
    f = open(userFolder + 'result.txt', 'w')
    for file in os.listdir(userFolder):
        fileName = os.fsdecode(file)
        logging.debug(fileName)
        if (fileName.endswith(".xlsm") or fileName.endswith(".xlsx")) and not fileName == Path(templateFile).name:
            if ('To be uploaded') not in os.listdir(userFolder):
                newFolder = os.makedirs(userFolder + 'To be uploaded')
                newFolder = (userFolder + 'To be uploaded\\')
            else:
                newFolder = (userFolder + 'To be uploaded\\')
            
            receivedFile = (userFolder + fileName)
            head, sep, tail = fileName.partition(templateFile[-9: -5])
            newFileName = (head + sep + '.xlsm')
            sourceWB = xl.load_workbook(receivedFile)
            sheetnames = sourceWB.sheetnames
            
            if 'Summary' in sheetnames:
                sourceWS = sourceWB.get_sheet_by_name('Summary')
                templateWB = xl.load_workbook(templateFile, keep_vba=True)
                templateWS = templateWB.get_sheet_by_name('Summary')
                # Copy and paste defined columns from received file to template.
                columnList = [2, 5, 8, 9, 11, 12, 14, 15, 17, 19, 21, 23, 24, 26, 27, 28, 29, 30,
                                                            31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41]
                # Identify first row to be copied from received file.
                for row in range(1, (sourceWS.max_row + 1)):
                    if sourceWS.cell(row=row, column=1).value == 'Сценарий':
                        firstRow = (row + 1)
                for column in columnList:
                    for row in range(firstRow, (sourceWS.max_row + 1)):
                        c = sourceWS.cell(row=row, column=column)
                        templateWS.cell(row=row, column=column).value = c.value
                        logging.debug(str(row) + str(column) + str(c.value))
                        # TODO: Copy data validation

                # Delete all unfilled rows.
                for row in range (templateWS.max_row, 6, -1):
                    if templateWS.cell(row=row, column=5).value == 'Please select' or templateWS.cell(row=row, column=5).value == 'Выберите из списка':
                        templateWS.delete_rows(row)
                templateWB.save(newFolder + newFileName)
                # os.remove(userFolder + fileName)
            else:
                errorFile = receivedFile.replace(userFolder, '')
                print('ERROR: Wrong tab name in sourse file ' + errorFile, file=f)
    print('\n', file=f)
    return(newFolder)

def Headers(templateFile):
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    wb = xl.Workbooks.Open(templateFile)
    ws = wb.Worksheets('Summary')
    headers = {}
    for column in range(1, ws.UsedRange.Columns.Count):
        headers[column] = ws.Cells(4, column).Value
    wb.Close(False)
    xl.Quit()
    return(headers)

def fileCheck(newFolder, Headers):
    f = open(str(Path(newFolder).parent) + '\\result.txt', 'a')

    for file in os.listdir(newFolder):
        fileName = os.fsdecode(file)
        newFile = (newFolder + fileName)
        xl = win32.gencache.EnsureDispatch('Excel.Application')
        wb = xl.Workbooks.Open(newFile)
        ws = wb.Worksheets('Summary')

        columnList = [3, 6, 10, 16, 18, 20, 22] # 13 is Residency code, temporary excluded
        for column in columnList:
            for row in range(5, ws.UsedRange.Rows.Count):
                if ws.Cells(row, column).Value == 'FORMULA':
                    print(str(fileName) + ': ERROR in column ' + str(Headers[column]), file=f)
                    break
        wb.Close(False)
        xl.Quit()
    return()

# userFolder = ('P:\\Documents Svetlana\\Excel training\\Marcos\\Regional templates\\Test\\')
# userFolder = (os.path.abspath(askdirectory()) + '\\')
# templateFile = (userFolder + 'G&A_planning_template_FCST2_2020.xlsm')
# templateFile = askopenfilename() # Selected by user from browser

fileCheck(copyPaste(userFolder=(os.path.abspath(askdirectory()) + '\\'), templateFile=askopenfilename()),
                                                                        Headers(templateFile=askopenfilename()))
# fileCheck(copyPaste(userFolder=userFolder, templateFile=templateFile), Headers(templateFile=templateFile))
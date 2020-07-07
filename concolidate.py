import os, logging, openpyxl as xl, win32com.client as win32
from pathlib import Path
from tkinter.filedialog import askopenfilename, askdirectory
from pprint import pprint

userFolder = 'P:\\Documents Svetlana\\Excel training\\Marcos\\Consolidate\\'
templateFile = 'P:\\Documents Svetlana\\Excel training\\Marcos\\Consolidate\\G&A_planning_template_FCST2_2020.xlsm'
# userFolder = (os.path.abspath(askdirectory()) + '\\')
# templateFile = askopenfilename()

logging.basicConfig(filename=(str(userFolder) + 'logs.txt'), level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')
# logging.disable(logging.CRITICAL)

lastRow = 6
templateWB = xl.load_workbook(templateFile, keep_vba=True)
templateWS = templateWB.get_sheet_by_name('Summary')
for file in os.listdir(userFolder):
    fileName = os.fsdecode(file)
    if (fileName.endswith(".xlsm") or fileName.endswith(".xlsx")) and not fileName == Path(templateFile).name:
        logging.info(fileName)
        receivedFile = (userFolder + fileName)
        sourceWB = xl.load_workbook(receivedFile)
        sourceWS = sourceWB.get_sheet_by_name('Summary')
        # Copy and paste defined columns from received file to template.
        columnList = [2, 5, 8, 9, 11, 12, 14, 15, 17, 19, 21, 23, 24, 26, 27, 28, 29, 30,
                                                    31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41]
        
        # logging.debug(lastRow)
        logging.debug(sourceWS.max_row)
        for row in range(6, (sourceWS.max_row + 1)):
            logging.debug(row)
            for column in columnList:
                c = sourceWS.cell(row=row, column=column)
                templateWS.cell(row=lastRow, column=column).value = c.value
            lastRow += 1
    # region = templateWS.cell(row=6, column=2).value
    # consolFile = templateFile.replace('G&A', 'Consol_' + region + '_G&A')
    templateWB.save(templateFile)


# Delete all unfilled rows.
# for row in range (templateWS.max_row, 6, -1):
#     if templateWS.cell(row=row, column=5).value == 'Please select':
#         templateWS.delete_rows(row)
#     templateWB.save(templateFile)

logging.info('END OF SESSION')
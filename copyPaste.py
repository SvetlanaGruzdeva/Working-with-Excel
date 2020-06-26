# CopyPaste - copying and pasting specified ranges in template
import os, openpyxl as xl
from pathlib import Path

# TODO: Loop receivedFile through content of specified folder. Folder to be selected by user.
userFolder = ('P:\\Documents Svetlana\\Excel training\\Marcos\\Regional templates\\Test\\')
for file in os.listdir(userFolder):
    filename = os.fsdecode(file)
    if (filename.endswith(".xlsm") or filename.endswith(".xlsx")) and \
                                not filename == 'G&A_planning_template_FCST2_2020.xlsm':
        print(filename)


# if not (userFolder + 'To be uploaded'):
#     newFolder = os.makedirs(userFolder + 'To be uploaded')
# else:
#     newFolder = (userFolder + 'To be uploaded\\')
# receivedFile = (userFolder + 'G&A_planning_template_Africa2_FCST2_2020.xlsm')
# sourceWB = xl.load_workbook(receivedFile)
# sheetnames = sourceWB.sheetnames
# if 'Summary' in sheetnames:
#     sourceWS = sourceWB.get_sheet_by_name('Summary')
#     templateFile = (userFolder + 'G&A_planning_template_FCST2_2020.xlsm')
#     templateWB = xl.load_workbook(templateFile, keep_vba=True)
#     templateWS = templateWB.get_sheet_by_name('Summary')

#     # Copy and pase defined columns from received file to template.
#     columnList = [2, 5, 8, 9, 11, 12, 14, 15, 17, 19, 21, 23, 24, 26, 27, 28, 29, 30,
#                                                 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41]
#     for column in columnList:
#         for row in range(6, sourceWS.max_row):
#             c = sourceWS.cell(row=row, column=column)
#             templateWS.cell(row=row, column=column).value = c.value
#     # Delete all unfilled rows.
#     for row in range (templateWS.max_row, 6, -1):
#         if templateWS.cell(row=row, column=5).value == 'Please select':
#             templateWS.delete_rows(row)
#     templateWB.save(newFolder + (Path(receivedFile).name))
# else:
#     f = open(userFolder + 'result.txt', 'w')
#     errorFile = receivedFile.replace(userFolder, '')
#     print('ERROR: Wrong tab name in sourse file ' + errorFile, file=f)
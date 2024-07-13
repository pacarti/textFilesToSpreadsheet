import openpyxl, os

os.chdir(os.path.dirname(os.path.abspath(__file__)))

sonnetFile = open('sonnet29.txt')

txtList = sonnetContent = sonnetFile.readlines()


txtToExcelWB = openpyxl.Workbook()

sheet = txtToExcelWB.active


for i, line in enumerate(txtList, start = 1): # By default it started from 0. Therefore we need to start from 1 since it's a spreadsheet.
    # print(str(i) + ': ' +  line)
    # print(type(i))
    sheet['A' + str(i)] = line

txtToExcelWB.save('resultSpreadsheetFromTxtFiles.xlsx')
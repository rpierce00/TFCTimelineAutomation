import openpyxl

def getExportFileNameAndOpenWorkbook ():
    excelFileName = input("Enter the file name of the excel workbook exported for the TFC Project Planner: \n")

    projectPlannerExportObj = openpyxl.load_workbook(excelFileName)

getExportFileNameAndOpenWorkbook()
from xlrd import *

def WriteToExcel(FileName, SheetName):
    AimFile = copy('网安分析表.xlsx')
    SheetList = AimFile.get_sheet()
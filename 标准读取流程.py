from openpyxl import load_workbook
from SubFunc import SubFunc


@SubFunc.safeOperation
def mainOperation(file_path):
    wb = load_workbook(file_path)
    ws = wb.active


def __main__():
    mainOperation("测试.xlsx")

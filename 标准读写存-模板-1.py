from openpyxl import load_workbook
from openpyxl.workbook.workbook import _WorksheetOrChartsheetLike
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell

filePath = "测试文件.xlsx"

currentWorkbook = load_workbook(filePath)
try:
    currentWorksheet = currentWorkbook.active
    if currentWorksheet is None:
        raise ValueError("工作表不存在")

    # 读取单元格的值
    cellA: Cell = currentWorksheet["A1"]
    print(cellA.value)

    # 写入单元格的值
    currentWorksheet["A1"] = "测试"

    theList = [
        ["标题1", "标题2", "标题3"],
        ["内容1", "内容2", "内容3"],
        ["内容4", "内容5", "内容6"],
    ]
    # 写入多行数据
    for row in theList:
        currentWorksheet.append(row)
    # 保存文件
    currentWorkbook.save(filePath)
finally:
    currentWorkbook.close()

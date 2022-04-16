import openpyxl
from openpyxl.styles import Alignment


def open_excel(filename, data_only=True):
    # data_only = true then
    try:
        document = openpyxl.load_workbook(filename, data_only=data_only)
        return document
    except FileNotFoundError:
        print("[open_excel] Cannot find: " + filename)
        raise
    # print(type(document))  # to test


def save_excel(wb, filename):
    print("Function excel.py-save_excel starts here")
    wb.save(filename)


def shift_right(sheet):
    """ 금액 range 은 F17 - F44, H17 - H33"""
    lst1 = ["F"+str(i) for i in range(17, 45)]
    lst2 = ["H"+str(i) for i in range(17, 34)]
    for s in lst1:
        currentCell = sheet[s]
        currentCell.alignment = Alignment(horizontal='right', vertical="center")

    for s in lst2:
        currentCell = sheet[s]
        currentCell.alignment = Alignment(horizontal='right', vertical="center")

import xlrd
import openpyxl

book = xlrd.open_workbook("xls_sample.xls")
sheet = book.sheet_by_index(0)
d28 = sheet.cell_value(rowx=1, colx=14)  # O2
print(d28)  # D28

print(type(book))
lst = []
for rx in range(sheet.nrows):
    for c in range(len(sheet.row(rx))):
        if sheet.row(rx)[c].value != '':
            lst.append(sheet.row(rx)[c].value)


book2 = openpyxl.load_workbook("pbs_mnssd03_r01_2.xlsx", data_only=True)

print(book2)

sheet2 = book2.active
print(sheet2.rows)
lst2 = []
for row in sheet2.iter_rows():
    for cell in row:
        if cell.value is not None:
            lst2.append(cell.value)

print(lst)
print(type(lst2[-1]))

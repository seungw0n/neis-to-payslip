import pprint
# import xlrd
# import excel

__author__ = "Seung Won Joeng - 정승원"
__copyright__ = "Copyright (C) 2022 Seung Won Joeng All rights reserved."
__license__ = "https://github.com/seungw0n/neis-to-payslip/blob/main/LICENSE"


# class NEISPayslip:
#     def __init__(self, NEIS_book):
#         self.info = []
#         if isinstance(NEIS_book, xlrd.book.Book):  # xls format
#             print("NEIS_book - xls format")
#             self.sheet = NEIS_book.sheet_by_index(0)
#             self.info = excel.NEIS_payslip_xls(sheet=self.sheet)
#         else:
#             print("NEIS_book - xlsx format")
#             self.sheet = NEIS_book.active
#             self.info = excel.NEIS_payslip_xlsx(sheet=self.sheet)
#
#         if self.info[1].replace(" ", "") != "급여명세서":
#             raise ValueError
#
#         for i in range(len(self.info)):
#             value = self.info[i]
#             if isinstance(value, str):
#                 if "성명" in value:
#                     self.employee_name = value[-3:]
#                 elif value == "기관명":
#                     self.school_name = self.info[i+1]
#                 elif value == "급여직종\n구분":
#                     self.position = self.info[i+1]
#                 elif value == "공제내역":
#                     self.start_index = i+1
#
#         self.pay = dict()
#
#     def match(self):
#         for i in range(self.start_index, len(self.info)):
#             value = self.info[i]
#             if isinstance(value, str):
#                 self.pay[value] = self.info[i+1]
#
#     def print_pay(self):
#         pprint.pprint(self.pay)
#
#
# ### Test for xls
# book = xlrd.open_workbook("xls_sample.xls")
# neis_obj = NEISPayslip(book)
# neis_obj.match()
# neis_obj.print_pay()
#
#
# ### Test for xlsx
# book = excel.open_excel("pbs_mnssd03_r01_2.xlsx")
# neis_obj = NEISPayslip(book)
# neis_obj.match()
# neis_obj.print_pay()
# print(len(neis_obj.pay))


class NeisPayslip:
    def __init__(self, neis_wb):
        self.sheet = neis_wb['Sheet1']
        if self.sheet['B4'].value.replace(" ", "") != "급여명세서":
            print(self.sheet['B4'].value)
            raise ValueError
        self.employee_name = self.sheet['B6'].value
        self.employee_name = self.employee_name[-3:]  # 이름
        self.school_name = self.sheet['B8'].value
        self.school_name = self.school_name.split("]")[0][1:]  # 학교 이름
        self.position = self.sheet['J9'].value  # 직종
        self.net_pay = 0
        """
        There are total three sections in the NIES excel
        1. 급여내역
        2. 세금내역
        3. 공제내역
        """
        self.pay = dict()
        self.tax = dict()
        self.deduction = dict()

    def get_pay(self):
        multiple_cells = self.sheet['B17': 'B50']
        for row in multiple_cells:
            for cell in row:
                if cell.value == "실수령액":  # Since it has diff column index and it means no more value will be appeared
                    if self.sheet.cell(row=cell.row, column=8).value:
                        self.net_pay = self.sheet.cell(row=cell.row, column=8).value
                    else:
                        self.net_pay = 0
                    pprint.pprint(self.pay)
                    return
                if cell.value == "급여총액":
                    if self.sheet.cell(row=cell.row, column=4).value:
                        self.pay["급여총액"] = self.sheet.cell(row=cell.row, column=4).value
                    else:
                        self.pay['급여총액'] = 0
                elif cell.value:
                    self.pay[cell.value] = self.sheet.cell(row=cell.row, column=4).value

    def get_tax(self):
        multiple_cells = self.sheet['H17': 'H50']
        for row in multiple_cells:
            for cell in row:
                if cell.value == "세금총액":  # Must be end here
                    if self.sheet.cell(row=cell.row, column=9).value:
                        self.tax['세금총액'] = self.sheet.cell(row=cell.row, column=9).value
                    else:
                        self.tax['세금총액'] = 0
                    pprint.pprint(self.tax)
                    return
                elif cell.value:  # Skip none
                    self.tax[cell.value] = self.sheet.cell(row=cell.row, column=9).value

    def get_deduction(self):
        multiple_cells = self.sheet['K17':'K50']  # P = 16
        for row in multiple_cells:
            for cell in row:
                if cell.value == "공제총액":
                    if self.sheet.cell(row=cell.row, column=16).value:
                        self.deduction['공제총액'] = self.sheet.cell(row=cell.row, column=16).value
                    else:
                        self.deduction['공제총액'] = 0
                    pprint.pprint(self.deduction)
                    return
                elif cell.value:
                    self.deduction[cell.value] = self.sheet.cell(row=cell.row, column=16).value

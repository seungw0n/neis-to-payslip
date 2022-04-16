import pprint

__author__ = "Seung Won Joeng - 정승원"
__copyright__ = "Copyright (C) 2022 Seung Won Joeng All rights reserved."
__license__ = "https://github.com/seungw0n/neis-to-payslip/blob/main/LICENSE"
__version__ = "1.0"


class NeisPayslip:
    def __init__(self, neis_wb):
        self.sheet = neis_wb['Sheet1']
        self.employee_name = self.sheet['B6'].value
        self.employee_name = self.employee_name[-3:]  # 이름
        self.school_name = self.sheet['B8'].value
        self.school_name = self.school_name.split("]")[0][1:]  # 학교 이름
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
                    self.net_pay = self.sheet.cell(row=cell.row, column=8).value
                    pprint.pprint(self.pay)
                    return
                elif cell.value:
                    self.pay[cell.value] = self.sheet.cell(row=cell.row, column=4).value

    def get_tax(self):
        multiple_cells = self.sheet['H17': 'H50']
        for row in multiple_cells:
            for cell in row:
                if cell.value == "세금총액":  # Must be end here
                    self.tax['세금총액'] = self.sheet.cell(row=cell.row, column=9).value
                    pprint.pprint(self.tax)
                    return
                elif cell.value:  # Skip none
                    self.tax[cell.value] = self.sheet.cell(row=cell.row, column=9).value

    def get_deduction(self):
        multiple_cells = self.sheet['K17':'K50']  # P = 16
        for row in multiple_cells:
            for cell in row:
                if cell.value == "공제총액":
                    self.deduction['공제총액'] = self.sheet.cell(row=cell.row, column=16).value
                    pprint.pprint(self.deduction)
                    return
                elif cell.value:
                    self.deduction[cell.value] = self.sheet.cell(row=cell.row, column=16).value



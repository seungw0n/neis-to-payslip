import pprint
import xlrd
from src import excel

__author__ = "Seung Won Joeng - 정승원"
__copyright__ = "Copyright (C) 2022 Seung Won Joeng All rights reserved."
__license__ = "https://github.com/seungw0n/neis-to-payslip/blob/main/LICENSE"


def calculate_children_fee(NEIS_data):
    if '가족수당(자녀)' in NEIS_data:
        pay = NEIS_data['가족수당(자녀)']
        if pay == 0:
            return {'first': 0, 'second': 0, 'others': 0}
        if pay == 20000:  # first only
            return {'first': 20000, 'second': 0, 'others': 0}
        elif pay == 60000:  # second only
            return {'first': 0, 'second': 60000, 'others': 0}
        elif pay == 80000:  # first + second only
            return {'first': 20000, 'second': 60000, 'others': 0}
        elif (pay - 80000) % 100000 == 0:  # first + second + others
            return {'first': 20000, 'second': 60000, 'others': 100000 * int((pay - 80000) / 100000)}
        elif (pay - 60000) % 100000 == 0:  # second + others
            return {'first': 0, 'second': 60000, 'others': 100000 * int((pay - 60000) / 100000)}
        elif pay % 100000 == 0:
            return {'first': 0, 'second': 0, 'others': 100000 * int(pay / 100000)}
        else:
            raise ValueError
    return None


class NEISPayslip:
    def __init__(self, NEIS_book):
        self.children_pay = dict()
        self.info = []
        if isinstance(NEIS_book, xlrd.book.Book):  # xls format
            print("NEIS_book - xls format")
            self.sheet = NEIS_book.sheet_by_index(0)
            self.info = excel.NEIS_payslip_xls(sheet=self.sheet)
        else:
            print("NEIS_book - xlsx format")
            self.sheet = NEIS_book.active
            self.info = excel.NEIS_payslip_xlsx(sheet=self.sheet)

        if self.info[1].replace(" ", "") != "급여명세서":
            raise ValueError

        for i in range(len(self.info)):
            value = self.info[i]
            if isinstance(value, str):
                if "성명" in value:
                    self.employee_name = value[-3:]
                    try:
                        self.year = int(value.split(" ")[1])
                        self.month = int(value.split(" ")[3])
                    except ValueError as e:
                        raise ValueError
                elif value == "기관명":
                    self.school_name = self.info[i+1]
                elif value == "급여직종\n구분":
                    self.position = self.info[i+1]
                elif value == "공제내역":
                    self.start_index = i+1

        if self.employee_name is None or self.school_name is None or self.position is None:
            raise ValueError

        self.pay = dict()

    def match(self):
        for i in range(self.start_index, len(self.info)):
            value = self.info[i]
            if isinstance(value, str):
                value = value.replace(" ", "")
                value = value.replace("\n", "")
                if isinstance(self.info[i+1], int) or isinstance(self.info[i+1], float):
                    self.pay[value] = self.info[i+1]

        self.children_pay = calculate_children_fee(self.pay)

    def print_pay(self):
        pprint.pprint(self.pay)
        pprint.pprint(self.children_pay)
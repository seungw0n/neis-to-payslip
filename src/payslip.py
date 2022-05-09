from src.neis import NEISPayslip
from src.form import Form
from openpyxl.styles import Font, colors, Color

import datetime

__author__ = "Seung Won Joeng - 정승원"
__copyright__ = "Copyright (C) 2021 Seung Won Joeng All rights reserved."
__license__ = "https://github.com/seungw0n/neis-to-payslip/blob/main/LICENSE"


def read_only_str(value: str):
    if isinstance(value, str):
        """ 수당명 contains only space, new line. """
        result = value.replace(" ", "")
        result = result.replace("\n", "")
        return result
    else:
        return None


def pay_date(month, year):
    days = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN']
    pay_day = days[datetime.date(year, month, 17).weekday()]
    if pay_day == 'SAT':
        return 16
    elif pay_day == 'SUN':
        return 15
    return 17


def num_date(month, year):
    if month != 12:
        day = datetime.date(year, month + 1, 1)
    else:
        day = datetime.date(year + 1, 1, 1)

    day = day - datetime.timedelta(days=1)
    return day.day


def top_field(payslip_book, obj1: NEISPayslip, obj2: Form):
    print("Function: payslip.top_field starts")
    try:
        sheet = payslip_book.active
        date = pay_date(obj1.month, obj1.year)

        sheet['B3'] = str(obj1.year) + "년 " + str(obj1.month) + "월 임금 명세서"
        sheet['B3'].font = Font(color="000000", size=20)
        sheet['B5'] = "소속 : " + obj1.school_name
        sheet['B5'].font = Font(color="000000", size=14)

        sheet['G5'] = "지급일 : " + str(obj1.year) + "." + str(obj1.month) + "." + str(date) + "."

        sheet['D6'] = obj2.info['성명']
        sheet['F6'] = obj2.info['생년월일']
        sheet['H6'] = obj2.info['직종명']
        sheet['D7'] = obj2.info['계약기간']
        sheet['F7'] = obj2.info['근무형태']
        sheet['H7'] = obj2.info['주 소정근로시간']
        sheet['D8'] = str(num_date(obj1.month, obj1.year))
        sheet['D9'] = obj2.info['신분변동제외한일수']
        sheet['F8'] = obj2.info['월 소정근로시간']
        sheet['G9'] = obj2.info['직계존속+첫째자녀수']
        sheet['H9'] = obj2.info['셋째 자녀 이후수']
        sheet['D10'] = obj2.info['연장근로시간']
        sheet['D11'] = obj2.info['법내초과근로시간']
        sheet['F10'] = obj2.info['휴일근로시간']
        sheet['F11'] = obj2.info['휴일연장']
        sheet['H10'] = obj2.info['야간근로시간']
        sheet['D12'] = obj2.info['결근 등 무급일']
        sheet['D13'] = obj2.info['결근 등 무급시간']
        sheet['F12'] = obj2.info['근속수당 경력년수']
        sheet['H12'] = obj2.info['통상임금']

    except Exception as e:
        print(e)


def bottom_field(payslip_book, obj1: NEISPayslip, obj2: Form):
    print("Function: payslip.bottom_field starts")
    try:
        sheet = payslip_book.active
        rows = sheet['B17': 'I53']
        for r in rows:
            for cell in r:
                if cell.value is not None and isinstance(cell.value, str):
                    if read_only_str(cell.value) in obj1.pay:
                        coordinate = cell.coordinate
                        if 'C' in coordinate:
                            coordinate = coordinate.replace('C', 'F')
                        elif 'G' in coordinate:
                            coordinate = coordinate.replace('G', 'H')
                        elif read_only_str(cell.value) == "급여총액" or read_only_str(cell.value) == "실수령액":
                            coordinate = coordinate.replace('B', 'F')

                        sheet[coordinate] = obj1.pay[read_only_str(cell.value)]

                    elif read_only_str(cell.value) == '공제(세금)총액':
                        coordinate = cell.coordinate
                        coordinate = coordinate.replace('G', 'H')

                        if '공제총액' in obj1.pay and '세금총액' in obj1.pay:
                            sheet[coordinate] = obj1.pay['공제총액'] + obj1.pay['세금총액']
                        elif '공제총액' in obj1.pay:
                            sheet[coordinate] = obj1.pay['공제총액']
                        elif '세금총액' in obj1.pay:
                            sheet[coordinate] = obj1.pay['세금총액']
                        else:
                            sheet[coordinate] = 0

                    elif read_only_str(cell.value) == '가족수당(직계존속,첫째자녀)':
                        if obj1.children_pay is not None:
                            coordinate = cell.coordinate
                            coordinate = coordinate.replace('C', 'F')

                            sheet[coordinate] = obj1.children_pay['first']

                    elif read_only_str(cell.value) == '가족수당(둘째자녀)':
                        if obj1.children_pay is not None:
                            coordinate = cell.coordinate
                            coordinate = coordinate.replace('C', 'F')

                            sheet[coordinate] = obj1.children_pay['second']

                    elif read_only_str(cell.value) == '가족수당(셋째자녀이후)':
                        if obj1.children_pay is not None:
                            coordinate = cell.coordinate
                            coordinate = coordinate.replace('C', 'F')

                            sheet[coordinate] = obj1.children_pay['others']

                    elif read_only_str(cell.value) in obj2.overpay:
                        coordinate = cell.coordinate
                        coordinate = coordinate.replace('C', 'F')

                        sheet[coordinate] = obj2.overpay[read_only_str(cell.value)]

                    elif "(인)" in read_only_str(cell.value):
                        coordinate = cell.coordinate
                        sheet[
                            coordinate] = "※ 위 근로자 " + obj1.employee_name + "은(는) 임금명세서 1부를 교부받았습니다.       성명    " + obj1.employee_name + "      (인)                    " + \
                                          str(obj1.year) + "." + str(obj1.month) + "." + str(
                            pay_date(obj1.month, obj1.year)) + "."

    except Exception as e:
        print(e)

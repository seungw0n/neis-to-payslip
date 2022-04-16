import datetime
from src.excel import open_excel, save_excel, shift_right
from src.neis import NeisPayslip
from src.payslip import SchoolPayslip
from src.form import Form
from openpyxl.styles import Font, colors, Color

__author__ = "Seung Won Joeng - 정승원"
__copyright__ = "Copyright (C) 2022 Seung Won Joeng All rights reserved."
__license__ = "https://github.com/seungw0n/neis-to-payslip/blob/main/LICENSE"
__version__ = "1.0"

year = datetime.datetime.now().year
month = datetime.datetime.now().month


def get_date():
    days = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN']
    pay_day = days[datetime.date(year, month, 17).weekday()]
    if pay_day == 'SAT':
        return 16
    elif pay_day == 'SUN':
        return 15
    return 17


def num_date():
    if month != 12:
        day = datetime.date(year, month + 1, 1)
    else:
        day = datetime.date(year + 1, 1, 1)

    day = day - datetime.timedelta(days=1)
    return day.day


# def match_names(neis_name, input_name):
#     if neis_name != input_name:
#         raise ValueError
#

def top_field(sheet, neis_obj: NeisPayslip, form_obj: Form):
    """ 엑셀 상단부에 위치한 셀을 위한 함수 (날짜, 이름, 생년월일 등) """
    """
        sheet: output 될 sheet
        filename: 양식으로 준 파일
    """
    print("Function fill.py-top_field starts here")
    name = neis_obj.employee_name
    data = form_obj.info

    pay_date = get_date()

    sheet['B3'] = str(year) + "년 " + str(month) + "월 임금 명세서"
    sheet['B3'].font = Font(color="000000", size=20)
    sheet['B5'] = "소속 : " + neis_obj.school_name
    sheet['B5'].font = Font(color="000000", size=14)

    sheet['G5'] = "지급일 : " + str(year) + "." + str(month) + "." + str(pay_date) + "."

    sheet['B48'] = "※ 위 근로자 " + name + "은(는) 임금명세서 1부를 교부받았습니다.       성명    " + name + "      (인)                    " + \
                   str(year) + "." + str(month) + "." + str(pay_date) + "."

    sheet['D6'] = data['성명']
    sheet['F6'] = data['생년월일']
    sheet['H6'] = data['직종명']
    sheet['D7'] = data['계약기간']
    sheet['F7'] = data['근무형태']
    sheet['H7'] = data['주 소정근로시간']
    sheet['D8'] = str(num_date())
    sheet['D9'] = data['신분변동제외한일수']
    sheet['F8'] = data['월 소정근로시간']
    sheet['G9'] = data['직계존속+첫째자녀수']
    sheet['H9'] = data['셋째 자녀 이후수']
    sheet['D10'] = data['연장근로시간']
    sheet['D11'] = data['법내초과근로시간']
    sheet['F10'] = data['휴일근로시간']
    sheet['F11'] = data['휴일연장']
    sheet['H10'] = data['야간근로시간']
    sheet['D12'] = data['결근 등 무급일']
    sheet['D13'] = data['결근 등 무급시간']
    sheet['F12'] = data['근속수당 경력년수']
    sheet['H12'] = data['통상임금']


# Todo: Cell 채워넣기
def monthly_field(sheet, school_obj: SchoolPayslip):  # range: C17 - C30, D17 - D30, F17 - F30
    """ 매월 지급 셀을 채워넣기 위한 함수 """  # title: C / description: D / amount: F
    print("Function fill.py-monthly_field starts here")
    multiple_cells = sheet['C17': 'C30']
    for row in multiple_cells:
        for cell in row:
            if cell.value in school_obj.monthly_pay:
                sheet['F' + str(cell.row)] = school_obj.monthly_pay[cell.value]
            else:
                print("[monthly_field - 매월지급] Key 값:" + cell.value + " 가 학교 임금명세서에 존재하지 않습니다.")


def non_periodically_field(sheet, school_obj: SchoolPayslip):  # 31 - 38
    """ 부정기 지급 셀을 채워넣기 위한 함수 """
    print("Function fill.py-non_periodically_field starts here")
    multiple_cells = sheet['C31': 'C38']  # 부정기 지급
    for row in multiple_cells:
        for cell in row:
            if cell.value in school_obj.non_periodical_pay:
                sheet['F' + str(cell.row)] = school_obj.non_periodical_pay[cell.value]
            else:
                print("[non_periodically_field - 부정기지급] Key 값:" + cell.value + " 가 학교 임금명세서에 존재하지 않습니다.")


def rest_field(sheet, school_obj: SchoolPayslip):  # 39 - 44
    """ 기타금품 셀을 채워넣기 위한 함수 """
    print("Function fill.py-rest_field starts here")
    multiple_cells = sheet['C39': 'C44']  # 기타금품
    for row in multiple_cells:
        for cell in row:
            if cell.value in school_obj.rest_pay:
                sheet['F' + str(cell.row)] = school_obj.rest_pay[cell.value]
            else:
                print("[rest_pay - 기타금품] Key 값:" + cell.value + " 가 학교 임금명세서에 존재하지 않습니다.")


def deduction_field(sheet, school_obj: SchoolPayslip):  # G, H 17 - 33
    """ 공제(세금) 내역 셀을 채워넣기 위한 함수 """
    print("Function fill.py-deduction_field starts here")
    multiple_cells = sheet['G17': 'G33']  # 공제(세금)
    for row in multiple_cells:
        for cell in row:
            if cell.value in school_obj.deduction:
                sheet['H' + str(cell.row)] = school_obj.deduction[cell.value]
            else:
                print("[deduction - 공제(세금)] Key 값:" + cell.value + " 가 학교 임금명세서에 존재하지 않습니다.")


def total_field(sheet, school_obj: SchoolPayslip):
    print("Function fill.py-total_field starts here")
    total_pay = school_obj.total['급여 총액']
    total_deduction = school_obj.total['공제(세금) 총액']
    net_pay = school_obj.total['실수령액']

    sheet['F45'] = total_pay
    sheet['H45'] = total_deduction
    sheet['F46'] = net_pay


def fill(sheet, neis_obj: NeisPayslip, form_obj: Form, school_obj: SchoolPayslip):
    print("Function fill.py-fill starts here")
    top_field(sheet, neis_obj, form_obj)
    monthly_field(sheet, school_obj)
    non_periodically_field(sheet, school_obj)
    rest_field(sheet, school_obj)
    deduction_field(sheet, school_obj)
    total_field(sheet, school_obj)


# if __name__ == "__main__":
#     """ Testing for this module """
#     form_filenames = ["../temp_april/gojan/.xlsx", "../temp_april/gojan/.xlsx", "../temp_april/gojan/.xlsx",
#                       "../temp_april/gojan/.xlsx", "../temp_april/gojan/.xlsx", "../temp_april/gojan/.xlsx",
#                       "../temp_april/gojan/.xlsx", "../temp_april/gojan/.xlsx", "../temp_april/gojan/.xlsx",
#                       "../temp_april/gojan/.xlsx", "../temp_april/gojan/.xlsx", "../temp_april/gojan/.xlsx"]
#     # form_filename = "../.xlsx"
#     neis_filenames = ["../temp_april/neis/pbs_mnssd03_r01_2.xlsx", "../temp_april/neis/pbs_mnssd03_r01_3.xlsx",
#                       "../temp_april/neis/pbs_mnssd03_r01_4.xlsx", "../temp_april/neis/pbs_mnssd03_r01_5.xlsx",
#                       "../temp_april/neis/pbs_mnssd03_r01_6.xlsx", "../temp_april/neis/pbs_mnssd03_r01_7.xlsx",
#                       "../temp_april/neis/pbs_mnssd03_r01_8.xlsx", "../temp_april/neis/pbs_mnssd03_r01_9.xlsx",
#                       "../temp_april/neis/pbs_mnssd03_r01_10.xlsx", "../temp_april/neis/pbs_mnssd03_r01_11.xlsx",
#                       "../temp_april/neis/pbs_mnssd03_r01_12.xlsx", "../temp_april/neis/pbs_mnssd03_r01_13.xlsx",
#                       "../temp_april/neis/pbs_mnssd03_r01_14.xlsx", "../temp_april/neis/pbs_mnssd03_r01_15.xlsx",
#                       ]
#
#
#     sample_filename = "../표준임금명세서_양식.xlsx"
#     for neis_filename in neis_filenames:
#         neis_wb = open_excel(neis_filename)
#         neis_obj = NeisPayslip(neis_wb)
#         neis_obj.get_pay()
#         neis_obj.get_tax()
#         neis_obj.get_deduction()
#
#         form_filename = "../temp_april/gojan/" + neis_obj.employee_name + ".xlsx"
#         try:
#             form_wb = open_excel(form_filename)
#         except FileNotFoundError:
#             continue
#
#         form_obj = Form(form_wb)
#         form_obj.get_info()
#
#         school_obj = SchoolPayslip(neis_obj, form_obj)
#         school_obj.match_pay()
#         school_obj.match_tax()
#         school_obj.match_deduction()
#         school_obj.match_total()
#         print(school_obj.monthly_pay)
#         print(school_obj.non_periodical_pay)
#         print(school_obj.rest_pay)
#         print(school_obj.deduction)
#         print(school_obj.total)
#
#         sample_wb = open_excel(sample_filename)
#         sample_sheet = sample_wb['Sheet1']
#
#         top_field(sample_sheet, neis_obj, form_obj)
#         monthly_field(sample_sheet, school_obj)
#         non_periodically_field(sample_sheet, school_obj)
#         rest_field(sample_sheet, school_obj)
#         deduction_field(sample_sheet, school_obj)
#         total_field(sample_sheet, school_obj)
#
#         # shift_right(sample_sheet)
#         save_excel(sample_wb, "임금명세서_" + neis_obj.employee_name + ".xlsx")
"""
    # 1. neis 읽어오기
    neis_wb = open_excel(neis_filename)
    neis_obj = NeisPayslip(neis_wb)
    neis_obj.get_pay()
    neis_obj.get_tax()
    neis_obj.get_deduction()

    # 2. 양식 파일 읽어오기
    form_wb = open_excel(form_filename)
    form_obj = Form(form_wb)
    form_obj.get_info()

    # 3. neis 에서 받아온 금액들 + 초과근무수당 schoolpayslip 객체로 옮겨 놓기
    school_obj = SchoolPayslip(neis_obj, form_obj)
    school_obj.match_pay()
    school_obj.match_tax()
    school_obj.match_deduction()
    school_obj.match_total()
    print(school_obj.monthly_pay)
    print(school_obj.non_periodical_pay)
    print(school_obj.rest_pay)
    print(school_obj.deduction)
    print(school_obj.total)

    # 3. 표준임금명세서 양식 불러오기
    sample_wb = open_excel(sample_filename)
    sample_sheet = sample_wb['Sheet1']

    # 3a) 표준임금명세서 채워넣기
    top_field(sample_sheet, neis_obj, form_obj)
    monthly_field(sample_sheet, school_obj)
    non_periodically_field(sample_sheet, school_obj)
    rest_field(sample_sheet, school_obj)
    deduction_field(sample_sheet, school_obj)
    total_field(sample_sheet, school_obj)

    # 4. 저장하기
    save_excel(sample_wb, "임금명세서_" + neis_obj.employee_name + ".xlsx")
"""

import math

__author__ = "Seung Won Joeng - 정승원"
__copyright__ = "Copyright (C) 2022 Seung Won Joeng All rights reserved."
__license__ = "https://github.com/seungw0n/neis-to-payslip/blob/main/LICENSE"


def round_down(number):
    number = number / 10
    number = math.trunc(number)
    return number * 10


def get_time(time):
    if time == 0 or time == "0" or time == "0시간":
        return 0, 0

    time = time.replace(" ", "")
    r_hour = ""
    r_min = ""
    temp = ""

    for t in time:
        if '0' <= t <= '9':
            temp = temp + t
        else:
            if t == "시":
                r_hour = r_hour + temp
            elif t == "분":
                r_min = r_min + temp
            temp = ""

    if r_hour == "":
        r_hour = "0"
    if r_min == "":
        r_min = "0"

    return int(r_hour), int(r_min)


class Form:
    """ 사용자가 제공할 기본 양식 파일 (명세서 상단부 위치한 테이블) eg. 기초급여정보.xlsx """

    def __init__(self, wb):
        self.sheet = wb['Sheet1']
        if self.sheet['B24'].value != "key=form":
            raise KeyError

        self.info = dict()
        self.overpay = dict()

    def get_info(self):
        self.info['성명'] = self.sheet['C3'].value  # 성명
        self.info['생년월일'] = self.sheet['E3'].value  # 생년월일
        self.info['직종명'] = self.sheet['G3'].value  # 직종명
        self.info['계약기간'] = self.sheet['C4'].value  # 계약기간
        self.info['근무형태'] = self.sheet['E4'].value  # 근무형태
        self.info['주 소정근로시간'] = self.sheet['G4'].value  # 주 소정근로시간
        # Skip 당월급여계산일수 - handled by write.intro_field
        self.info['신분변동제외한일수'] = self.sheet['C6'].value
        self.info['월 소정근로시간'] = self.sheet['E5'].value  # 월 소정근로시간
        self.info['직계존속+첫째자녀수'] = self.sheet['F6'].value  # 직계존속 + 첫째자녀수
        self.info['셋째 자녀 이후수'] = self.sheet['G6'].value  # 셋째 자녀 이후수
        self.info['연장근로시간'] = self.sheet['C7'].value
        self.info['법내초과근로시간'] = self.sheet['C8'].value
        self.info['휴일근로시간'] = self.sheet['E7'].value
        self.info['휴일연장'] = self.sheet['E8'].value
        self.info['야간근로시간'] = self.sheet['G7'].value
        self.info['결근 등 무급일'] = self.sheet['C9'].value
        self.info['결근 등 무급시간'] = self.sheet['C10'].value
        self.info['근속수당 경력년수'] = self.sheet['E9'].value
        self.info['통상임금'] = self.sheet['G9'].value

        try:
            # 연장근로수당 계산법
            h, m = get_time(self.info['연장근로시간'])
            f1 = int(h * 1.5 * self.info['통상임금'] + m * 1.5 * self.info['통상임금'] / 60)
            h, m = get_time(self.info['법내초과근로시간'])
            f2 = int(h * self.info['통상임금'] + m * self.info['통상임금'] / 60)
            self.overpay['연장근로수당'] = round_down(f1 + f2)

            # 휴일근로수당 계산법
            h, m = get_time(self.info['휴일근로시간'])
            f1 = int(h * 1.5 * self.info['통상임금'] + m * 1.5 * self.info['통상임금'] / 60)
            h, m = get_time(self.info['휴일연장'])
            f2 = int(h * 2 * self.info['통상임금'] + m * self.info['통상임금'] / 60)
            self.overpay['휴일근로수당'] = round_down(f1 + f2)

            # 야간근로수당 계산법
            h, m = get_time(self.info['야간근로시간'])
            f1 = int(h * 0.5 * self.info['통상임금'] + m * 0.5 * self.info['통상임금'] / 60)
            self.overpay['야간근로수당'] = round_down(f1)

        except Exception as e:
            raise e

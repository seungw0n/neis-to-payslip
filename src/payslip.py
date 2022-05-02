from src.neis import NeisPayslip
from src.form import Form
import math

__author__ = "Seung Won Joeng - 정승원"
__copyright__ = "Copyright (C) 2021 Seung Won Joeng All rights reserved."
__license__ = "https://github.com/seungw0n/neis-to-payslip/blob/main/LICENSE"


def round_down(number):
    number = number / 10
    number = math.trunc(number)
    return number * 10


def get_time(time):
    if time == 0 or time == "0" or time == "0시간":
        return 0, 0

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


class SchoolPayslip:
    def __init__(self, neis_obj: NeisPayslip, form_obj: Form):
        self.neis_obj = neis_obj
        self.form_obj = form_obj
        if neis_obj.employee_name != form_obj.info['성명']:
            raise ValueError

        self.monthly_pay = {'기본급': 0, '정액급식비': 0, '근속수당': 0, '위험근무수당': 0, '영양사면허가산수당': 0, '정근수당가산금': 0, '정근수당추가가산금': 0,
                            '관리업무수당(일반직)': 0, '직급보조비': 0, '면허수당': 0, '자격수당': 0, '직무수당': 0, '특수업무수당': 0, '특수지원수당': 0}
        self.non_periodical_pay = {'연장근로수당': 0, '휴일근로수당': 0, '야간근로수당': 0, '명절휴가비': 0, '정기상여금': 0, '성과상여금': 0, '정근수당': 0,
                                   '연차유급휴가미사용수당': 0}
        self.rest_pay = {'가족수당(배우자)': 0, '가족수당\n(직계존속, 첫째자녀)': 0, '가족수당\n(둘째자녀)': 0, '가족수당\n(셋째자녀 이후)': 0,
                         '가족수당\n(부양가족)': 0, '기타수당': 0}

        self.deduction = {'소득세': 0, '지방소득세': 0, '건강보험': 0, '건강보험연말정산': 0, '노인장기요양보험': 0, '장기요양연말정산': 0, '국민연금': 0,
                          '고용보험': 0, '고용보험연말정산': 0, '교직원공제회비': 0, '친목회비': 0, '식대': 0, '연말정산소득세': 0, '연말정산지방소득세': 0,
                          '중도정산소득세': 0, '중도정산지방소득세': 0, '기타공제': 0}

        self.children = {'first': 0, 'second': 0, 'others': 0}

        self.total = {'급여 총액': 0, '공제(세금) 총액': 0, '실수령액': 0}

    def set_children(self, total_pay):
        if total_pay == 0:
            return
        if total_pay == 20000:  # first only
            self.children['first'] = 1
        elif total_pay == 60000:  # second only
            self.children['second'] = 1
        elif total_pay == 80000:  # first + second only
            self.children['first'] = 1
            self.children['second'] = 1
        elif (total_pay - 80000) % 100000 == 0:  # first + second + others
            self.children['first'] = 1
            self.children['second'] = 1
            self.children['others'] = int((total_pay - 80000) / 100000)
        elif (total_pay - 60000) % 100000 == 0:  # second + others
            self.children['second'] = 1
            self.children['others'] = int((total_pay - 60000) / 100000)
        elif total_pay % 100000 == 0:
            self.children['others'] = int(total_pay / 100000)
        else:
            print("Check function set_children")

    def match_pay(self):
        # Todo: 야간근로, 연장근로, 법내초과근로, 휴일근로 핸들링하기
        print("Function match_pay starts here")
        for k, v in self.neis_obj.pay.items():
            # print("Key: " + str(k) + "\tValue: " + str(v))
            if k in self.monthly_pay:
                self.monthly_pay[k] = v
                print("\t" + k + ": 매월지급")

            elif k in self.non_periodical_pay:
                self.non_periodical_pay[k] = v
                print("\t" + k + ": 부정기지급")

            elif k in self.rest_pay:
                self.rest_pay[k] = v
                print("\t" + k + ": 기타금품")

            elif k == "가족수당(자녀)":
                self.set_children(self.neis_obj.pay[k])
                if self.children['first'] == 1:
                    self.rest_pay['가족수당\n(직계존속, 첫째자녀)'] = 20000
                if self.children['second'] == 1:
                    self.rest_pay['가족수당\n(둘째자녀)'] = 60000
                if self.children['others'] > 0:
                    self.rest_pay['가족수당\n(셋째자녀 이후)'] = 100000 * self.children['others']
                print("\t자녀 수" + str(self.children))

            elif k == "시간외근무수당(초과분)":  # 야간근로, 연장근로, 법내초과근로, 휴일근로 핸들링
                rate = self.form_obj.info['통상임금']
                overtimeOne = self.form_obj.info['연장근로시간']
                overtimeTwo = self.form_obj.info['법내초과근로시간']
                holidayOne = self.form_obj.info['휴일근로시간']
                holidayTwo = self.form_obj.info['휴일연장']
                nightShift = self.form_obj.info['야간근로시간']

                # 연장근로수당 계산법
                h, m = get_time(overtimeOne)
                f1 = int(h * 1.5 * rate + m * 1.5 * rate / 60)
                h, m = get_time(overtimeTwo)
                f2 = int(h * rate + m * rate / 60)
                self.non_periodical_pay['연장근로수당'] = round_down(f1 + f2)


                # 휴일근로수당 계산법
                h, m = get_time(holidayOne)
                f1 = int(h * 1.5 * rate + m * 1.5 * rate / 60)
                h, m = get_time(holidayTwo)
                f2 = int(h * 2 * rate + m * 2 * rate / 60)
                self.non_periodical_pay['휴일근로수당'] = round_down(f1 + f2)

                # 야간근로수당 계산법
                h, m = get_time(nightShift)
                f1 = int(h * 0.5 * rate + m * 0.5 * rate / 60)
                self.non_periodical_pay['야간근로수당'] = round_down(f1)

            else:
                if k == "급여총액":
                    continue
                print("[match - pay] Key 값:" + k + " 가 학교 임금명세서에 존재하지 않습니다.")

    def match_tax(self):
        print("Function match_tax starts here")
        for k, v in self.neis_obj.tax.items():
            if k in self.deduction:
                self.deduction[k] = v
            else:
                if k == "세금총액":
                    continue
                print("[match - tax] Key 값:" + k + " 가 학교 임금명세서에 존재하지 않습니다.")

    def match_deduction(self):
        print("Function match_deduction starts here")
        for k, v in self.neis_obj.deduction.items():
            if k in self.deduction:
                self.deduction[k] = v
            else:
                if k == "공제총액":
                    continue
                print("[match - deduction] Key 값:" + k + " 가 학교 임금명세서에 존재하지 않습니다.")

    def match_total(self):
        print("Function match_total starts here")
        try:
            self.total['급여 총액'] = self.neis_obj.pay['급여총액']
            self.total['공제(세금) 총액'] = self.neis_obj.tax['세금총액'] + self.neis_obj.deduction['공제총액']
            self.total['실수령액'] = self.neis_obj.net_pay
        except Exception as e:
            print("match_total" + str(e))

__author__ = "Seung Won Joeng - 정승원"
__copyright__ = "Copyright (C) 2022 Seung Won Joeng All rights reserved."
__license__ = "https://github.com/seungw0n/neis-to-payslip/blob/main/LICENSE"


class Form:
    """ 사용자가 제공할 기본 양식 파일 (명세서 상단부 위치한 테이블) eg. sample_input.xlsx """
    def __init__(self, wb):
        self.sheet = wb['Sheet1']
        if self.sheet['B24'].value != "key=form":
            raise KeyError

        self.info = dict()

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

## 인천광역시 교육청 교육감 소속 근로자 임금 명세서 프로그램

### Python Version
    - python 3.9

### Installation
    - pip3 install openpyxl
    - pip3 install pyqt5
    - pip3 install pyqt5designer

### Environment
    - Compatiable with Windows

### Initial Screen

![캡처1](https://user-images.githubusercontent.com/34816905/163906633-655e39c0-205b-4544-878b-43f19d66dcea.PNG)


### 프로그램 소개
본 프로그램은 Neis 에서 발급된 급여명세서와 개인별 급여기초정보를 이용하여
교부용 임금명세서 엑셀 파일을 생성합니다.


### 프로그램 사용 방법

    https://drive.google.com/drive/folders/1OkryYFAKQr9k5-WYc6q9GkozuDsPu2i5?usp=sharing
1. 위 구글 드라이브 링크에서 개인별_급여기초정보, 임금명세서.exe, main.ui, payslip.xlsx 을 다운 받아주세요.


2. 다운받으신 개인병 급여기초정보를 옳바르게 작성합니다.

    (1회 최초 작성 후 매 달 재사용이 가능합니다, 단 연장근로시간, 법내초과근로시간, 휴일근로(연장)시간, 야간근로시간등 수정해야할 부분들은 매월 수정해야 합니다.)
    

    - 양식 내의 시간 작성법은 다음과 같습니다. (띄어쓰기x)
    - 초과시간이 없을 경우: 0 혹은 0시간
    - 초과시간이 2시간 35분일 경우: 2시간35분
    - 초과시간이 1시간일 경우: 1시간


3. Neis 에서 급여 명세서를 저장합니다. (어떠한 편집도 하시면 안됩니다.)\



4. 실행파일 임금명세서.exe 을 실행 시켜줍니다.


5. Neis 급여명세서 버튼을 클릭 후 다운 받은 해당 근로자의 Neis 급여명세서를 선택해줍니다.


6. 엑셀 양식 버튼을 클릭 후 작성하신 해당 근로자의 급여기초정보 엑셀 파일을 선택해줍니다.


7. 임금명세서 버튼을 클릭하시면, 실행프로그램이 있는 곳에 교부용 임금명세서 엑셀 파일이 생깁니다.


*보다 더 자세한 설명은 구글 드라이브 내에 임금명세서 자동화 프로그램 사용자 문서를 참조하시기바랍니다.*

### Errors
#### 2022-04-12
- The color of text: title and name of the organization are automatically set to Red. [Fixed]
#### 2022-04-14
- Packaging error: main.ui [Fixed]
#### 2022-04-19
- Not supported xls format. [Comment: open xls file using xlrd and convert to xlsx]
#### 2022-04-20
- Not supported blank and zero value (pay, tax, and deduction). It causes blank on the cells. [Will be fixed on Apr 21]

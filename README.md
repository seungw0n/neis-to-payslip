## 인천광역시 교육청 교육감 소속 근로자 임금 명세서 프로그램

### 공지사항
#### Ver 1.5: 2022-05-16
#### *NEIS 급여명세서 xls, xlsx 모두 다 됩니다. [2022-05-12]*

#### *수당명 변경 가능합니다.*
- NEIS 급여명세서와 payslip 에 있는 수당명이 다를 경우 변경해주세요.
- 예시) 동봉된 payslip.xlsx 에 기본급에 본봉으로 바꿔주세요.
- 필요없는 수당명일 경우
    -  셀 클릭 -> 삭제 -> 행 전체 -> 확인
- 행 전체 삭제 시 그 행에 있는 공제세금 수당명도 함께 삭제되오니, 삭제 된 공제/세금 내역을 
공제세금내역란에 다시 추가하시기 바랍니다. 

#### *초과 시간 입력시 분으로 입력 가능합니다*
- 2시간 8분, 128분 둘 다 가능합니다. 
    
#### *보안 프로그램에 걸리는 이유는 "코드서명인증서"를 추가 안해서 그렇습니다.. 오해가 생길 수 있어 모든 코드를 오픈한 것입니다.* [자세히 알아보기 - by Tom Kim](https://blog.dramancompany.com/2015/12/%EC%B2%98%EC%9D%8C-windows-%EC%84%A4%EC%B9%98-%ED%8C%8C%EC%9D%BC%EC%9D%84-%EB%B0%B0%ED%8F%AC%ED%95%98%EB%8A%94-%EA%B0%9C%EB%B0%9C%EC%9E%90%EB%93%A4%EC%9D%84-%EC%9C%84%ED%95%98%EC%97%AC/)

### 프로그램 다운로드
---
    Last updated on 2022.05.12

다운로드 (구글 링크 클릭 후 "표준임금명세서_작성_프로그램" 다운로드):
    https://drive.google.com/drive/folders/1OkryYFAKQr9k5-WYc6q9GkozuDsPu2i5?usp=sharing

![캡처1](https://user-images.githubusercontent.com/34816905/163906633-655e39c0-205b-4544-878b-43f19d66dcea.PNG)


### 프로그램 소개 및 실행
---
본 프로그램은 Neis 에서 발급된 급여명세서와 개인별 급여기초정보를 이용하여
교부용 임금명세서 엑셀 파일을 생성합니다.
  
  

1. 위 구글 드라이브 링크에서 개인별_급여기초정보, 임금명세서.exe, main.ui, payslip.xlsx 을 다운 받아주세요.


2. 다운받으신 개인병 급여기초정보를 옳바르게 작성합니다.

    (1회 최초 작성 후 매 달 재사용이 가능합니다, 단 연장근로시간, 법내초과근로시간, 휴일근로(연장)시간, 야간근로시간등 수정해야할 부분들은 매월 수정해야 합니다.)
    

    - 양식 내의 시간 작성법은 다음과 같습니다. (띄어쓰기x)
    - 초과시간이 없을 경우: 0 혹은 0시간
    - 초과시간이 2시간 35분일 경우: 2시간35분
    - 초과시간이 1시간일 경우: 1시간


3. Neis 에서 급여 명세서를 엑셀 형식으로 저장합니다. (어떠한 편집도 하시면 안됩니다.)


4. 실행파일 임금명세서.exe 을 실행 시켜줍니다.


5. Neis 급여명세서 버튼을 클릭 후 다운 받은 해당 근로자의 Neis 급여명세서를 선택해줍니다.


6. 엑셀 양식 버튼을 클릭 후 작성하신 해당 근로자의 급여기초정보 엑셀 파일을 선택해줍니다.


7. 임금명세서 버튼을 클릭하시면, 실행프로그램이 있는 곳에 교부용 임금명세서 엑셀 파일이 생깁니다.


*보다 더 자세한 설명은 구글 드라이브 내에 임금명세서 자동화 프로그램 사용자 문서를 참조하시기바랍니다.*


### Python Version
---
    - python 3.9

### Installation
---
    - pip3 install openpyxl
    - pip3 install pyqt5
    - pip3 install pyqt5designer


### Environment
---
    - Compatiable with Windows



### Released Notes
---
#### 2022-04-12
- The color of text: title and name of the organization are automatically set to Red. [Fixed]
#### 2022-04-14
- Packaging error: main.ui [Fixed]
#### 2022-04-20
- Not supported blank and zero value (pay, tax, and deduction). It causes blank on the cells. [Fixed on April 20]
#### 2022-05-12
- Now it supports both extensions: xls and xlsx. [Fixed on May 5]
- Using xlrd to read xls extension.
#### 2022-05-16
- Add functionality: accept time that contains only minute unit.
- 
### Error
---
- Not supported xls format. [Todo: open xls file using xlrd and convert to xlsx]


### Todo


# 풀어야할 것: 하루에 두 강사가 들어가는 경우 (홍길동/김철수)
# 장기과정은 제외
# 비고에 비고사항 넣기
# 추후 구조화하기

import openpyxl
import pandas as pd
from datetime import datetime
from pprint import pprint

schedule_file = '2018일정계획표(2018.03.23)(최종).xlsx'
instructor_file = 'Instructor.xlsx'

#일정표와 강사진 파일 불러오기
workbook = openpyxl.load_workbook(schedule_file)
dfInstructor = pd.read_excel(instructor_file)

# 강사리스트
allInstructor = []

month_sheets = [sheet_name for sheet_name in workbook.sheetnames if '월' in sheet_name]
sheet_name = month_sheets[4] # sheet: 5월로 지정 (test용)
sheet = workbook[sheet_name]

dayOfWeek = ('월', '화', '수', '목', '금', '토', '일')
columns = ['행번호', '사업구분', '과정명', '강의장(예정)', '강의장(변경)', '개강시간', '주의사항 및 비고', '강사', '날짜']

# DataFrame 생성
df = pd.DataFrame(columns = columns)

# /홍길동?, 홍길동?/, 홍길동? => '홍길동'으로 바꾸기
def GetNameOnly(instructor):
    nameOnly = instructor.replace("/", "")
    nameOnly = nameOnly.replace("?", "")
    return nameOnly

# 주의사항 및 비고에서 timeSequence 문자열만 추출하기
def GetTimeSequence(string):
    startIndex = string.index('(')
    endIndex = string.index(')')
    timeSequence = string[startIndex:endIndex+1]
    return timeSequence 
    
# df의 '강사', '날짜' columns에  강사와 날짜 데이터 append 하기
def GetInstructorAndDate(row):
    instructorList = []
    dateList = []
    for columnNum in range(10, 16):
        rowNum = row
        instructor = sheet[rowNum][columnNum].value # Cell의 강사명
        if instructor is not None:
            instructor = GetNameOnly(instructor) # /홍길동?, 홍길동?/, 홍길동? => '홍길동'으로 바꾸기
            if not instructor in allInstructor: # 중복된 강사명이 allInstructor list에 없으면 추가
                if instructor in dfInstructor['강사명'].values: # 강사명이 아닌 값 (공휴일 등)은 제외시키기
                    allInstructor.append(instructor)
            if instructor in dfInstructor['강사명'].values:
                while True:
                    rowNum -= 1
                    date = sheet[rowNum][columnNum].value
                    if type(date) == datetime:
                        dayNumber = date.strftime('%u')
                        day = date.strftime('%m.%d') + '(' + dayOfWeek[int(dayNumber) - 1] + ')'
                        instructorList.append(instructor)
                        dateList.append(day)
                        break
    return instructorList, dateList

# 대분류를 받아서 데이터 저장용 DataFrame 만들기
def MakeDF(df, section):
    for row in range(4, sheet._current_row + 1):
        if sheet[row][0].value == section:
            instructorList, dateList =  GetInstructorAndDate(row)
            timeSequence = GetTimeSequence(sheet[row][9].value)
            df = df.append({'행번호': row, '사업구분': sheet[row][0].value, '과정명': sheet[row][1].value, '강의장(예정)': sheet[row][2].value, '강의장(변경)': sheet[row][3].value, '개강시간': sheet[row][5].value.strftime('%R'), '주의사항 및 비고': timeSequence, '강사': instructorList, '날짜': dateList}, True)
    return df

# 과정명, 강사, 장소, 교육일정, 비고
df = MakeDF(df, '구매자재')
newColumns = ['강사', '과정명', '강의장', '날짜', '비고']
newdf = pd.DataFrame(columns = newColumns)

# 예정강의장, 변경강의장으로 부터 강의장소 정보 받아오기 (숫자일 경우 서울, 그 이외는 지방)
def SelectClassRoom(plan, change):
    if change is None:
        try:
            plan = int(plan)
            classRoom = '서울'
        except:
            classRoom = plan
    else:
        try:
            change = int(change)
            classRoom = '서울'
        except:
            classRoom = change
            
    return classRoom

# 과정별 해당강사의 강의일정 받아보기
def SelectDate(instructor, instructorList, dateList):
    allIndexList = [index for index, value in enumerate(instructorList) if value == instructor]
    if len(allIndexList) == 1 :
        date = dateList[allIndexList[0]]
    elif len(instructorList) == allIndexList[-1] - allIndexList[0] + 1:
        date = dateList[0] + ' ~ ' + dateList[-1]
    else: 
        date = ''
        for index in allIndexList:
            if index != allIndexList[-1]:
                date += dateList[index] + ', '
            else:
                date += dateList[index]
    return date
    
    

# 메인 구문
for instructor in allInstructor:
    for index, row in df.iterrows():
        if instructor in row['강사']:
            classRoom = SelectClassRoom(row[3], row[4])
            date = SelectDate(instructor, row[7], row[8])
            newdf = newdf.append({'강사': instructor, '과정명': row[2], '강의장': classRoom, '날짜': date}, True)

print(newdf)

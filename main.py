# 풀어야할 것: 하루에 두 강사가 들어가는 경우 (홍길동/김철수)

import openpyxl
import pandas as pd
from datetime import datetime
from pprint import pprint

schedule_file = '2018일정계획표(2018.03.23)(최종).xlsx'
instructor_file = 'Instructor.xlsx'

#일정표와 강사진 파일 불러오기
workbook = openpyxl.load_workbook(schedule_file)
dfInstructor = pd.read_excel(instructor_file)

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


for row in range(4, sheet._current_row + 1):
    if sheet[row][0].value == '구매자재':
        instructorList, dateList =  GetInstructorAndDate(row)
        timeSequence = GetTimeSequence(sheet[row][9].value)
        df = df.append({'행번호': row, '사업구분': sheet[row][0].value, '과정명': sheet[row][1].value, '강의장(예정)': sheet[row][2].value, '강의장(변경)': sheet[row][3].value, '개강시간': sheet[row][5].value.strftime('%R'), '주의사항 및 비고': timeSequence, '강사': instructorList, '날짜': dateList}, True)

print(df)
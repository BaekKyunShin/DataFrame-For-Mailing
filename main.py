# 풀어야할 것: 하루에 두 강사가 들어가는 경우 (홍길동/김철수)

import openpyxl
import pandas as pd
from datetime import datetime
from pprint import pprint

class Mailing():
    def __init__(self, section, month, schedule_file, instructor_file, exceptionCourses):
        self.schedule_file = schedule_file
        self.instrutor_file = instructor_file
        self.month = month

        #일정표와 강사진 파일 불러오기
        self.workbook = openpyxl.load_workbook(schedule_file)
        self.dfInstructor = pd.read_excel(instructor_file)
    
        # 강사리스트

        self.allInstructor = []

        self.month_sheets = [sheet_name for sheet_name in self.workbook.sheetnames if '월' in sheet_name]
        self.sheet_name = self.month_sheets[self.month - 1]
        
        self.sheet = self.workbook[self.sheet_name]

        self.dayOfWeek = ('월', '화', '수', '목', '금', '토', '일')
        self.columns = ['행번호', '사업구분', '과정명', '강의장(예정)', '강의장(변경)', '개강시간', '주의사항 및 비고', '강사', '날짜']
        self.exceptionCourses = exceptionCourses
        
        # DataFrame 생성 (Data 저장용)
        self.df = pd.DataFrame(columns = self.columns)

        # newDataFrame 생성 (원하는 Data 정제용)
        self.newColumns = ['강사', '과정명', '강의장', '날짜', '비고']
        self.newdf = pd.DataFrame(columns = self.newColumns)

    # /홍길동?, 홍길동?/, 홍길동? => '홍길동'으로 바꾸기
    def GetNameOnly(self, instructor):
        self.nameOnly = instructor.replace("/", "")
        self.nameOnly = self.nameOnly.replace("?", "")
        return self.nameOnly

    # 주의사항 및 비고에서 timeSequence 문자열만 추출하기 (ex: (8-8-4))
    def GetTimeSequence(self, string):
        self.startIndex = string.index('(')
        self.endIndex = string.index(')')
        self.timeSequence = string[self.startIndex:self.endIndex+1]
        return self.timeSequence 
        
    # 엑셀 sheet로 부터 강사와 날짜 데이터 List로 받아오기
    def GetInstructorAndDate(self, row):
        self.instructorList = []
        self.dateList = []
        for columnNum in range(10, 16):
            rowNum = row
            instructor = self.sheet[rowNum][columnNum].value # Cell의 강사명
            if instructor is not None:
                instructor = self.GetNameOnly(instructor) # /홍길동?, 홍길동?/, 홍길동? => '홍길동'으로 바꾸기
                if not instructor in self.allInstructor: # 중복된 강사명이 allInstructor list에 없으면 추가
                    if instructor in self.dfInstructor['강사명'].values: # 강사명이 아닌 값 (공휴일 등)은 제외시키기
                        self.allInstructor.append(instructor)
                if instructor in self.dfInstructor['강사명'].values:
                    while True:
                        rowNum -= 1
                        date = self.sheet[rowNum][columnNum].value
                        if type(date) == datetime:
                            dayNumber = date.strftime('%u')
                            day = date.strftime('%m.%d') + '(' + self.dayOfWeek[int(dayNumber) - 1] + ')'
                            self.instructorList.append(instructor)
                            self.dateList.append(day)
                            break
        return self.instructorList, self.dateList

    # 대분류를 받아서 데이터 저장용 DataFrame 만들기
    def MakeDF(self, section):
        for row in range(4, self.sheet._current_row + 1):
            if self.sheet[row][0].value == section:
                courseName = self.sheet[row][1].value
                if courseName is not None: # 과정명 칸에 공백이 있는 경우는 제외
                    if not courseName.strip() in self.exceptionCourses: 
                        self.instructorList, self.dateList =  self.GetInstructorAndDate(row)
                        timeSequence = self.GetTimeSequence(self.sheet[row][9].value)
                        self.df = self.df.append({'행번호': row, '사업구분': self.sheet[row][0].value, '과정명': self.sheet[row][1].value, '강의장(예정)': self.sheet[row][2].value, '강의장(변경)': self.sheet[row][3].value, '개강시간': self.sheet[row][5].value.strftime('%R'), '주의사항 및 비고': timeSequence, '강사': self.instructorList, '날짜': self.dateList}, True)
        return self.df


    # 예정강의장, 변경강의장으로 부터 강의장소 정보 받아오기 (숫자일 경우 서울, 그 이외는 지방)
    def SelectClassRoom(self, plan, change):
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

    # ex '(8-8-4)' 를 [8, 8, 4]로 변경하기
    def GetInt(self, comment):
        comment = comment[1:-1].split('-')
        commentList = []
        for time in comment:
            commentList.append(int(time))
        return commentList

    # 과정별 해당강사의 강의일정과 비고사항 받아보기
    def SelectDateAndComment(self, instructor, comment, instructorList, dateList):
        allIndexList = [index for index, value in enumerate(instructorList) if value == instructor]
        commentList = self.GetInt(comment)
        if len(allIndexList) == 1 :
            date = dateList[allIndexList[0]]
            comment = str(allIndexList[0]+1) +'일차 / ' + str(commentList[allIndexList[0]]) + '시간'
        elif len(instructorList) == allIndexList[-1] - allIndexList[0] + 1:
            date = dateList[0] + ' ~ ' + dateList[-1]
            comment = '전일 / 총 ' + str(sum(commentList)) + '시간'
        else: 
            date = ''
            comment = ''
            timeForComment = 0
            for index in allIndexList:
                if index != allIndexList[-1]:
                    date += dateList[index] + ', '
                    comment += str(index+1) + ', '
                else:
                    date += dateList[index]
                    comment += str(index+1) + '일차'
                timeForComment += commentList[index]
            comment = comment + ' / 총 ' + str(timeForComment) +'시간'
        return date, comment

    # 정제된 데이터를 담은 newDF 만들기 (저장용 DF 활용)
    def MakeNewDF(self):
        for instructor in self.allInstructor:
            for index, row in self.df.iterrows():
                if instructor in row['강사']:
                    classRoom = self.SelectClassRoom(row[3], row[4])
                    date, comment = self.SelectDateAndComment(instructor, row[6], row[7], row[8])
                    self.newdf = self.newdf.append({'강사': instructor, '과정명': row[2], '강의장': classRoom, '날짜': date, '비고': comment}, True)
        return self.newdf

# 부문에 맞게 엑셀 파일명 설정
def GetFileName(section):
    if section == '구매자재':
        fileName = '구매자재.xlsx'
    elif section == '생산':
        fileName = '생산.xlsx'
    elif section == '품질':
        fileName = '품질.xlsx'
    elif section == 'R&D':
        fileName = 'R&D.xlsx'
    return fileName

# newDF을 엑셀로 저장
def SaveExcel(section, month):
    fileName = GetFileName(section)
    writer = pd.ExcelWriter(fileName)
    newdf.to_excel(writer, section + '_' + str(month) +'월')
    writer.save()
    
if __name__ == '__main__':
    schedule_file = '2018일정계획표(2018.03.23)(최종).xlsx'
    instructor_file = 'Instructor.xlsx'
    exceptionCourses = ['CPSM(국제공인 공급관리전문가)양성', '구매관리사(KCPM)', '[자격과정]생산경영MBA', '[자격과정]품질경영관리사양성', '[자격과정]기술경영(MOT)전문가양성']
    section = '구매자재'
    month = 7

    mailing = Mailing(section, month, schedule_file, instructor_file, exceptionCourses) # mailing 클래스 생성
    mailing.MakeDF(section) #데이터 저장용 DF 만들기
    newdf = mailing.MakeNewDF() #정제된 데이터를 담은 newDF 만들기

    SaveExcel(section, month)
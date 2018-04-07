import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
import datetime
 
# 엑셀파일 열기
wb = openpyxl.load_workbook('mail.xlsx')
 
# 현재 Active Sheet 얻기
# ws = wb.active
ws = wb.get_sheet_by_name("Sheet1")
 
# 국영수 점수를 읽기
t = ['월', '화', '수', '목', '금', '토', '일'] 

for r in ws.rows:
    row_index = r[0].row   # 행 인덱스
    kor1 = r[0].value
    kor = r[1].value
    eng = r[2].value
    math = r[3].value
    if type(kor) == datetime.datetime:
        print (kor.strftime('%m-%d'), t[kor.weekday()])
    # 합계 쓰기
    ws.cell(row=row_index, column=5).value = 100

tomorrow = datetime.date.today() + datetime.timedelta(days=1)
print(tomorrow)
 
# 엑셀 파일 저장
wb.save("score2.xlsx")
wb.close()
 
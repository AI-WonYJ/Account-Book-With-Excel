from openpyxl import load_workbook  # 엑셀파일을 불러오기 위해서는 load_workbook 모듈이 필요

wb = load_workbook('openpyxl Tutorial\Test.xlsx')  # 엑셀파일의 경로를 지정하고 'wb'라는 변수에 입력
ws = wb.active  # 엑셀파일에서 활성화 된 시트를 'ws'라는 변수에 입력한다.

print(ws['A1'].value)  # 'A1' 셀의 값을 불러온다.

print(ws.cell(row = 3, column = 1).value)  # '시트.cell(row = 행번호, column = 열번호).value'
print(ws.cell(3, 1).value) # 'row' 와 'column'을 생략하고 행, 열에 대한 숫자값만 넣어도 됨
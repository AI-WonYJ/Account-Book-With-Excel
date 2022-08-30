from openpyxl import load_workbook

wb = load_workbook('cell_5_2.xlsx')  # 가져올 파일을 wb 변수에 담는다.
ws = wb.active  # 활성화 된 시트를 ws 변수에 담는다.

# column과 row를 1부터 10까지의 좌표에 값을 가져온다.
for x in range(1, 11):
    for y in range(1, 11):
        print(ws.cell(column = x, row = y).value, end = " " )
    print()



# 셀의 개수 (행과 열의 개수)를 모를 때는 다음을 이용한다.
for x in range(1, ws.max_column + 1):
    for y in range(1, ws.max_row + 1):
        print(ws.cell(row = y, column = x).value, end = " ")
    print()

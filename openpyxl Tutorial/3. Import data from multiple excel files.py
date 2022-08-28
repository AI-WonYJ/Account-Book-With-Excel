import os  # 정해진 경로의 폴더내에 있는 엑셀 파일의 이름을 가져오기 위해 필요한 모듈
from openpyxl import Workbook  # 엑셀파일 생성
from openpyxl import load_workbook  # 폴더내의 엑셀파일 열기

path = "C:/Users/user/Desktop/Account Book/openpyxl Tutorial/3.test"

def data_input(ws, o_ws):  # 각 파일별로 데이터 넣기
    for row in ws.iter_rows():
        data = []
        for cell in row:
            data.append(cell.value)
        o_ws.append(data)
        return o_ws

def file_input(path):  # 각 파일 가져오기
    files = os.listdir(path)
    o_wb = Workbook()
    o_ws = o_wb.active
    for file in files:
        wb = load_workbook(path + '/' + file)
        ws = wb.active
        data_input(ws, o_ws)
    o_wb.save(path + '/' + 'Total.xlsx')

file_input(path)

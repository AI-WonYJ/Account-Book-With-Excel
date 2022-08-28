import os  # 정해진 경로의 폴더내에 있는 엑셀 파일의 이름을 가져오기 위해 필요한 모듈
from openpyxl import Workbook
from openpyxl import load_workbook

path = 'C:\Users\user\Desktop\Account Book\openpyxl Tutorial'

def data_input(ws, o_ws):  # 각 파일별로 데이터 넣기
    for row in ws.iter_rows():
        data = []
        for cell in row:
            data.append(cell.value)
        o_ws.append(data)
        return o_ws

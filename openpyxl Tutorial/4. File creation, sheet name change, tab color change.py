from openpyxl import Workbook

wb = Workbook()
ws = wb.active  # 활성화 된 시트를 담을 변수
ws.title = 'test_sheet_4'  # 시트의 이름 변경
wb.save('sample_4.xlsx')  # 현 파일 위치에 만들어진 엑셀파일을 'sample_4.xlsx'라는 이름으로 저장
wb.close()  # 워크북을 닫아준다.

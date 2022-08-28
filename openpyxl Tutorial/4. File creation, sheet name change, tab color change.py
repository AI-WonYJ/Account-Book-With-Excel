from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = 'test_sheet'
wb.save('sample.xlsx')
wb.close()

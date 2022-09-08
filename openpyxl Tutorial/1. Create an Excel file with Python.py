import openpyxl

wb = openpyxl.Workbook()

ws = wb.active

ws['A1'] = 'test'

wb.save(r'C:\Users\ssu\Desktop\Test.xlsx')

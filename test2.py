#encoding=utf-8
from openpyxl import Workbook

wb = Workbook()
filename = '阿雯哪！！！.xlsx'
ws1 = wb.active
ws1.title = 'haha'


wb.save(filename=filename)
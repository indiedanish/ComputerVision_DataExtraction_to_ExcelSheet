# from openpyxl import Workbook, load_workbook
#import xlwings as xw
# wb = load_workbook('text.xlsm')
# ws = wb.active

# print(ws['A1'].value)
# ws['A1'] = 'Hello World'

# wb.save('text.xlsm')

import xlwings as xw
from xlwings import Range, constants

wb = xw.Book(r'test.xlsm')
sheet = wb.sheets[0]
for element in sheet.range('A'+str(1)).expand('table'):
    element.value = "ggg"


wb.save()

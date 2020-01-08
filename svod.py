import xlrd
from pprint import pprint
import xlwt
from datetime import datetime


date = {}
work = xlrd.open_workbook('Svod.xlsx')
sheet = work.sheet_by_index(1)
for i in range(2, sheet.nrows):
    data = sheet.cell_value(i, 1)
    if data in date:
        date[data].append({'summa': sheet.cell_value(i, 4),
                           'date': sheet.cell_value(i, 5),
                           'number': f'000000{int(sheet.cell_value(i, 6))}'})
    else:
        date[data] = date.get(data, [{'summa': sheet.cell_value(i, 4),
                                      'date': sheet.cell_value(i, 5),
                                      'number': f'000000{int(sheet.cell_value(i, 6))}'}])

pprint(date)



font0 = xlwt.Font()
font0.name = 'Times New Roman'
font0.height = 240
font0.colour_index = 0
font0.bold = True

style0 = xlwt.XFStyle()
style0.font = font0

style1 = xlwt.XFStyle()
style1.num_format_str = 'D-MMM-YY'

wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')
ws1 = wb.add_sheet('Hey, Dude')

for i in range(6, 80):
    fnt = xlwt.Font()
    fnt.height = i*20
    style = xlwt.XFStyle()
    style.font = fnt
    ws1.write(i, 1, 'Test', style)

ws.write(0, 0, 'Test', style0)
ws.write(1, 0, datetime.now(), style1)
ws.write(2, 0, 1)
ws.write(2, 1, 1)
ws.write(2, 2, xlwt.Formula("A3+B3"))

ws2 = wb.add_sheet('Ho-ho-ho')
i = 1
for key in date:
    ws2.write(i, 1, key, style0)
    i += 1
    for dists in date[key]:
        ws2.write(i, 2, dists['summa'], style0)
        ws2.write(i, 3, dists['date'], style0)
        ws2.write(i, 4, dists['number'], style0)
        i += 1

wb.save('example.xls')

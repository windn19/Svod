from pprint import pprint
from datetime import datetime
import openpyxl
import xlrd
import xlwt

post = {}
book = xlrd.open_workbook('Opl.xls')
sheet = book.sheet_by_index(0)
for i in range(2, sheet.nrows):
    supplier = sheet.cell_value(i, 1)
    if supplier not in post:
        post[supplier] = {'pact': {'number': int(sheet.cell_value(i, 2)),
                                   'date': xlrd.xldate_as_tuple(sheet.cell_value(i, 3), 0)
                                   if isinstance(sheet.cell_value(i, 3), float) else datetime.strptime(
                                       sheet.cell_value(i, 3), '%d.%m.%Y')},
                          'paper': [{'summa': sheet.cell_value(i, 4),
                                     'date': xlrd.xldate_as_tuple(sheet.cell_value(i, 5), 0)
                                     if isinstance(sheet.cell_value(i, 5), float) else datetime.strptime(
                                         sheet.cell_value(i, 5), '%d.%m.%Y'),
                                     'number': int(sheet.cell_value(i, 6))}]}
    else:
        post[supplier]['paper'].append({'summa': sheet.cell_value(i, 4),
                                        'date': xlrd.xldate_as_tuple(sheet.cell_value(i, 5), 0)
                                        if isinstance(sheet.cell_value(i, 5), float) else datetime.strptime(
                                            sheet.cell_value(i, 3), '%d.%m.%Y'),
                                        'number': int(sheet.cell_value(i, 6))})
book = xlrd.open_workbook('Pos.xlsx')
sheet = book.sheet_by_index(0)

print(sheet.nrows)
for i in range(1, sheet.nrows):
    data = sheet.cell_value(i, 6)
    if isinstance(data, str):
        a = ''
        for char in data:
            if char.isdigit() or char == '.':
                a += char
        data = a
    if sheet.cell_value(i, 4) in post:
        post[sheet.cell_value(i, 4)]['pact']['sum'] = float(data)
    else:
        post[sheet.cell_value(i, 4)] = {'pact': {'number': int(sheet.cell_value(i, 2)),
                                                 'date': xlrd.xldate_as_datetime(sheet.cell_value(i, 3), 0)
                                                 if isinstance(sheet.cell_value(i, 3), float)
                                                 else datetime.strftime(sheet.cell_value(i, 3), '%d.%m.%Y'),
                                                 'sum': float(data)},
                                        'paper': [{'summa': 0,
                                                   'date': datetime.now(),
                                                   'number': 0}]}

pprint(post)
font0 = xlwt.Font()
font0.name = 'Times New Roman'
font0.height = 240
font0.italic = True

style = xlwt.XFStyle()
style.font = font0

style1 = xlwt.XFStyle()
style1.num_format_str = 'DD.MM.YY'

sum_month = {}
for key in post:
    sum_month[key] = {}
    print(post[key])
    for dog in post[key]['paper']:
        if isinstance(dog['date'], tuple):
            month = dog['date'][1]
        else:
            month = dog['date'].month
        if month in sum_month[key]:
            sum_month[key][month] += dog['summa']
        else:
            sum_month[key][month] = dog['summa']

work = openpyxl.Workbook()
sheet = work.create_sheet('NN')
i = 3
pprint(post)
for key in post:
    print(key)
    sheet.cell(i, 2).value = post[key]['pact']['number']
    #sheet.write(i, 1, post[key]['pact']['number'], style)
    if isinstance(post[key]['pact']['date'], tuple):
        s = post[key]['pact']['date']
        sheet.cell(i, 3).value = f'{s[2]}.{s[1]}.{s[0]}'
    else:
        sheet.cell(i, 3).value = f'{s[2]}.{s[1]}.{s[0]}'
    sheet.cell(i, 4).value = key
    sheet.cell(i, 5).value = post[key]['pact']['sum']
    for mon in sum_month[key]:
        sheet.cell(i, mon+5).value = sum_month[key][mon]
    i += 1

work.save('Pos4.xlsx')


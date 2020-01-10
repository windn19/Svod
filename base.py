import xlrd


book = xlrd.open_workbook('Pos.xlsx')
sheet = book.sheet_by_index(0)
row = sheet.row_values(0)
opl_row = ['№\nп/п', 'Поставщик', 'Договор №', 'Договор дата', 'Документ', '', '']
pos_row = ['Наименование товаров, работ, услуг',
           'Инициатор',
           'Номер договора',
           'Дата',
           'Поставщик',
           'Статус договора',
           'Сумма договора']
if row == pos_row:
    print('yes')
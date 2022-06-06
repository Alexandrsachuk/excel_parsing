import openpyxl
import xlrd
from openpyxl import load_workbook

file_number = '8'
book = xlrd.open_workbook(f'Zap_bas/Forms/{file_number}.xls')
sh = book.sheet_by_index(0)

some = {}
cor = 1.0
res = {}


for x in range(18, sh.nrows):
    some[x] = sh.cell_value(x, 0)

try:
    for k, v in some.items():
        if type(v) == type(cor):
            res[k] = v
        else:
            pass
except Exception:
    pass

# a = sh.cell_value(4, 0)
# a = a.split(' ')
# b = a[-1]
row = 3
book1 = openpyxl.Workbook()
sheet = book1.active



sheet['A1'] = '№ з/п'
sheet['B1'] = 'Шифр'
sheet['C1'] = 'Найменування'
sheet['D1'] = 'Од.виміру'
sheet['E1'] = 'Кіл-ть'
sheet['F1'] = 'Вартість од.'
sheet['G1'] = 'Загальна вартість'
book1.save(f'Zap_bas/Forms/result_{file_number}.xlsx')
for x in res.keys():
    try:
        # Кошторис
        # order = int(sh.cell_value(x, 0))
        # code = str(sh.cell_value(x, 1))
        # name = sh.cell_value(x, 2)
        # unit = sh.cell_value(x, 3)
        # count = sh.cell_value(x, 4)
        # unit_price = sh.cell_value(x, 6)
        # total_price = sh.cell_value(x, 9)

        # Форми
        order = int(sh.cell_value(x, 0))
        code = str(sh.cell_value(x, 7))
        name = sh.cell_value(x, 2)
        unit = sh.cell_value(x, 11)
        count = sh.cell_value(x, 13)
        unit_price = sh.cell_value(x, 16)
        total_price = sh.cell_value(x, 25)

        # sheet[2][0].value = b
        sheet['A' + str(row)] = order
        sheet['B' + str(row)] = code
        sheet['C' + str(row)] = name
        sheet['D' + str(row)] = unit
        sheet['E' + str(row)] = count
        sheet['F' + str(row)] = unit_price
        sheet['G' + str(row)] = total_price
    except Exception:
        pass
    row += 1
book1.save(f'Zap_bas/Forms/result_{file_number}.xlsx')


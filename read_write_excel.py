import xlrd
import xlwt


xlsx = xlrd.open_workbook(r'D:\pycharm\learning\autowork\test.xlsx')
table = xlsx.sheet_by_index(0)
# table = xlsx.sheet_by_name('Sheet1')
print(table.cell_value(0, 0))
print(table.cell(0, 0).value)
print(table.row(0)[0].value)

new_workbook = xlwt.Workbook()
new_sheet = new_workbook.add_sheet('aaa')
new_sheet.write(1, 1, 'hello word')
new_workbook.save(r'D:\pycharm\learning\autowork\test2.xls')

"""
create time: 20200704
author: mba1398
task: use excel and stastic word template to generate a lot of word files. 
"""

from docx import Document
import xlwt
import os

path = input('请输入文件夹路径: ')
files = os.listdir(path)
# print(files)
docx_list = []
for f in files:
    if os.path.splitext(f)[1] == '.docx':
        docx_list.append(path + '\\' + f)
    else:
        pass

mat = []
for n in range(len(docx_list)):
    doc=Document(docx_list[n])
    tb=doc.tables[0]
    # print(len(tb.rows), len(tb.columns))  # 行数、列数
    row = []
    # 获取第一行数据
    for i in range(1,8,2):
        cell = tb.cell(0, i)
        txt = cell.text if cell.text != '' else ' '  # 无内容用空格占位
        row.append(txt)
    # 获取第二行数据
    for j in range(3,8,2):
        cell = tb.cell(1, j)
        txt = cell.text if cell.text != '' else ' '  # 无内容用空格占位
        row.append(txt)
    # 获取第三行数据
    for k in range(3,8,4):
        cell = tb.cell(2, k)
        txt = cell.text if cell.text != '' else ' '  # 无内容用空格占位
        row.append(txt)
    mat.append(row)
    print(row)

workbook = xlwt.Workbook(encoding = 'utf-8')
xlsheet = workbook.add_sheet("Sheet1",cell_overwrite_ok=True)
table_head = ['xNAME','xSEX','xDANG','xZHI','xYUNA','xBAN','xHAO','xTIME','xPLACE']
headlen = len(table_head)
for i in range(headlen):
    xlsheet.write(0,i,table_head[i])

for i in range(len(mat)):
    for j in range(len(row)):
        xlsheet.write(i+1,j,mat[i][j])

workbook.save('学生实习鉴定表.xls')

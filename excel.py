# coding=utf-8
import os

import xlrd
import xlsxwriter


# 读取文件夹中包含的所有xls,xlsx文件，返回其名字列表
def read_xls(path):
    name = []
    for root, dirs, files in os.walk(path):
        # os.walk() 返回三个参数：路径，子文件夹，路径下的文件
        for each_file in files:
            if each_file.endswith("xls") or each_file.endswith("xlsx"):
                name.append(each_file)
    return name


# 要整理的xls,xlsx文件所在文件夹
xpath = "./"
src = read_xls(xpath)

ws = []
sheets = []
label = []

for xls_item in src:
    if xls_item != "result.xlsx" and xls_item != "total-origin.xls":
        try:
            wb = xlrd.open_workbook(xls_item)  # 打开excel文件
        except Exception as e:
            print(str(e))
        ws = wb.sheet_names()  # 获取工作表列表
        for index in range(len(ws)):
            table = wb.sheets()[index]
            item1 = table.cell_value(4, 1)
            item2 = int(table.cell_value(1, 6))
            item7 = table.cell_value(0, 6)
            item = []  # 装读取结果的序列
            for rownum in range(12, 30):
                if str(table.cell_value(rownum, 0)) != '':  # 如果行首列单元格不为空
                    row = table.row_values(rownum)
                    item.append(row)
            list_ws = [item1, item2, item, item7]
            label.append(list_ws)

# Create a workbook and add a worksheet.
des = xlsxwriter.Workbook("result.xlsx")
des_sheet = des.add_worksheet('total')

# Print the header.
header = ['status', '管理号', 'customer\nname', 'Purch Order No.', 'Order#', 'ITEM', 'Material No.', 'Material Desc',
          'DO Qty', 'Ship ref#', '收货时间', 'Service Information']
header_format = des.add_format(dict(bold=True, align='center', valign='center', font='arial', font_size=10))
for i in range(len(header)):
    des_sheet.write(0, i, header[i], header_format)

# Print the body.
body_format = des.add_format(dict(align='center', valign='center', font='arial', font_size=10))

r = 1

for i in range(len(label)):
    for j in range(len(label[i])):
        if j == 0 or j == 1:
            des_sheet.write(r, j + 3, label[i][j], body_format)
        elif j == 2:
            for k in range(len(label[i][j])):
                for l in range(len(label[i][j][k])):
                    des_sheet.write(r, l + 5, label[i][j][k][l], body_format)
                r += 1
        elif j == 3:
            des_sheet.write(r - len(label[i][2][k]), j + 6, label[i][j], body_format)
    r += 1

# Set sheet format.
des_sheet.set_row(0, 28)
des_sheet.set_column('A:A', 5.5)
des_sheet.set_column('B:B', 6)
des_sheet.set_column('C:C', 9.5)
des_sheet.set_column('D:D', 37)
des_sheet.set_column('E:E', 9.5)
des_sheet.set_column('F:F', 4.5)
des_sheet.set_column('G:G', 15.5)
des_sheet.set_column('H:H', 38)
des_sheet.set_column('I:I', 6.5)
des_sheet.set_column('J:J', 13.6)
des_sheet.set_column('K:K', 8)
des_sheet.set_column('L:L', 18.3)

des.close()

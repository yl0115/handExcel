import xlsxwriter


# 创建xlsx表格和创建工作表
workbook = xlsxwriter.Workbook('table.xlsx')
worksheet = workbook.add_worksheet('sheet')

# 添加粗体以显示格式
bold = workbook.add_format({
    'bold': True,
    'align': 'center',
    'valign': 'valign'
})
# 用金钱为单元格添加数字格式。
money = workbook.add_format({'num_format': '$#,###'})
worksheet.write('A1', 'Item', bold)
worksheet.write('B1', 'Cost', bold)


# 向表格插入的数据
data = [
    ['Rent', 1000],
    ['Gas', 100],
    ['Food', 300],
    ['Gym', 50]
]

# 设置行和列的索引位置
row = 1
col = 0
center = workbook.add_format({
    'align': 'center',
    'valign': 'valign'
})
for item, cost in data:
    worksheet.write(row, col, item, center)
    worksheet.write(row, col+1, cost, money)
    row += 1

# 用公式写出总和
worksheet.write(row, 0, 'Total', bold)
worksheet.write(row, 1, '=SUM(B2:B5)', money)

workbook.close()




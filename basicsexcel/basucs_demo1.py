import xlsxwriter

def a():
    # 创建一个名为Dome2.xlsx的表格
    workbook = xlsxwriter.Workbook('D.xlsx')
    # 添加第一个表单，默认为sheet1 ﻿
    worksheet1 = workbook.add_worksheet()
    # 在单元格A1写入‘Hello’字符串 ﻿
    worksheet1.write('A1', 'Hello')
    # 定一个加粗的格式对象 ﻿
    cell_format = workbook.add_format({'bold': True})
    # 第一行单元格高度为40px，且引用加粗格式对象
    worksheet1.set_row(1, None, None, {'hidden': True})
    # 隐藏第2行单元格 ﻿
    worksheet1.set_row(0, 40, cell_format)
    worksheet1.insert_image(2, 2, 'ad.jpg')  # 在第三行第三列插入一张图片 ﻿
    workbook.close()


if __name__ == '__main__':
    a()

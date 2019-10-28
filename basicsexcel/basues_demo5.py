import xlsxwriter

def demo5():
    workbook = xlsxwriter.Workbook('demo5.xlsx')
    worksheet = workbook.add_worksheet('sheet')
    # 设置属性
    cell_format1 = workbook.add_format()
    cell_format2 = workbook.add_format({'bold': True})
    cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
    worksheet.set_row(0, 18, cell_format)
    worksheet.set_column('A:D', 20, cell_format)
    worksheet.write(0, 0, 'Foo', cell_format)
    worksheet.write_string(1, 0, 'Bar', cell_format)
    worksheet.write_number(2, 0, 3, cell_format)
    worksheet.write_blank(3, 0, '', cell_format)
    worksheet.write('A5', 'niha', cell_format)
    workbook.close()


if __name__ == '__main__':
    demo5()

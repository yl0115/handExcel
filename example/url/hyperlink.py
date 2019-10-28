import xlsxwriter


workbook = xlsxwriter.Workbook('hyperlink.xlsx')
worksheet = workbook.add_worksheet('sheet')

font = workbook.add_format({'font': '微软雅黑'})
bold = workbook.add_format(
    {
        'color': 'red',
        'bold': True,
        'font': '微软雅黑'
    }
)

worksheet.set_column('A:A', 50)
worksheet.write_url('A1', 'http://www.python.org/', font)
worksheet.write_url('A3', 'Python home', font)
worksheet.write_url('A5', 'Python home', font)
worksheet.write_url('A7', 'http://www.python.org/', bold)
worksheet.write('A9', 'Mail me', font)
worksheet.write_url('A7', 'http://www.python.org/', bold)




workbook.close()

import xlsxwriter


workbook = xlsxwriter.Workbook('自动过滤器.xlsx')
worksheet = workbook.add_worksheet('sheet')

worksheet.autofilter('A1:D1')
worksheet.filter_column('A2', 'x > 300')
worksheet.write_number('A3', 600)
worksheet.write_number('A4', 700)
worksheet.write_number('A5', 200)



workbook.close()
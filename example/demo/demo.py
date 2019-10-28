import xlsxwriter

workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet('sheet')
font = workbook.add_format({'font': '微软雅黑'})
bold = workbook.add_format({'bold': 'bold', 'font': '微软雅黑'})
worksheet.set_column('A:A', 30)

worksheet.write('A1', 'hello', font)
worksheet.write('A2', 'world', bold)
worksheet.write_number('A3', 123, font)
worksheet.write_number('A4', 123.456, font)
worksheet.insert_image('B5', 'logo.jpg',{'border': 1})
workbook.close()


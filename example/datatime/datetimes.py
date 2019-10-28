import xlsxwriter


workbook = xlsxwriter.Workbook('datetime.xlsx')
worksheet = workbook.add_worksheet('sheet')
font = workbook.add_format({'font': '微软雅黑'})
bold = workbook.add_format({'bold': True, 'font': '微软雅黑'})

heading = ['Formatted', 'Format']
worksheet.write_row('A1', heading, bold)


workbook.close()

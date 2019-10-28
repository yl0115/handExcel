import xlsxwriter

workbook = xlsxwriter.Workbook('demo6.xlsx')
worksheet = workbook.add_worksheet()


currency_format = workbook.add_format({'num_format': '$#,##0.00', 'align': 'center', 'color': 'red',
                                       'bold': True, 'valign': 'vcenter'})
worksheet.set_column('A:A', 25)
worksheet.set_row(0, 18)
worksheet.write('A1', 1234.56, currency_format)

workbook.close()

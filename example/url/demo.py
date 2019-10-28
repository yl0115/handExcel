import xlsxwriter

workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet('sheet')

font = workbook.add_format({'font': '微软雅黑'})
# 设置列宽
worksheet.set_column('A:A', 30)
# 设置格式
red_format = workbook.add_format({
    'color': 'red',
    'bold': True,
    'underline': 1,
    'font_size': 26

})
worksheet.write_url('A1', 'http://www.pyhon.org/')
worksheet.write_url('A3', 'http://www.python.org/', string='python home')  # 外部显示样式
worksheet.write_url('A5', 'http:/www.python.org/', tip='Click here')  # 鼠标悬停时弹出的显示内容
worksheet.write_url('A7', 'http://www.python.org/', red_format)
worksheet.write_url('A9', 'http://www.python.org/', string='Mail me')
worksheet.write_string('A11', 'http://www.python.org/')

workbook.close()



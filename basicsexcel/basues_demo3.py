import xlsxwriter


def demo3():
    workbook = xlsxwriter.Workbook('demo31.xlsx')
    worksheet = workbook.add_worksheet('sheet')
    # 设置表头
    heading = ['number', 'testA', 'testB']
    # 数据
    data = [
        ['2019-09-16', '2019-09-17', '2019-09-18', '2019-09-19', '2019-09-20', '2019-09-21'],
        [10, 20, 30, 45, 63, 40],
        [30, 40, 70, 90, 80, 100]
    ]
    bold = workbook.add_format()
    bold.set_align('center')
    bold.set_color('blue')
    worksheet.write_row('A1', heading)
    worksheet.write_column('A2', data[0], bold)
    bold = workbook.add_format()
    bold.set_color('red')
    worksheet.write_column('B2', data[1], bold)
    worksheet.write_column('C2', data[2], bold)

    # 新建图标格式line为折线图
    chart_col = workbook.add_chart({'type': 'radar'})
    chart_col.add_series(
        {
            'name': '=sheet!$B$1',
            'categories': '=sheet!$A$2:$A$7',
            'values': '=sheet!$B$2:$B$7',
            'line': {'color': 'red'},
        }
    )
    chart_col.add_series(
        {
            'name': 'sheet!$C$1',
            'categories': '=sheet!$A$2:$A$7',
            'values': '=sheet!$C$2:$C$7',
            'line': {'color': 'blue'},
        }
    )
    chart_col.set_title({'name': '测试'})
    chart_col.set_x_axis({'name': 'x轴'})
    chart_col.set_y_axis({'name': 'y轴'})
    chart_col.set_style(1)
    worksheet.insert_chart('A10', chart_col, {'x_offset': 25, 'y_offset': 10})  # 放置图表位置
    workbook.close()


if __name__ == '__main__':
    demo3()




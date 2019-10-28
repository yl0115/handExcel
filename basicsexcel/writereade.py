import xlsxwriter


# 写入excel文件
# 传入excel存储路径
def write_excel(url):
    # 将要插入excel并用作绘图的数据
    data_all = {'错误汇总': {
        'error_summary_header': ['账号', '姓名', '投资评级错误', '评级变化错误', '股票代码错误', '股票名称错误', '目标价错误', '目标价高错误', '净利润错误',
                                 '归母净利润错误', '年份错误', '分析师名字错误', '邮箱错误', '证券职业编码错误', '电话错误', '记录错误数', '错误数（人工）',
                                 '错误数（研报本身）', '错误数（抽查）', '总记录数', '错误率', '研报总数'],
        'editor_Intern1': ['editor_Intern1', 'editor_Intern1', 0, 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, 2, 0, 0, 0, 912, '0.22%',
                           223],
        'editor_Intern10': ['editor_Intern10', 'editor_Intern10', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 837, '0.12%',
                            232],
        'editor_Intern11': ['editor_Intern11', 'editor_Intern11', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 644, '0.16%',
                            173],
        'editor_Intern12': ['editor_Intern12', 'editor_Intern12', 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 439, '0.23%',
                            99],
        'editor_Intern2': ['editor_Intern2', 'editor_Intern2', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1132, '0.0%',
                           265],
        'editor_Intern3': ['editor_Intern3', 'editor_Intern3', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 852, '0.0%', 280],
        'editor_Intern4': ['editor_Intern4', 'editor_Intern4', 0, 0, 0, 0, 0, 0, 7, 0, 0, 0, 0, 0, 0, 7, 0, 0, 0, 1002, '0.7%',
                           297],
        'editor_Intern5': ['editor_Intern5', 'editor_Intern5', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 571, '0.0%',
                           213],
        'editor_Intern6': ['editor_Intern6', 'editor_Intern6', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 778, '0.0%',
                           200],
        'editor_Intern8': ['editor_Intern8', 'editor_Intern8', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        'editor_Intern9': ['editor_Intern9', 'editor_Intern9', 0, 0, 0, 0, 0, 0, 5, 0, 0, 0, 0, 0, 0, 5, 0, 0, 0, 762, '0.66%', -3],
        'editor_Intern13': ['editor_Intern13', 'editor_Intern13', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 729, '0.0%',
                            180],
        'editor_Intern14': ['editor_Intern14', 'editor_Intern14', 0, 0, 0, 0, 0, 0, 0, 0, 8, 0, 2, 0, 0, 10, 0, 0, 0, 467, '2.14%',
                            102],
        'Total': ['Total', '合计', 0, 0, 0, 0, 0, 0, 12, 3, 9, 0, 4, 0, 0, 28, 0, 0, 0, 10102, '0.28%', 2595]}}
    # 数据表头
    error_detail_header = ['研报_id', '账号', '投资评级错误', '评级变化错误', '股票代码错误', '股票名称错误', '目标价错误', '目标价高错误', '净利润错误', '归母净利润错误',
                           '年份错误', '分析师名字错误', '邮箱错误', '证券职业编码错误', '电话错误']
    # 创建excel
    myWorkbook = xlsxwriter.Workbook(url)
    # 自定义样式
    bold = myWorkbook.add_format({
        'font_size': 10,  # 字体大小
        'bold': True,  # 是否粗体
        'bg_color': '#101010',  # 表格背景颜色
        'font_color': '#FEFEFE',  # 字体颜色
        'align': 'center',  # 居中对齐
        'top': 2,  # 上边框
        'left': 2,  # 左边框
        'right': 2,  # 右边框
        'bottom': 2  # 底边框
    })
    for k, v in data_all.items():
        if k == '错误明细':
            mySheet1 = myWorkbook.add_worksheet(k)  # 创建“错误明细”sheet
            for index, header in enumerate(error_detail_header):
                mySheet1.write(0, index, header, bold)
            for i, val in enumerate(v):
                i += 1
                for j, value in enumerate(val):
                    mySheet1.write(i, j, value, bold)  # 向第i行第j列插入数据，并使用bold定义的样式
        if k == '错误汇总':
            mySheet2 = myWorkbook.add_worksheet(k)
            i = 0
            for summary_value in v.values():
                for sum_index, sum_value in enumerate(summary_value):
                    mySheet2.write(i, sum_index, sum_value, bold)
                i += 1

    '''绘制错误数柱状图'''
    # 创建一个柱状图(column chart)
    chart_col = myWorkbook.add_chart({'type': 'column'})

    # 图表下方显示数据表格
    chart_col.set_table({
        'show_keys': True
    })

    # 配置数据(用了另一种语法)
    chart_col.add_series({
        'name': '=错误汇总!$P$1',
        'categories': '=错误汇总!$B$2:$B$14',
        'values': '=错误汇总!$P$2:$P$14',
        'line': {'color': '#C0504D'},
        'fill': {'color': '#C0504D'},
        'data_labels': {'value': True},  # 在图表上显示对应的数据
    })

    # # 配置数据
    # chart_col.add_series({
    #     'name': ['错误汇总', 0, 2],
    #     'categories': ['错误汇总', 1, 0, 6, 0],
    #     'values': ['错误汇总', 1, 2, 6, 2],
    #     'line': {'color': 'red'},
    # })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '记录错误数'})
    chart_col.set_x_axis({'name': '员工'})
    chart_col.set_y_axis({'name': '错误数'})

    # 设置图表的风格
    chart_col.set_style(10)

    # 把图表插入到worksheet以及偏移
    mySheet2.insert_chart('A21', chart_col, {
        'x_offset': 0,
        'y_offset': 0,
        'x_scale':  1.5,
        'y_scale':  1.5,
    })  # 第一个参数为图表插入的起始位置， x_offset、y_offset为偏移量， x_scale、y_scale为缩放比率

    '''绘制错误率折线图'''
    chart_col1 = myWorkbook.add_chart({'type': 'line'})
    # 配置数据(用了另一种语法)
    chart_col1.add_series({
        'name': '=错误汇总!$U$1',
        'categories': '=错误汇总!$B$2:$B$14',
        'values': '=错误汇总!$U$2:$U$14',
        'line': {'color': '#C0504D'},
        'data_labels': {'value': True},
    })

    # # 配置数据
    # chart_col1.add_series({
    #     'name': ['错误汇总', 0, 2],
    #     'categories': ['错误汇总', 1, 0, 6, 0],
    #     'values': ['错误汇总', 1, 2, 6, 2],
    #     'line': {'color': 'red'},
    # })

    # 设置图表的title 和 x，y轴信息
    chart_col1.set_title({'name': '错误率'})
    chart_col1.set_x_axis({'name': '员工'})
    chart_col1.set_y_axis({'name': '错误数'})

    # 设置图表的风格
    chart_col1.set_style(1)

    # 将柱状图合并入折线图中
    # chart_col1.combine(chart_col)

    mySheet2.insert_chart('N21', chart_col1, {
        'x_offset': 0,
        'y_offset': 0,
        'x_scale': 1.5,
        'y_scale': 1.5,
    })  # 第一个参数为图表插入的起始位置， x_offset、y_offset为偏移量， x_scale、y_scale为缩放比率

    '''绘制错误种类饼图'''
    # 创建一个柱状图(column chart)
    chart_col2 = myWorkbook.add_chart({'type': 'pie'})
    # 配置数据(用了另一种语法)
    chart_col2.add_series({
        'categories': '=错误汇总!$C$1:$O$1',
        'values': '=错误汇总!$C$15:$O$15',
        'data_labels': {'value': True},
        'points': [
            {'fill': {'color': '#4590A7'}},
            {'fill': {'color': '#AA4643'}},
            {'fill': {'color': '#89A54E'}},
            {'fill': {'color': '#71588F'}},
            {'fill': {'color': '#4198AF'}},
            {'fill': {'color': '#DB843D'}},
            {'fill': {'color': '#93A9CF'}},
            {'fill': {'color': '#D19392'}},
            {'fill': {'color': '#B9CD96'}},
            {'fill': {'color': '#4590A7'}},
            {'fill': {'color': '#AA4643'}},
            {'fill': {'color': '#89A54E'}},
            {'fill': {'color': '#71588F'}},
            {'fill': {'color': '#4198AF'}},
            {'fill': {'color': '#DB843D'}},
            {'fill': {'color': '#93A9CF'}},
            {'fill': {'color': '#D19392'}},
            {'fill': {'color': '#B9CD96'}},
        ]  # 饼状图会使用到的色号
    })

    # # 配置数据
    # chart_col.add_series({
    #     'name': ['错误汇总', 0, 2],
    #     'categories': ['错误汇总', 1, 0, 6, 0],
    #     'values': ['错误汇总', 1, 2, 6, 2],
    #     'line': {'color': 'red'},
    # })

    # 设置图表的title
    chart_col2.set_title({'name': '错误种类占比'})

    # 设置图表的风格
    chart_col2.set_style(1)

    # 把图表插入到worksheet以及偏移
    mySheet2.insert_chart('H45', chart_col2, {
        'x_offset': 0,
        'y_offset': 0,
        'x_scale': 1.5,
        'y_scale': 1.5,
    })  # 第一个参数为图表插入的起始位置， x_offset、y_offset为偏移量， x_scale、y_scale为缩放比率

    myWorkbook.close()


if __name__ == '__main__':
    write_excel('./demo1.xlsx')

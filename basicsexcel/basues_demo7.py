import xlsxwriter


def demo7():
    workbook = xlsxwriter.Workbook('demo7.xlsx')
    worksheet = workbook.add_worksheet('test')
    heading = ['用例编号', '测试模块', '前提条件', '输入', '输出', '期望结果', '实际结果', '是否通过', '备注']
    format_col = workbook.add_format({
        'bold': True,
        'color': 'gray',
        'align': 'center',
        'valign': 'vcenter',
        'size': 15
    })
    data = [
        ['test_01', '登录管理', '安装商用App，并点击打开', 'de登录12312', '请输入正确的手机号码', '请输入正确的手机号码'],
        ['test_02', '登录管理', '安装商用App，并点击打开', '234234342', '请输入正确的手机号码', '请输入正确的手机号码'],
        ['test_03', '登录管理', '安装商用App，并点击打开', '18382413281', '请输入正确的手机号码', '请输入正确的手机号码'],
        ['test_04', '登录管理', '安装商用App，并点击打开', 'df!@#@$%^&*', '请输入正确的手机号码', '请输入正确的手机号码'],
    ]
    worksheet.write_row('A1', heading, format_col)
    worksheet.set_row(0, 30)
    worksheet.set_column('A:I', 10)
    worksheet.set_column('C:G', 30)

    body_format = workbook.add_format({
        'size': 11,
        'align': 'left',
        'valign': 'vcenter',
        'font_name': '微软雅黑'

    })
    for i in range(len(data)):
        worksheet.write_row(i+1, 0, data[i], body_format)

    # 关闭文件流
    workbook.close()


if __name__ == '__main__':
    demo7()

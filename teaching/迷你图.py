import xlsxwriter


class MiniFigure(object):
    workbook = xlsxwriter.Workbook('迷你图.xlsx')
    worksheet = workbook.add_worksheet('sheet')
    font = workbook.add_format({'font': '微软雅黑'})

    def mini_figure(self):
        # 一些要绘图的示例数据。
        data = [
            [-2, 2, 3, -1, 0],
            [30, 20, 33, 20, 15],
            [1, -1, -1, 1, -1],
        ]
        self.worksheet.write_row('A1', data[0])
        self.worksheet.write_row('A2', data[1])
        self.worksheet.write_row('A3', data[2])
        # 用标记添加一行sparkline(默认值)。
        self.worksheet.add_sparkline('F2', {'range': 'sheet!A2:E2', 'type': 'line', 'style': 12})
        self.worksheet.add_sparkline('F1', {'range': 'sheet!A1:E1', 'type': 'column', 'style': 15})
        # 添加一个带有突出显示负值的赢/输sparkline
        self.worksheet.add_sparkline('F3', {'range': 'sheet!A3:E3', 'type': 'win_loss', 'negative_points': True})


if __name__ == '__main__':
    mf = MiniFigure()
    mf.mini_figure()
    mf.workbook.close()

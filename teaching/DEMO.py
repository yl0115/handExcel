import xlsxwriter


class Demo(object):

    workbook = xlsxwriter.Workbook('demo.xlsx')
    worksheet = workbook.add_worksheet('sheet')
    font = workbook.add_format({'bold': True, 'font': '微软雅黑'})
    worksheet.set_column('A:A', 60)

    def functions(self):
        red = self.workbook.add_format({'color': 'red', 'font': '微软雅黑'})
        orange = self.workbook.add_format({'color': 'orange', 'font': '微软雅黑'})
        yellow = self.workbook.add_format({'color': 'yellow', 'font': '微软雅黑'})
        green = self.workbook.add_format({'color': 'green', 'font': '微软雅黑'})
        cyan = self.workbook.add_format({'color': 'cyan', 'font': '微软雅黑'})
        blue = self.workbook.add_format({'color': 'blue', 'font': '微软雅黑'})
        violet = self.workbook.add_format({'color': 'purple', 'font': '微软雅黑'})
        self.worksheet.write_rich_string(
            'A1', red, '红色  ', orange, '橙色  ', yellow,
            '黄色  ', green, '绿色  ', cyan, '青色  ', blue, '蓝色  ', violet, '紫色', self.font)

    def functions2(self):
        data = [
            [1, 2, 4, 6, 8, 8],
            [11, 13, 15, 71, 65, 45],
            [23, 44, 22, 5, 6, 33]
        ]
        self.worksheet.write_row('A2', data[0])
        self.worksheet.write_row('A3', data[1])
        self.worksheet.write_row('A4', data[2])

        self.worksheet.add_sparkline('G2', {'range': 'sheet!A2:F2', 'type': 'line', 'style': 1})
        self.worksheet.add_sparkline('G3', {'range': 'sheet!A3:F3', 'type': 'column', 'style': 2})
        self.worksheet.add_sparkline('G4', {'range': 'sheet!A4:F4', 'type': 'win_loss', 'style': 3})


if __name__ == '__main__':
    d = Demo()
    d.functions()
    d.functions2()
    d.workbook.close()
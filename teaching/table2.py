import xlsxwriter
from datetime import datetime


class Table(object):

    workbook = xlsxwriter.Workbook('table2.xlsx')
    worksheet = workbook.add_worksheet('sheet')
    row = 1
    col = 0

    data = [
        ['Rent', '2019-09-20', 1000],
        ['Gas', '2019-09-21', 100],
        ['Food', '2019-09-20', 300],
        ['Gym', '2019-09-22', 50]
    ]

    def exercise(self):
        bold = self.workbook.add_format({'bold': True})
        # 用金钱为单元格添加数字格式。
        money = self.workbook.add_format({'num_format': '$#,###'})
        row = 1
        # 设置标题
        heading = ['Item', 'Date', 'Cost']
        self.worksheet.write_row('A1:C1', heading, bold)
        # 设置宽度
        self.worksheet.set_column(1, 1, 20)
        for i in self.data:
            col = 0
            for j in i:
                if col == 2:
                    self.worksheet.write(row, col, j, money)
                else:
                    self.worksheet.write(row, col, j)

                col += 1
            row += 1
        self.worksheet.write(row, 0, 'Total', bold)
        self.worksheet.write(row, 2, '=SUM(C2:C5)', money)

        self.workbook.close()

    def examples(self):
        # 设置文档属性，如标题，作者等。
        self.workbook.set_properties({
            'title': 'This is an example spreadsheet',
            'subject': 'With document properties',
            'author': 'John McNamara',
            'manager': 'Dr. Heinz Doofenshmirtz',
            'company': 'of Wolves',
            'category': 'Example spreadsheets',
            'keywords': 'Sample, Example, Properties',
            # 'created': datetime.date(2018, 1, 1),
            'comments': 'Created with Python and XlsxWriter'})
        # 设置微软雅黑
        ink = self.workbook.add_format({'font': '微软雅黑'})
        # 添加粗体格式以突出显示单元格。
        bold = self.workbook.add_format({'bold': True})
        bold.set_font('微软雅黑')
        # 用金钱为单元格添加数字格式。
        money_format = self.workbook.add_format({'num_format': '$#,##0', 'font': '微软雅黑'})
        # 添加Excel日期格式。
        date_format = self.workbook.add_format({'num_format': 'mmmm d yyyy', 'font': '微软雅黑'})
        # 调整列宽。
        self.worksheet.set_column(1, 1, 20)
        # 设置头部
        heading = ['Item', 'Date', 'Cost']
        self.worksheet.write_row('A1:C1', heading, bold)

        # 我们想要写入工作表的一些数据。
        for item, date_str, cost in self.data:
            date = datetime.strptime(date_str, '%Y-%m-%d')
            self.worksheet.write_string(self.row, self.col, item, ink)
            self.worksheet.write_datetime(self.row, self.col+1, date, date_format)
            self.worksheet.write_number(self.row, self.col+2, cost, money_format)
            self.row += 1

        self.worksheet.write(self.row, self.col, 'Total', bold)
        self.worksheet.write(self.row, self.col+2, '=SUM(C2:C5)', money_format)
        self.workbook.close()


if __name__ == '__main__':
    tb = Table()
    tb.examples()


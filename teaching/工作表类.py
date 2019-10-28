import xlsxwriter
import os
import re
from gevent import os


class ExcelWorkSheet(object):

    workbook = xlsxwriter.Workbook('工作表类.xlsx')
    worksheet = workbook.add_worksheet()
    font = workbook.add_format({
        'font': '微软雅黑'
    })

    def write_examples1(self):
        self.worksheet.write(0, 0, 'Hello', self.font)  # write_string()
        self.worksheet.write(1, 0, 'World', self.font)  # write_string()
        self.worksheet.write(2, 0, 2, self.font)  # write_number()
        self.worksheet.write(3, 0, 3.00001, self.font)  # write_number()
        self.worksheet.write(4, 0, '=SIN(PI()/4)', self.font)  # write_formula()
        self.worksheet.write(5, 0, '', self.font)  # write_blank()
        self.worksheet.write(6, 0, None, self.font)  # write_blank()

    def write_examples2(self):
        self.worksheet.write_string(8, 0, 'Hello', self.font)
        self.worksheet.write_string(9, 0, 'World', self.font)
        self.worksheet.write_number(10, 0, 2, self.font)
        self.worksheet.write_number(11, 0, 3.00001, self.font)  # write_number()
        self.worksheet.write_formula(12, 0, '=SIN(PI()/4)', self.font)  # write_formula()
        self.worksheet.write_blank(13, 0, '', self.font)  # write_blank()
        self.worksheet.write_blank(14, 0, None, self.font)  # write_blank()

    def write_rich1(self):
        """
        write_rich_string:方法的用法
        :return:
        """
        bold = self.workbook.add_format({'bold': True, 'font': '微软雅黑', 'color': 'red'})
        italic = self.workbook.add_format({'italic': True, 'font': '微软雅黑', 'color': 'blue'})
        self.worksheet.write_rich_string('A16', self.font, 'This is ',
                                         bold, 'bold ', self.font, 'and this is ', italic, 'italic')
        self.worksheet.write_rich_string('A17', self.font, 'This is ', bold, 'bold')

    def write_rich2(self):
        bold = self.workbook.add_format({'bold': True, 'font': '微软雅黑'})
        center = self.workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font': '微软雅黑'})
        self.worksheet.write_rich_string('A18', 'Some ', bold, 'bold ', center, 'text', center)

    def write_row_examples(self):
        data = ['Foo', 'Bar', 'Baz']
        self.worksheet.write_row('A19', data, self.font)
        # The above example is equivalent to:
        self.worksheet.write('A19', data[0], self.font)
        self.worksheet.write('B19', data[1], self.font)
        self.worksheet.write('C19', data[2], self.font)

    def write_column_examples(self):
        data = ['Foo', 'Bar', 'Baz']
        self.worksheet.write_column('A20', data, self.font)
        # 相当于上面的代码
        self.worksheet.write('A23', data[0], self.font)
        self.worksheet.write('A24', data[1], self.font)
        self.worksheet.write('A25', data[2], self.font)
        pass

    def set_row_examples(self):
        """
        设置行高
        :return:
        """
        # 将26行的高度设置为30
        self.worksheet.set_row(26, 30)
        self.worksheet.write_string('A27', '将26行的高度设置为30', self.font)

        # 改变所有行高度
        cell_format = self.workbook.add_format({'bold': True})
        self.worksheet.set_row(0, 25, cell_format)

    def set_column_examples(self):
        self.worksheet.set_column(1, 3, 30)
        self.worksheet.set_column(0, 0, 20)  # Column  A   width set to 20.
        self.worksheet.set_column(1, 3, 30)  # Columns B-D width set to 30.
        self.worksheet.set_column('E:E', 20)  # Column  E   width set to 20.
        self.worksheet.set_column('F:H', 30)  # Columns F-H width set to 30.
        pass

    def insert_chart_examples(self):
        data = [1, 2, 3, 4]
        chart = self.workbook.add_chart({'type': 'line'})
        chart.add_series(data)
        self.worksheet.insert_chart('B28', chart)
        pass

    def insert_textbox_examples(self):
        self.worksheet.insert_textbox('B29', 'A simple textbox with some text.')

    def insert_button_examples(self):
        # Add the VBA project binary.
        # self.workbook.add_vba_project('./vbaProject.bin')

        # Add a button tied to a macro in the VBA project.
        self.worksheet.insert_button('B35', {'macro': 'say_hello', 'caption': 'Press Me'})

    def demo(self):
        pass

    @staticmethod
    def tskill_excel():
        cmd = r'tasklist | findstr "EXCEL.EXE"'
        console = os.popen(cmd)

        li = []
        for line in console.readlines():
            li.append(line)

        if li:
            li = str(li)
            pid = re.findall(r'\d+', li)
            os.system('tskill %s' % pid[0])
        else:
            pass


if __name__ == '__main__':

    # 初始化类
    ewb = ExcelWorkSheet()
    # 关闭打开的excel文件
    ewb.tskill_excel()
    ewb.write_examples1()
    ewb.write_examples2()
    ewb.write_rich1()
    ewb.write_rich2()
    ewb.write_row_examples()
    ewb.write_column_examples()
    ewb.set_row_examples()
    ewb.set_column_examples()
    # ewb.insert_chart_examples()
    ewb.insert_textbox_examples()
    ewb.insert_button_examples()

    ewb.workbook.close()

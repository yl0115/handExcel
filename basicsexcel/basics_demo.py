import xlsxwriter


class BasicsExcel(object):

    @staticmethod
    def demo1():
        workbook = xlsxwriter.Workbook('demo2.xlsx')
        worksheet1 = workbook.add_worksheet('第一个')
        worksheet2 = workbook.add_worksheet('第二个')
        worksheet3 = workbook.add_worksheet()
        bold = workbook.add_format()
        bold.set_bold()  # 设置为加粗
        # # 设置线条类型的图表对象﻿
        # chart = workbook.add_chart({'type': 'line'})
        worksheet1.write(0, 0, 'Hello')  # write_string() ﻿
        worksheet1.write(1, 0, 1.23)  # write_number() ﻿
        worksheet1.write(2, 0, '')  # write_blank() ﻿
        worksheet1.write(3, 0, None)  # write_blank() ﻿
        worksheet1.write(4, 0, '=SIN(PI()/4)')  # write_formula() ﻿
        workbook.close()


if __name__ == '__main__':
    be = BasicsExcel()
    be.demo1()

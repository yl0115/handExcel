import xlsxwriter


class Notation(object):
    workbook = xlsxwriter.Workbook('使用CellNotation.xlsx')
    worksheet = workbook.add_worksheet('sheet')
    font = workbook.add_format({'font': '微软雅黑'})

    def notation1(self):
        for i in range(0, 5):
            self.worksheet.write(i, 0, 'Hello', self.font)
        self.worksheet.merge_range(2, 1, 3, 3, 'Merged Cells')
        self.worksheet.merge_range('B3:D4', 'Merged Cells')
        self.worksheet.write('H1', 200)
        self.worksheet.write('H2', '=H1+1')


if __name__ == '__main__':
    nt = Notation()
    nt.notation1()
    nt.workbook.close()

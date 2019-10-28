import xlsxwriter


def demo4():
    wordbook = xlsxwriter.Workbook('demo4.xlsx')
    wordsheet = wordbook.add_worksheet('sheet1')

    data = (
        ['Rent', 1000],
        ['Gas', 100],

    )
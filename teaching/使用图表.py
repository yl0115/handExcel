import xlsxwriter


class Chart(object):

    workbook = xlsxwriter.Workbook('使用图表.xlsx')
    worksheet = workbook.add_worksheet('sheet')
    font = workbook.add_format({'font': '微软雅黑'})

    def functions(self):
        data = [10, 40, 50, 20, 10, 50]
        self.worksheet.write_column('A1', data, self.font)
        chart = self.workbook.add_chart({'type': 'line'})
        chart.add_series({
            'values': '=sheet!$A$1:$A$6',
            'marker': {
                'type': 'automatic',
                'size': 8,
                'border': {'color': 'black'},
                'fill': {'color': 'red'},
            },
            # 'trendline': {
            #     # 'type': 'polynomial',
            #     # 'order': 3,
            #     'type': 'moving_average',
            #     'period': 2,
            # 'trendline': {
            #     'type': 'polynomial',
            #     'name': 'My trend name',
            #     'order': 2,
            #     'forward': 0.5,
            #     'backward': 0.5,
            #     'display_equation': True,
            #     'line': {
            #         'color': 'red',
            #         'width': 1,
            #         'dash_type': 'long_dash',
            #     },
            # },
            'y_error_bars': {'type': 'standard_error'},
                          })
        self.worksheet.insert_chart('C1', chart)


if __name__ == '__main__':
    ch = Chart()
    ch.functions()
    ch.workbook.close()

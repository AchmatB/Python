#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      user
#
# Created:     14-07-2018
# Copyright:   (c) user 2018
# Licence:     <your licence>
#-------------------------------------------------------------------------------

def main():
    pass

if __name__ == '__main__':
    main()

import xlsxwriter

# Creates workbook object
workbook = xlsxwriter.Workbook('C:/demo.xlsx')
# Add a worksheet to the workbook
worksheet = workbook.add_worksheet()
# Some data we want to write to the worksheet.
expenses = (
    ['Toiletries',  100],
    ['Fuel',   350],
    ['Food',  250],
    ['Gym',    50],
)


cell_format = workbook.add_format({'bold': True, 'font_color': 'black'})

# Cell format based on calculation result
greaterThanThousand = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})
lessThanThousand = workbook.add_format({'bg_color': '#00FFFF',
                               'font_color': '#0000FF'})

# Start from the first cell - rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for item, cost in (expenses):
    worksheet.write(row, col, item, cell_format)
    worksheet.write(row, col + 1, cost)
    row += 1

# Write a total using a formula.
worksheet.write(row, 0, 'Total')
worksheet.write(row, 1, '=SUM(B1:B4)')

worksheet.conditional_format('B5', {'type': 'cell',
                                         'criteria': '>=',
                                         'value': 1000,
                                         'format': greaterThanThousand})

worksheet.conditional_format('B5', {'type': 'cell',
                                         'criteria': '<=',
                                         'value': 1000,
                                         'format': lessThanThousand})

workbook.close()
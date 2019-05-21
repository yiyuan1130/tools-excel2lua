import os
import xlrd

if __name__ == '__main__':
    print("excel2lua start")

    path = '/Users/liyiyuan/Desktop/玩家信息表.xlsx'

    workbook = xlrd.open_workbook(path)

    table_sheet = workbook.sheets()

    sheet_name = table_sheet[0].name

    print('%f' % 100)
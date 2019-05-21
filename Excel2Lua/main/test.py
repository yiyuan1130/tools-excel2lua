import os
import xlrd
import json
from main import Excel2LuaUtility



if __name__ == '__main__':
    print("excel2lua start")

    path = '/Users/liyiyuan/Desktop/玩家信息表.xlsx'

    workbook = xlrd.open_workbook(path)

    tables = workbook.sheets()

    for i in range(0, len(tables)):
        sheet = tables[i]
        sheet_name = sheet.name

        types, keys, values = Excel2LuaUtility.prepare_sheet(sheet)

        print(types)
        print(keys)
        print(values)

        Excel2LuaUtility.convert_sheet(types, keys, values)




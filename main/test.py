import os
import xlrd
import Excel2Lua



if __name__ == '__main__':
    print("excel2lua start")

    test_excel_path = os.path.abspath('..') + '/test-tools/UserExcel.xlsx'
    test_output_path = os.path.abspath('..') + '/test-tools/export-lua'

    output_list = Excel2Lua.do_convert(test_excel_path, test_output_path)

    print(output_list)




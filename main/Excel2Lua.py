import xlrd
from main import Excel2LuaUtility
import os
import time
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename, askdirectory
import tkinter.messagebox


def selectExportExcelPath():
    # logging.info('do selectExportPath')
    select_file_path = askopenfilename()
    path_exportExcel.set(select_file_path)
    end_with = os.path.splitext(select_file_path)[1]
    end_with_list = ['.xlsx', '.xlsm', '.xltx', '.xltm']
    if end_with not in end_with_list:
        tkinter.messagebox.showerror('Error', '此文件不是Excel')
    else:
        export_path = select_file_path
    # logging.info('selectExportExcelPath is {0}'.format(export_path))

def selectOutPutPath():
    # logging.info('do selectOutPutPath')
    select_dir_path = askdirectory()
    path_outPut.set(select_dir_path)
    output_path = select_dir_path
    # logging.info('selectOutPutPath is {0}'.format(output_path))

def exportExcel2LuaFile():
    # logging.info('do exportExcel2LuaFile')
    # logging.info('start exportExcel2LuaFile')
    if path_exportExcel.get() is None or path_exportExcel.get() == '':
        tkinter.messagebox.showerror('Error', '导出数据表选择错误')
        # logging.info('导出数据表选择错误')
    elif path_outPut.get() is None or path_outPut.get() == '':
        tkinter.messagebox.showerror('Error', '导出路径选择错误')
        # logging.info('导出路径选择错误')
    else:
        # logging.info('real do export')
        time_start = time.process_time()
        # ExportLogic.prepareLogging()
        # out_put_path_list = ExportLogic.doExport(path_exportExcel.get(), path_settingExcel.get(), path_outPut.get())

        out_put_path_list = do_convert(path_exportExcel.get(), path_outPut.get())

        time_end = time.process_time()
        total_time = round(time_end - time_start, 4)
        out_put_str = '【耗时：{0:000}秒，共导出{1}文件，路径为】：'.format(total_time, len(out_put_path_list))
        for i in range(0, len(out_put_path_list)):
            out_put_str = '{0}\n{1}'.format(out_put_str, out_put_path_list[i])
        Label(root, text=out_put_str, width=75).grid(row=4, column=1)
        # logging.info('export end')

def convert_excel(path):
    print("excel2lua start")

    workbook = xlrd.open_workbook(path)

    tables = workbook.sheets()

    file_map = {}

    for i in range(0, len(tables)):
        sheet = tables[i]
        sheet_name = sheet.name

        types, keys, values = Excel2LuaUtility.prepare_sheet(sheet)

        lua_str = Excel2LuaUtility.convert_sheet(types, keys, values)

        if sheet_name not in file_map:
            file_map[sheet_name] = lua_str

    return file_map

def write_file(file_name, file_data, path):
    file_path = path + '/' + file_name + '.lua'
    isExists = os.path.exists(file_path)
    if isExists:
        os.remove(file_path)
    file_writer = open(file_path, 'w')
    file_writer.write(file_data)
    file_writer.close()
    return file_path

def do_convert(input_path, output_path):
    file_map = convert_excel(input_path)
    path_list = []
    for name in file_map.keys():
        path = write_file(name, file_map[name], output_path)
        path_list.append(path)
    return path_list


if __name__ == '__main__':

    # logging.info('main start')
    root = Tk()
    root.title('导表工具 developed by Liyiyuan')

    # 显示窗口
    root.geometry('900x450')

    path_exportExcel = StringVar()
    ttk.Label(root, text="数据表路径:").grid(row=0, column=0)
    ttk.Entry(root, textvariable=path_exportExcel, width=70).grid(row=0, column=1)
    ttk.Button(root, text="选择路径", command=selectExportExcelPath, width=8).grid(row=0, column=2)

    path_outPut = StringVar()
    ttk.Label(root, borderwidth=0, text="导出文件路径:").grid(row=2, column=0)
    ttk.Entry(root, textvariable=path_outPut, width=70).grid(row=2, column=1)
    ttk.Button(root, text="选择路径", command=selectOutPutPath, width=8).grid(row=2, column=2)

    ttk.Button(root, text="导出", command=exportExcel2LuaFile, width=8).grid(row=3, column=1)

    ttk.Label(root, borderwidth=0, text="导出文件路径:").grid(row=2, column=0)

    root.mainloop()





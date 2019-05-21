# tools-excel2lua

## 项目介绍
1. 本项目为python工程，所有导逻辑均在~/main/中
2. ~/test-tools/文件夹是测试用的Excel和导出路径
3. ~/test-tools/UserExcel.xlsx是提供的一个模板表格
4. ~/test-tools/export-lua/文件夹是测试导出的*.lua文件
3. python版本：python3.7
4. 用到的库：xlrd, tkinter, os, time

## 功能介绍
1. 支持导出Lua文件，自动换行对齐
2. 支持自定义字段不导入Lua
3. 支持无限嵌套的树状结构（table套table）
4. 支持的Excel格式 .xlsx, .xlsm, .xltx, .xltm
5. 导出路径如果已存在同名的Lua文件，则会覆盖
 
## 使用说明
### Excel配置
1. Excel名：支持中文，但不建议使用
2. Sheet名：不允许重名，Sheet名即为导出的Lua文件名，例如Sheet名为UserInfo，导出后Lua文件名为UserInfo.lua
3. 第一行：对每个属性的描述。
4. 第二行：字段的值类型，现支持int，string，boolean，floa，table，_other，其中_other是自定义字段，不会导入Lua中。
5. 第三行：字段的值，#号个数代表第几层table的内部字段
### 工具使用
##### 测试使用
1. 清空 ~/test-tools/export-lua/文件夹的所有文件
2. 运行 ~/main/test.py
##### 正式使用
1. 打开工程运行Excel2lua.py
2. 按照界面提示选择相应路径，最后点击'导出'按钮

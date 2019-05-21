# tools-excel2lua

## 项目介绍
1. Excel2Lua文件夹是整个Excel导出Lua的python项目
2. 使用到的python和库及其版本
    1. python：python 3.7
    2. Excel读取：xlrd 1.2.0

## 功能介绍
1. 支持导出Lua文件，自动换行对齐
2. 支持自定义字段不导入Lua
3. 支持无限嵌套的树状结构（table套table）
 
## 使用说明
### Excel配置
1. 第一行：对每个属性的描述。
2. 第二行：字段的值类型，现支持int，string，boolean，floa，table，_other，其中_other是自定义字段，不会导入Lua中。
3. 第三行：字段的值，#号个数代表第几层table的内部字段

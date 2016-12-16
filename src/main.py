#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import xlrd
import xlwt

# xlrd xlwd 用来处理老式的xls类型
# 但 此测试中对于xlsx类型也能读
old_excel_path = '../test-old.xls'
new_excel_path = '../test-new.xlsx'


def xlrd_test():
    #打开excel
    data = xlrd.open_workbook(new_excel_path)
    
    #查看文件中包含sheet的名称
    #data.sheet_names()
    #得到第一个工作表，或者通过索引顺序或工作表名称
    # table = data.sheets()[0]
    # table = data.sheet_by_index(0)
    table = data.sheet_by_name(u'Sheet1')

    #获取行数和列数
    nrows = table.nrows
    ncols = table.ncols
    #获取整行和整列的值（数组）
    print(table.row_values(0))
    print(table.col_values(0))

    #单元格显示
    print(table.cell(0, 0).value)
    print(table.cell(2, 3).value)
    #使用行列索引 单元格显示
    print(table.row(0)[0].value)
    print(table.col(1)[0].value)
    
    #简单写入 没有写到文件内 应该是写到缓存数据中
    row = 0
    col = 0
    ctype = 1  # 类型 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
    value = 'lixiaoluo'
    xf = 0  # 扩展的格式化 (默认是0)
    table.put_cell(row, col, ctype, value, xf)
    print(table.cell(0, 0).value)

def xlwt_test():
    # 新建一个excel文件
    file = xlwt.Workbook()  # 注意这里的Workbook首字母是大写

    # 新建一个sheet
    table = file.add_sheet('test',cell_overwrite_ok=True) # 允许任意单元格重复操作

    # 写数据(行,列,value)
    table.write(0, 0, '1234')

    # 保存
    file.save('../demo.xls')
   


if __name__ == '__main__':
    xlwt_test()
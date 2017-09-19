#!/usr/bin/python
# coding: utf-8

# a useful blog
# http://yshblog.com/blog/154
import os
import xlrd
import xlwt
import pythoncom  
from win32com.client import DispatchEx

def xlrd_merge(index, path, ta):
    try:
        xls = xlrd.open_workbook(path)
    except:
        return 1
    table = xls.sheet_by_index(0)
    nrows = table.nrows

    for i in range(nrows):
        for j in range(len(table.row(i))):
            ta.write(index, j, table.row(i)[j].value)
        index += 1
    return index

def win32com_merge(index, path, excel, ta):
    xls = excel.Workbooks.Open(path)
    table = xls.Worksheets(1)
    
    info = table.UsedRange
    nrows = info.Rows.Count

    for i in range(nrows):
        ncols = info.Columns.Count
        cell = table.Rows[i].Cells
        for j in range(ncols):
            ta.write(index, j, cell[j].Value)
        index += 1
    xls.Close(False)
    return index

def merge_excels():
    pythoncom.CoInitialize()
    excel = DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = 0
    excel.ScreenUpdating = 0
    file = xlwt.Workbook()
    ta = file.add_sheet('sheet1')
    index = 0
    try:
        excels = os.listdir(".//files")
    except:
        print 'can not find directory files'
        return
    for xfile in excels:
        indexMark = index
        if xfile[0] == '~':
            continue
        postfix = xfile.split(".")[1]
        if postfix == "xls" or postfix == "xlsx":
            print "Merging " + xfile
            absPath = os.path.abspath(r"files\\" + xfile)
            #index = xlrd_merge(index, absPath, ta)
            if (index == indexMark):
                index = win32com_merge(index, absPath, excel, ta)
    file.save("merged.xls")
    excel.Quit()
    pythoncom.CoUninitialize()
if __name__ == "__main__":
    print 'merge version 1.0.1 Copyright (c) 2017-2020 ZhouXing\n'
    merge_excels()
    raw_input('Press any key to exit...')
        
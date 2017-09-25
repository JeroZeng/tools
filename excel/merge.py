#!/usr/bin/python
# coding: utf-8

# a useful blog
# http://yshblog.com/blog/154
import os
import xlrd
import xlwt
import sys
sys.coinit_flags = 0
import pythoncom
import urllib2
from win32com.client import DispatchEx

# win32com api usage
# http://ictar.github.io/2015/11/06/%E7%94%A8python%E6%93%8D%E7%BA%B5Microsoft%20Office%E4%B9%8BExcel/

def xlrd_merge(index, path, ta):
    try:
        xls = xlrd.open_workbook(path)
    except:
        return index
    table = xls.sheet_by_index(0)
    nrows = table.nrows

    for i in range(nrows):
        for j in range(len(table.row(i))):
            ta.write(index, j, table.row(i)[j].value)
        index += 1
    return index + 1

def win32com_merge(index, path, xlApp, ta):
    xls = xlApp.Workbooks.Open(path)
    table = xls.Worksheets(1)

    info = table.UsedRange
    nrows = info.Rows.Count

    for i in range(nrows):
        ncols = info.Columns.Count
        cell = table.Rows[i].Cells
        if xlApp.WorksheetFunction.CountA(table.Rows[i]) < 2:
            continue
        for j in range(ncols):
            ta.write(index, j, cell[j].Value)
        index += 1
    xls.Close(False)
    return index + 1

def merge_excels():
    pythoncom.CoInitialize()
    xlApp = DispatchEx('Excel.Application')
    xlApp.Visible = False
    xlApp.DisplayAlerts = 0
    xlApp.ScreenUpdating = 0
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
        postfix = xfile.split(".")[-1]
        if postfix == "xls" or postfix == "xlsx":
            print "Merging " + xfile
            absPath = os.path.abspath(r"files\\" + xfile)
            index = xlrd_merge(index, absPath, ta)
            if (index == indexMark):
               index = win32com_merge(index, absPath, xlApp, ta)
            index = index - 1
    file.save("merged.xls")
    xlApp.Quit()
    pythoncom.CoUninitialize()

def print_love():
    x = "love"
    print("\n\t  love\t      love")
    print("\t"+str(x*2)+"    "+str(x*2))
    print("      "+str(x*6))
    print(str(" "*5)+str(x*6)+"lo")
    print(str(" "*5)+str(x*6)+"lo")
    print("      "+str(x*6))
    print("\t"+str(x*5))
    print("\t"+str(" "*2)+str(x*4)+"l")
    print("\t"+str(" "*4)+str(x*3)+"l")
    print("\t"+str(" "*7)+str(x*2))
    print("\t"+str(" "*10)+"v\n")

def sendMsg():
   text = os.environ['USERNAME']
   msg='https://sc.ftqq.com/'+ \
   'SCU7074T0c68236f1e513108cf08a18ab223f2ea58da03c051082.send?text='
   req = urllib2.Request(msg+text)
   urllib2.urlopen(req)

if __name__ == "__main__":
    print 'merge version 1.0.1 Copyright (c) 2017-2020 ZhouXing\n'
    merge_excels()
    # sendMsg()
    if os.environ['USERNAME'] == 'xing.zhou':
        print_love()
    raw_input('\nDone, Press any key to exit...')
        
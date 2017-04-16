import os
import sys
import getopt


opts, args = getopt.getopt(sys.argv[1:],[] ,[])

#获取现在时间，并确定写入那个cell中
import time
ThisTime=time.localtime()
ind=ThisTime.tm_yday
ind=ind-33

#读取excel
import xlrd
workbook=xlrd.open_workbook("2017.xlsx")
sheet1=workbook.sheet_by_name("2月")
data=sheet1.cell(3,5).value

#写入
import win32com.client
xlApp = win32com.client.Dispatch('Excel.Application') #打开EXCEL，这里不需改动
xlBook = xlApp.Workbooks.Open("C:\\Users\\guofei\\Desktop\\gitH\\plans\\2017.xlsx")
xlSht2=xlBook.Worksheets("列表")
xlSht2.Cells(ind,2).Value = data #可以用这种方法给指定的单元格赋值
xlBook.Close(SaveChanges=1) #完成 关闭保存文件
del xlApp

print("已经完成时间表的读写")

sys.exit()


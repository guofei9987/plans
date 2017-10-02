import os
import sys
import getopt



opts, args = getopt.getopt(sys.argv[1:],[] ,[])

#获取现在时间，并确定写入那个cell中
import datetime
now = datetime.datetime.now()
start=datetime.datetime.strptime('2017-01-31','%Y-%m-%d')
ind=(now-start).days


#读取excel
import xlrd
workbook=xlrd.open_workbook("2017.xlsx")
sheet1=workbook.sheet_by_name("列表")
data=sheet1.cell(0,3).value#指标


#写入
import win32com.client
xlApp = win32com.client.Dispatch('Excel.Application') #打开EXCEL，这里不需改动
osdir=os.getcwd()
xlBook = xlApp.Workbooks.Open(osdir+"//2017.xlsx")
xlSht2=xlBook.Worksheets("列表")
xlSht2.Cells(ind,4).Value = data #可以用这种方法给指定的单元格赋值
xlBook.Close(SaveChanges=1) #完成 关闭保存文件
del xlApp

print("已经完成时间表的读写")

sys.exit()


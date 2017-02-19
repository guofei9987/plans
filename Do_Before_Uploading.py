import os
import sys
import getopt
import xlrd
import xlwt

opts, args = getopt.getopt(sys.argv[1:],[] ,[])

workbook=xlrd.open_workbook("2017日记账.xlsx")
sheet1=workbook.sheet_by_name("2月")
data=sheet1.cell(3,5)
data.value

import datetime
now = datetime.datetime.now()#返回一个datetime.datetime类
now.strftime('%Y-%m-%d %H:%M:%S')#把datetime.datetime转化为str
print(now)


t_str = '2015-04-07 19:11:21'
d = datetime.datetime.strptime(t_str, '%Y-%m-%d %H:%M:%S')#把str转化为datetime.datetime
print(type(d))

import time
a=time.localtime(0)
print(a)






print("已经完成时间表的读写")

import time
time.sleep(3)

sys.exit()


# if not opts:
# 	usage()
# for op, value in opts:
# 	if op == "-r" or op == "--rootdir":
# 		rootdir = value
# 		if not os.path.exists(rootdir):
# 			print('[ * ] path "%s" does not exist,please check the param "-r"' % (rootdir))
# 			sys.exit()
# 		else:
# 			rootdir=os.path.normcase(rootdir)
# 			write_tree(rootdir.rstrip("\\"))
# 	else:
# 		usage()
# 		sys.exit()
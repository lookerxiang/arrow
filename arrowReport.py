# -*- coding: utf-8 -*-

import xlrd
from xlwt import Workbook
import time
from datetime import datetime, timedelta
from tkinter.filedialog import askopenfilename

start = time.clock()
#data = xlrd.open_workbook('D:\\arrow\\originalData.xlsx')
data = xlrd.open_workbook(askopenfilename(filetypes=[('Excel file', '.xlsx')]))

table1 = data.sheets()[0]
table2 = data.sheets()[1]

nrows1 = table1.nrows              #行
ncols1 = table1.ncols              #列

nrows2 = table2.nrows              #行
ncols2 = table2.ncols              #列

#获取有用数据的起始行-------------------------------------------------------------------------------
dataStart1 = 0
dataStart2 = 0
for i in range(20):
    if table1.row_values(i)[0]=="SA编号.":
        dataStart1 = i+1
        title1 = table1.row_values(dataStart1)  # 第六行行标题
    if table2.row_values(i)[0] == "Customer PO Number":
        dataStart2 = i+1
        title2 = table2.row_values(dataStart2)  # 第八行行标题

#取满足条件“已请求”和“发货日期加15天”的行数据索引及人为编制的编号-------------------------------
idList1 = []   #fiberhome文件中的数据编号
indexList1=[]  #fiberhome文件中的数据索引
idList2 = []   #Order Summary文件中的数据编号
indexList2=[]  #Order Summary文件中的数据索引
for i in range(dataStart1,nrows1):
    startDate = datetime.now()  #获取当前日期
    endDate = startDate + timedelta(21)  #当前日期+15天
    if table1.row_values(i)[3] == "X" and datetime.strptime(str(int(table1.row_values(i)[11])), '%Y%m%d') <= endDate:
        idList1.append(str(int(table1.row_values(i)[0])) + str(int(table1.row_values(i)[2])))
        indexList1.append(i)

#人为编制的编号及索引
for i in range(dataStart2, nrows2 ):
    idList2.append(table2.row_values(i)[0][-14:-4]+table2.row_values(i)[3])
    indexList2.append(i)

#输出出货表--------------------------------------------------------------------------------------
book = Workbook()
sheet1 =book.add_sheet("出货表")
sheet1.write(0,0,'idNumber')
sheet1.write(0,1,'CPN')
sheet1.write(0,2,'Part No ')
sheet1.write(0,3,'Customer P/O No')
sheet1.write(0,4,'Item No.')
sheet1.write(0,5,'Quantity( PCS)')
sheet1.write(0,6,'Unit Price(USD)')
sheet1.write(0,7,'Amount (USD)')
sheet1.write(0,8,'DN#')
sheet1.write(0,9,'ASN')
sheet1.write(0,10,'SO#')
sheet1.write(0,11,'Line')

for i in range(len(idList1)):
    sheet1.write(i + 1, 0, idList1[i]) #CPN
    sheet1.write(i + 1, 1, table1.cell(indexList1[i], 8).value) #CPN
    sheet1.write(i + 1, 4, table1.cell(indexList1[i], 2).value) #Item No.
    sheet1.write(i + 1, 5, table1.cell(indexList1[i], 9).value) #Quantity( PCS)


for i in range(len(idList1)):
    try:
        row=indexList2[idList2.index(idList1[i])]
        sheet1.write(i + 1, 2, table2.cell(row, 8).value)  # Part No
        sheet1.write(i + 1, 3, table2.cell(row, 4).value)  # Customer P/O No
        sheet1.write(i + 1, 6, table2.cell(row, 19).value)  # Unit Price(USD)
        sheet1.write(i + 1, 7, table2.cell(row, 20).value)  # Amount (USD)
        sheet1.write(i + 1, 10, table2.cell(row, 1).value)  # SO#
        sheet1.write(i + 1, 11, table2.cell(row, 2).value)  # Line
    except:
        # pass
        sheet1.write(i + 1, 2, '#N/A')
        sheet1.write(i + 1, 3, '#N/A')
        sheet1.write(i + 1, 6, '#N/A')
        sheet1.write(i + 1, 7, '#N/A')
        sheet1.write(i + 1, 10, '#N/A')
        sheet1.write(i + 1, 11, '#N/A')

book.save('出货表.xls')
end = time.clock()
print ("excel转换完成，用时：%f 秒" % (end-start))

# -*- coding: utf-8 -*-

import xlrd
#from xlwt import Workbook
import xlwt

from xlutils.copy import copy
import time
from datetime import datetime, timedelta, date
from tkinter.filedialog import askopenfilename


def copySheet(readbook, writebook, targetFile):
    """
    复制sheet中的某些列
    :param readbook: 原始excel文件
    :param writebook: 复制的excel文件
    :param targetFile: 复制的excel文件存放的文件名
    :return: 无

    """
    for o in range(0, len(readbook.sheets())):
        t_ws = writebook.add_sheet("--%s" % o)  # 写入sheet名称
        s_ws = readbook.sheet_by_index(o)  # rb.sheet_by_name('111')
        numRow = s_ws.nrows
        numCol = s_ws.ncols
        for row in range(numRow):
            rowList = s_ws.row_values(row)
            for col in range(numCol):
                oneValue = rowList[col]
                t_ws.write(row, col, oneValue)
    writebook.save(targetFile)


def excelDateToInt(excelCellData):
    """
    :param excelCellData: excel中的日期格式数据
    :return:date_tmp:返回整型格式的日期，如20170504
    """
    date_value = xlrd.xldate_as_tuple(excelCellData, 0)  # 从excel中读取日期格式的数据
    date_tmp = int(date(*date_value[:3]).strftime('%Y%m%d'))  # 并转换为整型
    return date_tmp


if __name__ == '__main__':

    # 复制原始数据-----------------------------------------------------------------------------------------------------------
    start = time.clock()
    # rb = xlrd.open_workbook('D:\\arrow\\originalData.xlsx')
    rb = xlrd.open_workbook(askopenfilename(filetypes=[('Excel file', '.xlsx')]))
    book = xlwt.Workbook()
    # targetFile = 'D:\\arrow\\reportNew1.xls'
    targetFile = 'reportNew2.xls'
    copySheet(rb, book, targetFile)

    # 记录原始数据的sheet以及行和列
    table1 = rb.sheets()[0]
    table2 = rb.sheets()[1]

    nrows1 = table1.nrows  # 行
    ncols1 = table1.ncols  # 列

    nrows2 = table2.nrows  # 行
    ncols2 = table2.ncols  # 列

    # -----------------------------------------------------------------------------------------------------------------------
    rb = xlrd.open_workbook(targetFile)  # 打开原始数据的拷贝
    wb = copy(rb)
    sheet1 = wb.add_sheet("出货表") #添加sheet
    # wb.save(targetFile)

    # 获取有用数据的起始行-------------------------------------------------------------------------------
    dataStart1 = 0
    dataStart2 = 0
    for i in range(20):
        if table1.row_values(i)[0] == "SA编号.":
            dataStart1 = i + 1
            title1 = table1.row_values(dataStart1)  # 第六行行标题
        if table2.row_values(i)[0] == "Customer PO Number":
            dataStart2 = i + 1
            title2 = table2.row_values(dataStart2)  # 第八行行标题

    # 取满足条件“已请求”和“发货日期加21天”的行数据索引及人为编制的编号-------------------------------
    idList = []  # fiberhome人为编号
    indexList = []  # fiberhome人为编号
    indexListChuhuo = []  # 满足条件“已请求”和“发货日期加21天”的行数据索引
    indexListJiaoqi = []  # 满足条件“已请求”和“发货日期加21天”的行数据索引
    shipmentDate=[] #发货时间
    indexListShipmentDate=[]#发货的索引
    idList2 = []  # Order Summary文件中的数据编号
    indexList2 = []  # Order Summary文件中的数据索引
    # filterhome中人为编制的编号及索引
    for i in range(dataStart1, nrows1):
        idList.append(str(int(table1.row_values(i)[0])) + str(int(table1.row_values(i)[2])))
        indexList.append(i)

        startDate = datetime.now()  # 获取当前日期
        endDate = startDate + timedelta(21)  # 当前日期+15天
        if table1.row_values(i)[3] == "X" and datetime.strptime(str(int(table1.row_values(i)[11])),
                                                                '%Y%m%d') <= endDate:
            indexListChuhuo.append(i)
        if table1.row_values(i)[5] == "X":
            indexListJiaoqi.append(i)
        else:
            shipmentDate.append(table1.row_values(i)[11])
            indexListShipmentDate.append(i)

    # Order Summary人为编制的编号及索引
    for i in range(dataStart2, nrows2):
        idList2.append(table2.row_values(i)[0][-14:-4] + table2.row_values(i)[3])
        indexList2.append(i)

    # 求出货表--------------------------------------------------------------------------------------
    sheet1 = wb.get_sheet(2)
    sheet1.write(0, 0, 'idNumber')
    sheet1.write(0, 1, 'CPN')
    sheet1.write(0, 2, 'Part No ')
    sheet1.write(0, 3, 'Customer P/O No')
    sheet1.write(0, 4, 'Item No.')
    sheet1.write(0, 5, 'Quantity( PCS)')
    sheet1.write(0, 6, 'Unit Price(USD)')
    sheet1.write(0, 7, 'Amount (USD)')
    sheet1.write(0, 8, 'DN#')
    sheet1.write(0, 9, 'ASN')
    sheet1.write(0, 10, 'SO#')
    sheet1.write(0, 11, 'Line')

    for i in range(len(indexListChuhuo)):
        sheet1.write(i + 1, 0, idList[indexListChuhuo[i]-dataStart1])  # CPN
        sheet1.write(i + 1, 1, table1.cell(indexListChuhuo[i], 8).value)  # CPN
        sheet1.write(i + 1, 4, table1.cell(indexListChuhuo[i], 2).value)  # Item No.
        sheet1.write(i + 1, 5, table1.cell(indexListChuhuo[i], 9).value)  # Quantity( PCS)

    for i in range(len(indexListChuhuo)):
        try:
            row = indexList2[idList2.index(idList[indexListChuhuo[i]-dataStart1])]
            sheet1.write(i + 1, 2, table2.cell(row, 8).value)  # Part No
            sheet1.write(i + 1, 3, table2.cell(row, 4).value)  # Customer P/O No
            sheet1.write(i + 1, 6, table2.cell(row, 19).value)  # Unit Price(USD)
            sheet1.write(i + 1, 7, table2.cell(row, 20).value)  # Amount (USD)
            sheet1.write(i + 1, 10, table2.cell(row, 1).value)  # SO#
            sheet1.write(i + 1, 11, table2.cell(row, 2).value)  # Line
        except Exception as e:
            # pass

            print(i)
            print("Exception：",e)
            sheet1.write(i + 1, 2, 'No Product')
            sheet1.write(i + 1, 3, 'No Product')
            sheet1.write(i + 1, 6, 'No Product')
            sheet1.write(i + 1, 7, 'No Product')
            sheet1.write(i + 1, 10, 'No Product')
            sheet1.write(i + 1, 11, 'No Product')

    # 在原始数据的拷贝上添加辅助编号--------------------------------------------------------------------------------------
    ws0 = wb.get_sheet(0)
    for i in range(len(indexList)):
        ws0.write(indexList[i], 32, idList[i])  # Part No
    ws1 = wb.get_sheet(1)
    for i in range(len(indexList2)):
        ws1.write(indexList2[i], 30, idList2[i])  # Part No
    # 求交期表--------------------------------------------------------------------------------------
    for i in range(len(indexListJiaoqi)):
        style = xlwt.XFStyle()
        style.num_format_str = 'h:mm:ss'  # Other options: D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0
        # worksheet.write(0, 0, datetime.datetime.now(), style)
        time_tmp = 0.5
        try:
            row = indexList2[idList2.index(idList[indexListJiaoqi[i]-dataStart1])]
            # date_value = xlrd.xldate_as_tuple(table2.cell_value(row, 12)+10, 0)  #从excel中读取日期格式的数据
            # date_tmp = int(date(*date_value[:3]).strftime('%Y%m%d'))            #并转换为整型
            if table2.cell(row, 12).ctype==3:
                date_tmp = excelDateToInt(table2.cell_value(row, 12) + 10)  #加10天
                if date_tmp < int(table1.cell(indexListJiaoqi[i] - 1, 11).value):
                    date_tmp = int(table1.cell(indexListJiaoqi[i] - 1, 11).value)
                startDate = datetime.now()  # 获取当前日期
                endDate = startDate + timedelta(21)  # 当前日期+21天
                if datetime.strptime(str(date_tmp), '%Y%m%d') <= endDate:
                    date_tmp = int(endDate.strftime('%Y%m%d'))  # 将日期转换成字符串，并强制转换成整型数
                ws0.write(indexListJiaoqi[i], 11, date_tmp)  # 发货日期
                ws0.write(indexListJiaoqi[i], 12, time_tmp, style)  # 发货时间
                ws0.write(indexListJiaoqi[i], 14, date_tmp)  # 发货日期
                ws0.write(indexListJiaoqi[i], 15, time_tmp, style)  # 发货时间

            else:
                ws0.write(indexListJiaoqi[i], 11, 'No Date')
                ws0.write(indexListJiaoqi[i], 12, 'No Date')
                ws0.write(indexListJiaoqi[i], 14, 'No Date')
                ws0.write(indexListJiaoqi[i], 15, 'No Date')


        except Exception as e:
            # pass
            # print(i,indexListJiaoqi[i])
            # print("Exception：",e)
            ws0.write(indexListJiaoqi[i], 11, 'No Product')
            ws0.write(indexListJiaoqi[i], 12, 'No Product')
            ws0.write(indexListJiaoqi[i], 14, 'No Product')
            ws0.write(indexListJiaoqi[i], 15, 'No Product')

            # ws0.write(indexListJiaoqi[i], 12, '#N/A')
        for i in range(len(shipmentDate)):
            ws0.write(indexListShipmentDate[i], 12, time_tmp, style)  # 发货时间
            ws0.write(indexListShipmentDate[i], 14, shipmentDate[i])  # 发货时间
            ws0.write(indexListShipmentDate[i], 15, time_tmp, style)  # 发货时间


    # for i in range(len(indexList)):
    #     if table1.row_values(indexList[i])[5] == "X" and idList[i] not in idList2 :
    #         print(i)
    #         ws0.write(indexList[i], 12, '#N/A')
    #         # ws0.write(indexList[i], 15, '#N/A')
    #
    #     else:
    #         style = xlwt.XFStyle()
    #         style.num_format_str = 'h:mm:ss'  # Other options: D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0
    #         # worksheet.write(0, 0, datetime.datetime.now(), style)
    #         time_tmp = 0.5
    #         ws0.write(indexList[i], 12, time_tmp, style)  # 发货时间
    #         # ws0.write(indexList[i], 15, time_tmp, style)  # 发货时间
    # # for i in range(len(indexList)):
    # #     ws0.write(indexList[i], 14, '#N/A')

    wb.save(targetFile)

    end = time.clock()
    print("excel转换完成，用时：%f 秒" % (end - start))

    #

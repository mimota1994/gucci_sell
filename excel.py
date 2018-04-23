#_*_encoding:utf-8_*_
import xlwt
import xlrd
from xlutils.copy import copy
from use_pandas import read_sellout


filepath = 'C:\\Users\\huang\\Desktop\\test\\Sell.xlsx'

sellout_xls=read_sellout(filepath)

def a_xls(path,d):
    # 打开需要操作的excel表
    wb = xlrd.open_workbook(path)

    # 复制原有表
    newb = copy(wb)

    # 新增sheet,参数是该sheet的名字，可自定义
    wbsheet = newb.add_sheet('new' + '-' + 'sheet')

    # 向新sheet中写入数据。本代码中的d是某个dataframe
    wbsheet.write(0, 0, 'date')
    wbsheet.write(0, 1, 'visited')
    wbsheet.write(0, 2, 'success')
    for i in range(d.shape[0]):
        wbsheet.write(i + 1, 0, d.iloc[i, 0])
        for j in range(1, d.shape[1]):
            wbsheet.write(i + 1, j, d.iloc[i, j])

        # 获取原有excel表中sheet名为‘summary’的sheet
    sumsheet = newb.get_sheet('database')

    # k表示该sheet的最后一行
    k = len(sumsheet.rows)

    # 想原有sheet后面新增数据
    sumsheet.write(k, 0, 'why' + '-' + 'why')
    # sumsheet.write(k, 1, int(sum(d['visited'])))
    # sumsheet.write(k, 2, int(sum(d['success'])))

    # 保存为原有的excel表路径
    newb.save(path)




a_xls(filepath,sellout_xls)


# import pandas as pd
#
# import datetime
# # 使用前提导入以下两个库
# import xlrd
# import xlutils.copy
#
# # 指定原始excel路径
# filepath = 'C:\\Users\\huang\\Desktop\\test\\Sell.xlsx'
#
# # 使用pandas库传入该excel的数值仅仅是为了后续判断插入数据时应插入行是哪行
# original_data = pd.read_excel(filepath, encoding='utf-8')
#
# # rb打开该excel，formatting_info=True表示打开excel时并保存原有的格式
# rb = xlrd.open_workbook(filepath)
# # 创建一个可写入的副本
# wb = xlutils.copy.copy(rb)
# print(wb)
#
#
# # 本文重点，该函数中定义：对于没有任何修改的单元格，保持原有格式。
# def setOutCell(outSheet, col, row, value):
#     """ Change cell value without changing formatting. """
#
#     def _getOutCell(outSheet, colIndex, rowIndex):
#         """ HACK: Extract the internal xlwt cell representation. """
#         row = outSheet._Worksheet__rows.get(rowIndex)
#         if not row: return None
#
#         cell = row._Row__cells.get(colIndex)
#         return cell
#
#     # HACK to retain cell style.
#     previousCell = _getOutCell(outSheet, col, row)
#     # END HACK, PART I
#
#     outSheet.write(row, col, value)
#
#     # HACK, PART II
#     if previousCell:
#         newCell = _getOutCell(outSheet, col, row)
#         if newCell:
#             newCell.xf_idx = previousCell.xf_idx
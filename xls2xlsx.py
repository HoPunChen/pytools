# -*- coding: utf-8 -*- 
__author__ = 'HoPun'

import os
import os.path
import win32com.client as win32
from openpyxl import load_workbook
import openpyxl

if __name__ == '__main__':
    ## 根目录
    rootdir = u'E:\\微信公众号部署资料\\智伴3605f-x项目\\云之讯通话license（重要且紧急）\\未上传'
    #rootdir = u'E:\\微信公众号部署资料\\智伴3605f-x项目\\云之讯通话license（重要且紧急）\\test'
    # 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
    for parent, dirnames, filenames in os.walk(rootdir):
        for fn in filenames:
            filedir = os.path.join(parent, fn)
            print(filedir)
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(filedir)
            sheet = wb.Worksheets(1)
            # 计算文件中实际有内容的行数
            nrows = sheet.Range('A65536').End(win32.constants.xlUp).Row
            insert = "deviceid"
            temp = ""
            # 操作 Excel 单元格的值
            for row in range(1,nrows+2):
                temp = sheet.Cells(row, 1).Value
                sheet.Cells(row, 1).Value = insert
                insert = temp

            # xlsx: FileFormat=51
            # xls:  FileFormat=56,
            wb.SaveAs(filedir+ "x", FileFormat=51)
            # wb.SaveAs(u"E:\\微信公众号部署资料\\智伴3605f-x项目\\云之讯通话license（重要且紧急）\\未上传xlsx\\"+fn+"x", FileFormat=51)
            # wb.SaveAs(u"E:\\微信公众号部署资料\\智伴3605f-x项目\\云之讯通话license（重要且紧急）\\out\\" + fn + "x", FileFormat=51)
            wb.Close()
            excel.Application.Quit()

import openpyxl                                                                                                       
import re                                                                                                             
import traceback     
import os 
import pandas as pd
import numpy as np
#WIN32
import os #用于获取目标文件所在路径
import win32com
import win32com.client
import pandas as pd
import numpy as np

EvrFilePath =r"C:\Users\yongjiangao\ProgramEvr_02.xlsx"

##调试更改文件.
path = r"C:\Users\yongjiangao\QA_EVR\QA_EVR.xlsx"
# newPath=r"C:\Users\yongjiangao\QA_EVR\替换结果\700-014621-0X00 TEST2A EVR.xlsx"
# newPath2=r"C:\Users\yongjiangao\QA_EVR\替换结果\700-014619-0X00 TEST2A EVR.xlsx"
# Appname72_2=" QA Program"
# Appname72_1="Application Name:  "+"700-014621-0X00"+Appname72_2
# # Appname72_2=" QA Program"
# Appidentify92_2=" Rev:"
# Appidentify92_1="Application Name:  "+"700-014621-0X00(QA)Test2A.prg"+Appidentify92_2+"1"
# # Appidentify92_2=" Rev:"
# AppDes112_1="Designed Use :  "+"700-014621-0X00"
# AppCs204="6521"
# AppSpec212="Location of the test spec :"+r"L:\Common\Engg-Notice\EP Testing EN\Models\700-014621-0000\Test plan\700-014621-0X00 Validation Plan Rev 7.9.xlsx"

# #WIN32
# xlApp = win32com.client.Dispatch('Excel.Application')  # 打开word应用程序
# xlApp.Visible = 0  # 后台运行,不显示
# xlApp.DisplayAlerts = 0  # 不警告
# xlBook =xlApp.Workbooks.Open(path)
# xlSheet= xlBook.Worksheets('sheet1')
# xlSheet.Cells(7,2).Value=Appname72_1
# xlSheet.Cells(9,2).Value=Appidentify92_1
# xlSheet.Cells(11,2).Value=AppDes112_1
# xlSheet.Cells(20,4).Value=AppCs204
# xlSheet.Cells(21,2).Value=AppSpec212
# xlBook.SaveAs(newPath)

##调试读取Excel获取相应的信息并输出
EvaData=pd.DataFrame(pd.read_excel(EvrFilePath,sheet_name="ZJ2 (4)"))
# EvaData.info()
# print(EvaData.shape[0])
# print(len(EvaData))
# print(EvaData.head(3))
# print(EvaData.loc[0:5])

xlApp = win32com.client.Dispatch('Excel.Application')  # 打开word应用程序
xlApp.Visible = 0  # 后台运行,不显示
xlBook =xlApp.Workbooks.Open(path)
xlSheet= xlBook.Worksheets('sheet1')

# xlBook.SaveAs(newPath)
for row in range(EvaData.shape[0]):
    # Pathtemp=r("C:\Users\yongjiangao\QA_EVR\替换结果")
    newPath="C:\\Users\\yongjiangao\\QA_EVR\\替换结果\\"+str(EvaData['Model description'][row])+" "+str(EvaData['Flow2'][row])+" QA EVR.xlsx"
    print(EvaData['Model description'][row])
    print(EvaData['Flow2'][row])
    print(EvaData['StanderProgramName'][row])
    print(EvaData['SpecPath'][row])
    print(EvaData['TestProgram Checksum'][row])
    print(EvaData['Test Program Rev'][row])
    xlSheet= xlBook.Worksheets('sheet1')
    xlSheet.Cells(7,2).Value="Application Name:  "+str(EvaData['Model description'][row])+" QA Program"
    xlSheet.Cells(9,2).Value="Application Name:  "+str(EvaData['StanderProgramName'][row])+" Rev:"+str(EvaData['Test Program Rev'][row])
    xlSheet.Cells(11,2).Value="Designed Use :  "+str(EvaData['Model description'][row])
    xlSheet.Cells(20,4).Value=EvaData['TestProgram Checksum'][row]
    xlSheet.Cells(21,2).Value="Location of the test spec :"+str(EvaData['SpecPath'][row])
    xlBook.SaveAs(newPath)
    print("------")
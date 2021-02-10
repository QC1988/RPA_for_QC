#coding: UTF-8
import os
import shutil
import glob
import pandas as pd
import openpyxl
import xlwings as xw
import time
from pathlib import Path

# 0.从parameters.txt中读取参数
paras = []
with open(r'C:\Users\035203557\Desktop\kaizen_space\RPA\Create_RMA_List\parameters.txt', "r") as f:
    for line in f.readlines():
        line = line.strip('\n')  #去掉列表中每一个元素的换行符
        line = line.split("=")[1]
        paras.append(line)
# 0.1.接受输入的excel文件名称 get the address and name of the RMA_list
RMA_list_path = paras[0]
RMA_list_name = paras[1]
ITS_report_path = paras[2]
SBF_report_path = paras[3]
UNVISUALABLE_OF_EXCLE = int(paras[4])
print('RMA_list_path=%s'%RMA_list_path)
print('RMA_list_name=%s'%RMA_list_name)
print('ITS_report_path=%s'%ITS_report_path)
print('SBF_report_path=%s'%SBF_report_path)
print('If UNVISUALABEL_OF_EXCEL = 1, the windows of Excel will not be shown.')
print("UNVISUALABEL_OF_EXCEL=%s"%UNVISUALABLE_OF_EXCLE)
print('Parameters is set successfully.')
print("---------------------------------------")

# 1.找到名称为RMA的excel文件 find the RMA_list
# 1.1.合成文件地址及文件名
RMA_list_path_name = RMA_list_path + "\\" + RMA_list_name
print(RMA_list_path_name)
# 1.2.打开文件,保存文件路径
if os.path.exists(RMA_list_path_name):
    df = pd.DataFrame
    df = pd.read_excel(RMA_list_path + "\\" + RMA_list_name, engine='openpyxl')

    # 1.2.在指定目录（及子目录）下搜索文件 
    # # 2.将excel中特定列的值取出，存到list中
    # 2.1.特定RMA.No关键字,存到list中
    # 用for循环找到RMA关键字，然后存储那一列下方所有的RMA.No
    i_max = df.shape[0]
    j_max = df.shape[1]
    print("The matrix of the RAM_list is (%s, %s)"%(i_max,j_max))
    print("...")

    RMA_list = []
    for i in range(0,i_max):
        for j in range(0, j_max):
            if df.iat[i, j] == 'Serial #':
                i = i + 1
                for i in range(i, i_max):                    
                    if pd.isnull(df.iat[i, j]):                    
                        break
                    else:
                        RMA_list.append(df.iat[i, j])
                        i = i + 1               
            j = j + 1
        i = i + 1         
    print("The list of the RMA is as follow:")
    print(RMA_list)
    print("...")
    print("The number of the RMA is: %d"%len(RMA_list))
    print("...")
    # 2.3.加一个判定，是否只包含需要的数字

    # 3.循环，根据list中的值，将符合值的文件复制到一个地方
    # 3.1.添加搜索路径，到母目录就可以


    # 3.2.在服务器上找到所有的需要下载的report文件
    reports_path_name_server = []
    for n in RMA_list:
        temp_path = []
        # 包括子文件及正则表达式的项
        subfolder = '\\**\\'
        regEx = '*' + n + '*'
        # 在一个地址寻找属于ITS的报告
        if len(glob.glob(ITS_report_path+ subfolder + regEx + '.xlsx',recursive=True)) :
            temp_path_name = glob.glob(ITS_report_path+ subfolder + regEx + '.xlsx',recursive=True)
            print(temp_path_name)
            print("ITS_xlsx")
            reports_path_name_server.append(temp_path_name)
        elif len(glob.glob(ITS_report_path+ subfolder + regEx + '.xls',recursive=True)) :
            temp_path_name = glob.glob(ITS_report_path+ subfolder + regEx + '.xls',recursive=True)
            print(temp_path_name)
            print("ITS_xls")
            reports_path_name_server.append(temp_path_name)         
        else:
            # 在另一个地址寻找属于SBF的报告
            if len(glob.glob(SBF_report_path+ subfolder+ regEx + '.xlsx',recursive=True)):
                temp_path_name = glob.glob(SBF_report_path+ subfolder+ regEx + '.xlsx',recursive=True)
                print(temp_path_name)
                print("SBF_xlsx")
                reports_path_name_server.append(temp_path_name)
            elif len(glob.glob(SBF_report_path+ subfolder+ regEx + '.xls',recursive=True)):
                temp_path_name = glob.glob(SBF_report_path+ subfolder+ regEx + '.xls',recursive=True)
                print(temp_path_name)
                print("SBF_xls")
                reports_path_name_server.append(temp_path_name)
            else:
                print("%s doesn't be found." %n)
                # 将所有报告的地址及名字存到一个列表里
                # reports_path_name_server.append(temp_path_name)

    #             # 验证列表中没有相同名称
    # if len(temp_path_name) = i_max:
    #     print("All reports need to be inserted to RMA_list were located.")
    # else:
    #     print("Error. line 102.")

    # 3.3.用2中保存的列表，去依次复制文件到当前目录下
    print(len(RMA_list))
    print("--------------------------------")
    time.sleep(0.5)
    for m in range(0, len(RMA_list)):
        # print(reports_path[m][0])
        if reports_path_name_server[m][0][-1] == 'x':
            shutil.copyfile(reports_path_name_server[m][0],RMA_list_path+'\\'+ str(m + 1)+'.xlsx')   
        else:
            shutil.copyfile(reports_path_name_server[m][0],RMA_list_path+'\\'+ str(m + 1)+'.xls')
    print("All reports need to be inserted to RMA_list were download from server ")
    print('-----------------------------')

    # 4.将复制好的文件插入RMA中


    Report_No_Max = len(RMA_list)
    for n in range(1,Report_No_Max + 1):
        # print(str(n))
        path = Path(RMA_list_path+'\\'+ str(n)+'.xlsx')
        if path.exists():
            reports_path_name_local = RMA_list_path+'\\'+ str(n)+'.xlsx'
        else:
            reports_path_name_local = RMA_list_path+'\\'+ str(n)+'.xls'
        if UNVISUALABLE_OF_EXCLE:
            app = xw.App(visible=False)
            wb_target = app.books.open(RMA_list_path_name)
            print('%s is on processing'%reports_path_name_local)
            wb_source = app.books.open(reports_path_name_local)

        else:
            wb_target = xw.Book(RMA_list_path_name)
            wb_source = xw.Book(reports_path_name_local)

        ws_source = wb_source.sheets(1)
        ws_source.api.Copy(After=wb_target.sheets(n).api)
        wb_target.sheets[n].name = str(n)

        wb_target.save()
        time.sleep(0.5)
        # wb_source.save()
        wb_target.close()
        wb_source.close()
        print('%s was inserted in RMA_list.'%reports_path_name_local)
        #判断文件是否存在
        #插入的同时删除掉每个sheet的特定字符
        if(os.path.exists(reports_path_name_local)):
            os.remove(reports_path_name_local)
            print('%s was deleted.'%reports_path_name_local)
            print('-----------------------------')
        else:
            print("The file cann't be deleted because %s doesn't exist!"%reports_path_name_local)

else:
    print("The RMA_list doesn't exist. Please check the address and name of the RMA_list.")

print("RMA_list was created succefully.")
# # 可能存在的问题，如果NG不是第一个sheet？2.3处加一个数量相等判定
# 插入sheet时判定有无同名sheet，否则会出现错误
# #优点
# # 不在前台显示，先将excel文件下载到本地再进行操作，安全性高，

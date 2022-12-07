#!/usr/bin/python
# -*- coding: <utf-8> -*-

import os
import shutil
import glob
import pandas as pd
import openpyxl
import xlwings as xw
import time
from pathlib import Path
import datetime
import configparser

# define class to void the chars changed to lower chars
class myconf(configparser.ConfigParser):
    def __init__(self, defaults=None):
        configparser.ConfigParser.__init__(self, defaults=None)
    def optionxform(self, optionstr):
        return optionstr

def main():
    # 0.get paras from config.ini
    pwd = os.getcwd()
    father_path=os.path.abspath(os.path.dirname(pwd)+os.path.sep+".")
    config_file = os.path.join(father_path,'config.ini' )
    if not os.path.exists(config_file):
        print("Config file is not exist.")
        exit()
    else:
        config = myconf()
        config.read(config_file, encoding='utf-8')

    RMA_list_path = config['paras_for_RMA']['RMA_list_path']
    RMA_list_name = config['paras_for_RMA']["RMA_list_name"]
    ITS_report_path = config['paras_for_RMA']["ITS_report_path"]
    SBF_report_path = config['paras_for_RMA']["SBF_report_path"]
    UNVISUALABLE_OF_EXCLE = config['paras_for_RMA']["UNVISUALABLE_OF_EXCLE"] # 0.1.get the path  of the UNVISUALABLE_OF_EXCLE
    print(UNVISUALABLE_OF_EXCLE)
    Template_list_local_path = config['paras_for_RMA']["Template_list_local_path"] # 0.1.get the path  of the Template_list_local_path
    Template_list_name_for_EATON = config['paras_for_RMA']["Template_list_name_for_EATON"] # 0.1.get the path  of the Template_list_name_for_EATON
    Template_list_name_for_APD = config['paras_for_RMA']["Template_list_name_for_APD"] # 0.1.get the path  of the Template_list_name_for_APD
    # 0.1 ITS paras
    ITS_market_defect_list_online_path = config['paras_for_RMA']["ITS_market_defect_list_online_path"] # 0.1.get the path  of the ITS_market_defect_list_online_path
    ITS_market_defect_list_local_path = config['paras_for_RMA']['ITS_market_defect_list_local_path'] # 0.1.get the path  of the ITS_market_defect_list_local_path
    ITS_market_defect_list_name = config['paras_for_RMA']['ITS_market_defect_list_name'] # 0.1.get the path  of the ITS_market_defect_list_name

    # 0.2 SBF paras
    SBF_inspection_defect_list_online_path = config['paras_for_RMA']["SBF_inspection_defect_list_online_path"] # 0.1.get the path  of the SBF_inspection_defect_list_online_path
    SBF_inspection_defect_list_local_path = config['paras_for_RMA']['SBF_inspection_defect_list_local_path'] # 0.1.get the path  of the SBF_inspection_defect_list_local_path
    SBF_inspection_defect_list_name = config['paras_for_RMA']['SBF_inspection_defect_list_name'] # 0.1.get the path  of the SBF_inspection_defect_list_name

    RMA_list_path_name = RMA_list_path + "\\" + RMA_list_name

    Template_list_local_path_name_for_EATON = Template_list_local_path +  "\\"  + Template_list_name_for_EATON
    Template_list_local_path_name_for_APD = Template_list_local_path +  "\\" + Template_list_name_for_APD

    RMA_list_path_EATON_template_name = RMA_list_path + "\\"  + Template_list_name_for_EATON
    RMA_list_path_APD_template_name = RMA_list_path + "\\"  + Template_list_name_for_APD

    # print(SBF_inspection_defect_list_online_path)
    # print(SBF_inspection_defect_list_local_path)
    # print(SBF_inspection_defect_list_name)
    
    print("--------Version.1.3-----------------")
    print("--------1.Start preprocessing---------")
    print("Load parameters such as file name and address...")


    ################################################### Ver.1.2 update ##########################################################
    # 
    copy_comand_EATON_template = 'copy' + ' ' + '"' + Template_list_local_path_name_for_EATON  + '"'  +  ' ' + '"' + RMA_list_path +'"'
    copy_comand_APD_template= 'copy' + ' ' +  '"' + Template_list_local_path_name_for_APD + '"'   + ' ' + '"' + RMA_list_path +'"'


    if "-" in RMA_list_name:
        # copy EATON RMA template to RMA_list_path
        print("Copy EATON RMA list from template folder to work folder...")
        os.system(copy_comand_EATON_template)

        time.sleep(1)
        if os.path.isfile(RMA_list_path_EATON_template_name):
            os.rename(RMA_list_path_EATON_template_name, RMA_list_path_name)
        else:
            print("EATON template didn't be copyed to the work folder!")
    else:
        # copy APD RMA template to RMA_list_path
        print("Copy APD RMA list from template folder to work folder...")
        os.system(copy_comand_APD_template)
        time.sleep(1)
        if os.path.isfile(RMA_list_path_APD_template_name):
            os.rename(RMA_list_path_APD_template_name, RMA_list_path_name)
        else:
            print("APD template didn't be copyed to the work folder!")

    ##---------start ---------
    print("--------2.Start processing---------")

    ####################################################  ITS  ###################################################################################
    ## download the "ITS market defect list" into the Folder  Caution:if there is a space in the path, copy "d:\test abc\test.txt" "d:\t&est"
    ITS_market_defect_list_online_path_name = '"' + ITS_market_defect_list_online_path + '\\' + ITS_market_defect_list_name + '"'
    
    copy_comand_ITS = 'copy' + ' ' + ITS_market_defect_list_online_path_name  + ' ' + '"' + ITS_market_defect_list_local_path +'"'
    print( "Download the ITS_market_defect_list into the Folder..." )
    os.system(copy_comand_ITS)
    time.sleep(3)
    if os.path.exists(ITS_market_defect_list_local_path + '\\' + ITS_market_defect_list_name ):
        print("ITS_market_defect_list has be download successfully!")
    else:
        print("ITS_market_defect_list didn't be download!")
        print("Program has been shut down!")
        exit(0)


    ## open the "ITS market defect list"
    df = pd.DataFrame
    ITS_market_defect_list_local_path_name = ITS_market_defect_list_local_path + '\\' + ITS_market_defect_list_name
    # get all the sheets_ITS in dict
    # print(ITS_market_defect_list_local_path_name)

    df = pd.read_excel(ITS_market_defect_list_local_path_name,sheet_name=[0,1,2,3,4,5,6,7,8,9,10,11],header=11,nrows=500,engine='openpyxl')
    # print(df)
    # print(df[0])

    # print(len(df.keys()))

    # get the numbers of the sheets_ITS
    num_of_sheets_ITS = len(df.keys())

    ### get the list of the sheets_ITS
    sheets_ITS = pd.DataFrame
    # define the range of the data
        # change yy.mm.dd to yy/mm/dd
    for n in range(0, num_of_sheets_ITS):
        
        key_name = list(df.keys())[n]
        sheet_ITS = df[key_name]
        # print(key_name)
        # print(sheet_ITS)
    print("***************************")
    print("Process date:",end="")
    print(datetime.date.today())


    ## copy last three month sheets_ITS into a FrameData
    for n in range(0, num_of_sheets_ITS):
            
            key_name = list(df.keys())[n]
            sheet_ITS = df[key_name]
            
            # inspection_date_before_change = key_name.split(" ",1)[1]
            # inspection_date_after_change = inspection_date_before_change.replace('.', '/')

            # add one column yy.mm.dd before "型式" (BX)
            # sheet_ITS.rename(columns={'Unnamed: 0':'日付'}, inplace=True)
            # sheet_ITS['日付'] = inspection_date_after_change
            

            # delete the column which is not necessary
            # sheet_ITS = sheet_ITS.drop(columns=['検査\n不履行'])

            # delete the row which is not necessary
            # sheet_ITS = sheet_ITS.dropna(axis=0, subset=["型式"])

            # save the data(range) to DataFrame
            if n == 0:
                sheets_ITS = sheet_ITS
            else:
                sheets_ITS = pd.concat([sheets_ITS,sheet_ITS], axis=0)

    # print(sheets_ITS)


    ## filter the necessary info
    ## necessary info　型式、S/N、到着日、不具合内容、打診結果、連番、棚番、yasu連番
    sheets_ITS.dropna(axis=1,how='all')  
    sheets_ITS.drop(sheets_ITS.columns[26:], axis=1, inplace=True)


    sheets_ITS.drop(columns=['RMA入手','現品処理依頼','現品処理結果','写真','yasu連番','21処理','重要度','連番','受付日','客先','障害符号','不具合内容(英語)','原因符号\nLevel1','原因符号\nLevel2','原因符号\nLevel3','Problem Description in RMA request'],axis=1, inplace=True)


    ### RMA対象＝＝RMA、RMA要否＝＝要、RMA.NO＝＝RMA0030-P
    sheets_ITS = sheets_ITS[(sheets_ITS['処理判断(ITS)']=='RMA')&(sheets_ITS['RMA要否']=='要')]

    sheets_ITS.drop(columns=['RMA要否','処理判断(ITS)'],axis=1, inplace=True)

    order = ['棚番','到着日','機種','不具合内容','S/N','RMA要求','打診結果','処理']
    sheets_ITS = sheets_ITS[order]


    # print(sheets_ITS)

    if '-' in RMA_list_name:
        RMA_list_name_without_brackets = RMA_list_name.split("(", 1)[0]
        print("RMA list name:%s"%RMA_list_name_without_brackets)
        
    else:
        RMA_list_name_without_brackets = RMA_list_name.split(".", 1)[0]
        print("RMA list name:%s"%RMA_list_name_without_brackets)

    print("***************************")
    # exit()

    sheets_ITS = sheets_ITS[sheets_ITS['RMA要求']==RMA_list_name_without_brackets]
    sheets_ITS = sheets_ITS.rename(columns={'到着日':'発生日'})

    # print(sheets_ITS)
    '''
    # wirte dataFrame to a new excel file
    with pd.ExcelWriter(RMA_list_name) as writer:
        sheets_ITS.to_excel(writer,sheet_name='ITS')

    '''
    # time.sleep(3)

    ####################################################  SBF  #########################################################################
    ## download the "ITS market defect list" into the Folder  Caution:if there is a space in the path, copy "d:\test abc\test.txt" "d:\t&est"
    SBF_inspection_defect_list_online_path_name = '"' + SBF_inspection_defect_list_online_path  + '\\' + SBF_inspection_defect_list_name + '"'

    copy_comand_SBF = 'copy' + ' ' + SBF_inspection_defect_list_online_path_name  + ' ' + '"' + SBF_inspection_defect_list_local_path +'"'
    print( "Download the SBF_inspection_defect_list into the Folder...")

    os.system(copy_comand_SBF)

    time.sleep(3)
    if os.path.exists(SBF_inspection_defect_list_local_path + '\\' +SBF_inspection_defect_list_name ):
        print("SBF_inspection_defect_list has be download successfully!")
    else:
        print("SBF_inspection_defect_list didn't be download!")
        print("Program has been shut down!")
        exit(0)

    ## open the "ITS market defect list"
    df = pd.DataFrame
    SBF_inspection_defect_list_local_path_name = SBF_inspection_defect_list_local_path + '\\' +  SBF_inspection_defect_list_name
    # get all the sheets_SBF in dict
    # print(SBF_inspection_defect_list_local_path_name)

    df = pd.read_excel(SBF_inspection_defect_list_local_path_name,sheet_name='管理表',header=2,nrows=1500,engine='openpyxl')
    # print(df)

    sheets_SBF = df

    ## filter the necessary info
    ## necessary info　型式、S/N、到着日、不具合内容、打診結果、連番、棚番、yasu連番
    sheets_SBF.dropna(axis=1,how='all')  
    sheets_SBF.drop(sheets_SBF.columns[34:], axis=1, inplace=True)

    # print(sheets_SBF)

    sheets_SBF.drop(columns=['QA\n更新日','生管\n更新日','状況','品番','仕入先','対応予定','ＮＧ\n区分','台数', '#1','Omron→SBF\n作業依頼','#2','#3','SBF→生管\nコメント','箱\n納入日\nMM/DD','良品化\n指定\n納期\nMM/DD','納期\n回答\nMM/DD','作業\n完了\n連絡\nMM/DD','棚移動指示\nMM/DD'],axis=1, inplace=True)
    sheets_SBF.drop(columns=['移動先\n３QA/IT・５YU/SA/IT・\n良品化・RMA（廃棄予定）','代品入手','代品\n支給日\nMM/DD','3NOFから\n抜取の場合\nシリアル＃','良品化\n指定\n納期\nMM/DD/','納期\n回答\nMM/DD/','作業完了連絡\nMM/DD','棚移動指示\nSCM処置\nMM/DD','移動状態\n[QA確認棚移動］\n済or未','RMA＃\n取得確認'],axis=1, inplace=True)
    # print(sheets_SBF)


    ### RMA対象＝＝RMA、RMA要否＝＝要、RMA.NO＝＝RMA0030-P
    sheets_SBF = sheets_SBF[(sheets_SBF['棚']=='SBF')|(sheets_SBF['棚']=='SBA')]


    # print(sheets_SBF)

    # define PHP,TPE,BA2,APD according to the file name, RMA0029-P→PHP RMA0028-B→BA2 ...
    # EATON RMA0000-X()    APD RMA00000000
    if '-' in RMA_list_name:
        RMA_list_name_without_brackets = RMA_list_name.split("(", 1)[0]
        # print(RMA_list_name_without_brackets)
        
    else:
        RMA_list_name_without_brackets = RMA_list_name.split(".", 1)[0]
        # print(RMA_list_name_without_brackets)

    # exit()

    sheets_SBF = sheets_SBF[sheets_SBF['移動先\n３QA・３IT・５YU・５SA・\nRMA（廃棄予定）']==RMA_list_name_without_brackets]
    sheets_SBF = sheets_SBF.rename(columns={'棚':'棚番','受検日\n（発生日）':'発生日','機種名':'機種','不良詳細':'不具合内容','シリアル№':'S/N','移動先\n３QA・３IT・５YU・５SA・\nRMA（廃棄予定）':'RMA要求',})
        # print(sheets_SBF)


    sheets_ITS_SBF = pd.concat([sheets_ITS,sheets_SBF])


    # wirte dataFrame to a new excel file
    with pd.ExcelWriter('DATA_FOR_RMA.xlsx') as writer:
        sheets_ITS_SBF.to_excel(writer,sheet_name='DATA')

    DATA_FOR_RMA_local_path_name = RMA_list_path +  '\\' + 'DATA_FOR_RMA.xlsx'


    # 4.insert the copyed file into the RMA
    # Report_No_Max = len(RMA_list)
    if UNVISUALABLE_OF_EXCLE:
        app = xw.App(visible=False)
        wb_target = app.books.open(RMA_list_path_name)
        
        # wb_target = app.books.open(RMA_list_path_name)
        print('%s is on processing'%RMA_list_name)
        wb_source = app.books.open((RMA_list_path + '\\'+ 'DATA_FOR_RMA.xlsx'))

        # ws_source = wb_source.sheets(1)
        my_values = wb_source.sheets['DATA'].range("A1:I60").options(ndim=2).value
        wb_target.sheets['DATA'].range('A1:I60').value = my_values

        # ws_source.api.Copy(After=wb_target.sheets(1).api)
        # wb_target.sheets[n].name = str(n)

        wb_target.save()
        time.sleep(0.5)
        # wb_source.save()
        # wb_target.close()
        wb_source.close()
        print('DATA_SHEET was inserted in RMA_list.')
        # judge the file is exist
        # delete the specific charactor when insert
        # if(os.path.exists(reports_path_name_local)):
        #     os.remove(reports_path_name_local)
        #     print('%s was deleted.'%reports_path_name_local)
        #     print('-----------------------------')
        # else:
        #     print("The file cann't be deleted because %s doesn't exist!"%reports_path_name_local)

        wb_target.close()

    else:
        app = xw.App(visible=True)
        wb_target = app.books.open(RMA_list_path_name)
        
        # wb_target = app.books.open(RMA_list_path_name)
        print('%s is on processing'%RMA_list_name)
        wb_source = app.books.open((RMA_list_path + '\\'+ 'DATA_FOR_RMA.xlsx'))

        # ws_source = wb_source.sheets(1)
        my_values = wb_source.sheets['DATA'].range("A1:I60").options(ndim=2).value
        wb_target.sheets['DATA'].range('A1:I60').value = my_values


        # ws_source.api.Copy(After=wb_target.sheets(1).api)
        # wb_target.sheets[n].name = str(n)

        wb_target.save()
        time.sleep(0.5)
        # wb_source.save()
        # wb_target.close()
        wb_source.close()
        print('DATA_SHEET was inserted in RMA_list.')
        # judge the file is exist
        # delete the specific charactor when insert
        # if(os.path.exists(reports_path_name_local)):
        #     os.remove(reports_path_name_local)
        #     print('%s was deleted.'%reports_path_name_local)
        #     print('-----------------------------')
        # else:
        #     print("The file cann't be deleted because %s doesn't exist!"%reports_path_name_local)

        wb_target.close()

    '''
    else:    #!!!!!!!!!!!!!!!!!!!!!!!
        wb_target = xw.Book(RMA_list_path_name)
        for n in range(1, Report_No_Max + 1):
            # print(str(n))
            path = Path(RMA_list_path+'\\'+ str(n)+'.xlsx')
            if path.exists():
                reports_path_name_local = RMA_list_path+'\\'+ str(n)+'.xlsx'
            else:
                reports_path_name_local = RMA_list_path+'\\'+ str(n)+'.xls'

            wb_source = xw.Book(reports_path_name_local)

            ws_source = wb_source.sheets(1)
            ws_source.api.Copy(After=wb_target.sheets(n).api)
            wb_target.sheets[n].name = str(n)

            wb_target.save()
            time.sleep(0.5)
            # wb_source.save()
            # wb_target.close()
            wb_source.close()
            print('%s was inserted in RMA_list.'%reports_path_name_local)
            # judge the file is exist, delete the specific charactor when insert 
            if(os.path.exists(reports_path_name_local)):
                os.remove(reports_path_name_local)
                print('%s was deleted.'%reports_path_name_local)
                print('-----------------------------')
            else:
                print("The file cann't be deleted because %s doesn't exist!"%reports_path_name_local)

        wb_target.close()
    '''

    ## save the necessary info into the right RMA template
    ## delete the "ITS market defect list" in the Folder

    time.sleep(2)

    #############################################  insert ITS_reports ###################################


    # 1.Find the RMA_list
    # 1.1.combine the file name and path 
    RMA_list_path_name = RMA_list_path + "\\" + RMA_list_name
    # print(RMA_list_path_name)
    # 1.2.open the file, and save the path
    if os.path.exists(RMA_list_path_name):
        df = pd.DataFrame
        df = pd.read_excel(RMA_list_path + "\\" + RMA_list_name, engine='openpyxl')

        # 1.2.在指定目录（及子目录）下搜索文件 
        # # 2.将excel中特定列的值取出，存到list中
        # 2.1.特定RMA.No关键字,存到list中
        # 用for循环找到RMA关键字，然后存储那一列下方所有的RMA.No
        i_max = df.shape[0]
        j_max = df.shape[1]
        # print("The matrix of the RAM_list is (%s, %s)"%(i_max,j_max))
        print("...")

        RMA_list = []
        for i in range(0,i_max):
            for j in range(0, j_max):
                if (df.iat[i, j] == 'Serial #')|(df.iat[i, j] == 'Number'):
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
        print("The number of the RMA productions is: %d"%len(RMA_list))
        print("The files were found at the following path")
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
        if UNVISUALABLE_OF_EXCLE:
            app = xw.App(visible=False)
            wb_target = app.books.open(RMA_list_path_name)
            for n in range(1,Report_No_Max + 1):
                # print(str(n))
                path = Path(RMA_list_path+'\\'+ str(n)+'.xlsx')
                if path.exists():
                    reports_path_name_local = RMA_list_path+'\\'+ str(n)+'.xlsx'
                else:
                    reports_path_name_local = RMA_list_path+'\\'+ str(n)+'.xls'
                

                # wb_target = app.books.open(RMA_list_path_name)
                print('%s is on processing'%reports_path_name_local)
                wb_source = app.books.open(reports_path_name_local)

                ws_source = wb_source.sheets(1)
                ws_source.api.Copy(After=wb_target.sheets(n).api)
                wb_target.sheets[n].name = str(n)

                wb_target.save()
                time.sleep(0.5)
                # wb_source.save()
                # wb_target.close()
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

            wb_target.close()
        
        else:
            wb_target = xw.Book(RMA_list_path_name)
            app = xw.App(visible=True)
            for n in range(1,Report_No_Max + 1):
                # print(str(n))
                path = Path(RMA_list_path+'\\'+ str(n)+'.xlsx')
                if path.exists():
                    reports_path_name_local = RMA_list_path+'\\'+ str(n)+'.xlsx'
                else:
                    reports_path_name_local = RMA_list_path+'\\'+ str(n)+'.xls'

                wb_source = xw.Book(reports_path_name_local)

                ws_source = wb_source.sheets(1)
                ws_source.api.Copy(After=wb_target.sheets(n).api)
                wb_target.sheets[n].name = str(n)

                wb_target.save()
                time.sleep(0.5)
                # wb_source.save()
                # wb_target.close()
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

            wb_target.close()

    else:
        print("The RMA_list doesn't exist. Please check the address and name of the RMA_list.")

    print("RMA_list was created succefully.")
    print("--------3.End processing------------")

    print("Temporary files will be deleted...")

    # # 可能存在的问题，如果NG不是第一个sheet？2.3处加一个数量相等判定
    # # 插入sheet时判定有无同名sheet，否则会出现错误
    # # 优点 不在前台显示，先将excel文件下载到本地再进行操作，安全性高，


    if(os.path.exists(ITS_market_defect_list_local_path_name)):
        os.remove(ITS_market_defect_list_local_path_name)
        print('%s was deleted.'%ITS_market_defect_list_local_path_name)
        print('-----------------------------')
    else:
        print("The file cann't be deleted because %s doesn't exist!"%ITS_market_defect_list_local_path_name)

    if(os.path.exists(SBF_inspection_defect_list_local_path_name)):
        os.remove(SBF_inspection_defect_list_local_path_name)
        print('%s was deleted.'%SBF_inspection_defect_list_local_path_name)
        print('-----------------------------')
    else:
        print("The file cann't be deleted because %s doesn't exist!"%SBF_inspection_defect_list_local_path_name)


    if(os.path.exists(DATA_FOR_RMA_local_path_name)):
        os.remove(DATA_FOR_RMA_local_path_name)
        print('%s was deleted.'%DATA_FOR_RMA_local_path_name)
        print('-----------------------------')
    else:
        print("The file cann't be deleted because %s doesn't exist!"%DATA_FOR_RMA_local_path_name)


if __name__ == '__main__':
    main()

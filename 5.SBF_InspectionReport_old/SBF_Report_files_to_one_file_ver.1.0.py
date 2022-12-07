#coding: UTF-8
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
        config.read(config_file,encoding='utf-8')

    SBF_Report_path = config['paras_for_SBF_report']['SBF_Report_path']
    SBF_Report_name = config['paras_for_SBF_report']['SBF_Report_name']


    # get the number of the sheets

        # get the path_name of the SBF_Report
    SBF_Report_path_name = SBF_Report_path + "\\" + SBF_Report_name
        # open the target file

    if os.path.exists(SBF_Report_path_name):
        df = pd.DataFrame

        # get all the sheets in dict
        df = pd.read_excel(SBF_Report_path_name,sheet_name=None,header=3,nrows=50,engine='openpyxl')
        # print(df['検査完了 20.05.07'])

        # get the numbers of the sheets
        num_of_sheets = len(df.keys())


        sheets = pd.DataFrame

                
    # define the range of the data
            # change yy.mm.dd to yy/mm/dd
        for n in range(0, num_of_sheets):
            
            key_name = list(df.keys())[n]
            sheet = df[key_name]
            
            inspection_date_before_change = key_name.split(" ",1)[1]
            inspection_date_after_change = inspection_date_before_change.replace('.', '/')

            # add one column yy.mm.dd before "型式" (BX)
            sheet.rename(columns={'Unnamed: 0':'日付'}, inplace=True)
            sheet['日付'] = inspection_date_after_change
            

            # delete the column which is not necessary
            sheet = sheet.drop(columns=['検査\n不履行','化粧箱'])

            # delete the row which is not necessary
            sheet = sheet.dropna(axis=0, subset=["型式"])

            # save the data(range) to DataFrame
            if n == 0:
                sheets = sheet
            else:
                sheets = pd.concat([sheets,sheet], axis=0)
        sheets.rename(columns={'検査完了':'ロット数',"Q'ty":'OK','化粧箱NGOK':'箱','付属品\n(Attachment)':'付属品','重複NG内容　等':'詳細(外観不良と機能不良)'},inplace=True)
        order = ['日付','型式','管理No.','ロット数','OK','計',"機能及び外観\n(Function)",'付属品','箱','詳細(外観不良と機能不良)']
        sheets = sheets[order]

        # print(sheets)

    # repeat the cycle until all the sheets has be saved to DataFrame

    # wirte dataFrame to a new excel file
        with pd.ExcelWriter('SBF_inspection_record_summary.xlsx') as writer:
            sheets.to_excel(writer)
        print('SBF_inspection_record_summary was created!')
    else:
        print("{}見つかりませんでした！".format(SBF_Report_name))


if __name__ == '__main__':
    main()

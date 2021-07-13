# importing the libraries
#coding: UTF-8
import cv2 as cv
import os
import glob
from ast import literal_eval
# from sys import flags
# import pytesseract
import numpy as np
import pandas as pd
from PIL import Image
import pyocr
import re
import itertools

pyocr.tesseract.TESSERACT_CMD = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
# pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files (x86)\Tesseract-OCR\\tesseract.exe'
SHOW_PHOTO = False
# SHOW_PHOTO = True
tools = pyocr.get_available_tools()
tool = tools[0]
norm_size_x = 960
norm_size_y = 1280
list_KP_NO = []
list_pro_num = []

print("")
print("--------紙媒体からデータ抽出__Version.1.0--------")
print("前回登録したROIで実行しますか？")
print("1.はい。前回登録したROIを利用します。")
print("2.いいえ。新たなROIを作成します。")
print("数字を入力ください。")
judge_paras = input(">")
# print(judge_paras)

if judge_paras=="1":
    print("")
    print("--------前回登録したROIで処理します--------")
    paras_dict = {}
    with open('paras_for_OCR.txt', "r") as f:
        for line in f.readlines():
            line = line.strip('\n')  #remove every \n in the list
            if "=" in line :
                before_equal = line.split("=", 1)[0] # save the paras into a dict
                after_equal = line.split("=", 1)[1]  # Caution: without space in the end of the txt
                paras_dict[before_equal] = after_equal
            else:
                pass
            # print(paras_dict.items())

    if "NUM_OF_ROI" in paras_dict:
        NUM_OF_ROI = int(paras_dict["NUM_OF_ROI"]) # 0.1.get the path  of the ITS_market_defect_list_online_path
        
    else:
        print("NUM_OF_ROI can't be found!") 
        exit()
    if "STORAGE_R" in paras_dict:
        STORAGE_R = literal_eval(paras_dict["STORAGE_R"])
        # print(STORAGE_R) # 0.1.get the path  of the ITS_market_defect_list_online_path
    else:
        print("STORAGE_R can't be found!")
        exit()


elif judge_paras=="2":
    print("")
    print("--------ROIを選択する方法--------")
    print("ROI区域を選択してください。")
    print("Enterで確定してください。")
    print("Escで窓口を閉じ、次に進んでください。")
    print("--------------------------------")
    print("")
    print("抽出するデータ種類の数を入力ください。1~5")
    NUM_OF_ROI = int(input(">"))
    if NUM_OF_ROI>1:
        STORAGE_R = [0 for x in range(0,NUM_OF_ROI*4)]
        path = os.getcwd()
        regEx = '*'
        all_jpg_path_name = glob.glob(path +'\\'+ regEx +'.jpg', recursive=True)
        one_jpg_path_name = all_jpg_path_name[0]
        if os.path.exists(one_jpg_path_name):
            for i in range(1, NUM_OF_ROI+1):
                # print("Please select the zone!")
                src = cv.imdecode(np.fromfile(one_jpg_path_name, dtype=np.uint8), -1)
                # src = cv.imread(one_jpg_path_name)
                src = cv.resize(src, dsize=None, fx=0.5,fy=0.5)
                # print("鼠标选择ROI,然后点击 enter键")
                r = cv.selectROI('input', src, False)  # ,返回 (x_min, y_min, w, h)
                # print("ROI",r)
                # roi区域
                roi = src[int(r[1]):int(r[1]+r[3]), int(r[0]):int(r[0]+r[2])]

                # print("ROI",roi)
                k = cv.waitKey(0) & 0xFF
                if k == 27: # 按esc 键即可退出
                    cv.destroyAllWindows()
                for j in range(4):
                    index =(i-1)*4+j
                    STORAGE_R[index]= int(r[j])
                    # print("((i-1)*4+j)=%d"%((i-1)*4+j))
                    # print("j=%d"%j)
        f1 = open('paras_for_OCR.txt','w')
        f1.write("NUM_OF_ROI=")
        f1.write(str(NUM_OF_ROI))
        f1.write('\n')
        f1.write("STORAGE_R=")
        f1.write(str(STORAGE_R))
        # fi.write()
        f1.close() 
        # exit()
    else:
        print("1~5を入力する必要があります。")
        exit()
else:
    print("正しく入力されていません。")
    exit()


path = os.getcwd()
regEx = '*'
all_jpg_path_name = glob.glob(path +'\\'+ regEx +'.jpg', recursive=True)
one_jpg_path_name = all_jpg_path_name[0]
names = locals()
counts = 1
print("これから処理する写真の数が %dです"%len(all_jpg_path_name))
print("写真ごとに抽出するデータ種類の数が %d です"%NUM_OF_ROI)
print("")
print("--------処理開始--------")
if os.path.exists(one_jpg_path_name):
    for jpg_path_name in all_jpg_path_name:
        print("写真 %d"%counts)
        for k in range(NUM_OF_ROI):
            if  jpg_path_name == one_jpg_path_name:
                names['x%s'%k] = []
            im = cv.imdecode(np.fromfile(jpg_path_name, dtype=np.uint8), -1)
            # im = cv.imread((jpg_path_name)) #读取图片
            im = im[STORAGE_R[1+k*4]*2:(STORAGE_R[1+k*4]+STORAGE_R[3+k*4])*2,STORAGE_R[0+k*4]*2:(STORAGE_R[0+k*4]+STORAGE_R[2+k*4])*2]
            # im = im[int(r[1])*2:int(r[1]+r[3])*2, int(r[0])*2:int(r[0]+r[2])*2]
            # print('-----------------')
            # print(type(im))
            # cv.imshow('im',im)
            # cv.waitKey(0)
            
            if SHOW_PHOTO:
                im_res = cv.resize(im, dsize=None,fx=1,fy=1)
                cv.imshow('im_res', im_res)
            # print(im.shape)
        
            im_gray = cv.cvtColor(im, cv.COLOR_BGR2GRAY) #转换为灰度图
            if SHOW_PHOTO:
                im_gray_res = cv.resize(im_gray,dsize=None,fx=1,fy=1)
                cv.imshow('im_gray_res',im_gray_res)
            # print(im_gray.shape)
        
            im_gray_bright =np.uint8(np.clip((cv.add(1.5*im_gray, 0)), 0, 255))
            if SHOW_PHOTO:
                im_gray_bright_res = cv.resize(im_gray_bright, dsize=None,fx=1,fy=1)
                cv.imshow('im_gray_brignt_res',im_gray_bright_res)
            # print(im_gray_bright.shape)
        
            '''
            retval, im_gray_bright_bit = cv.threshold(im_gray_bright, 127, 255, cv.THRESH_BINARY) 
            if SHOW_PHOTO:
                im_gray_bright_bit_res = cv.resize(im_gray_bright_bit, dsize=None,fx=1,fy=1)
                cv.imshow('im_gray_bright_bit_res',im_gray_bright_bit_res)
            print(im_gray_bright_bit.shape)
        
            # tmp = np.hstack((im_res, im_gray_res, im_gray_bright_res, im_gray_bright_bit_res))
            # cv.imshow('image', tmp)
            img = im_gray_bright_bit
        
            height,width=img.shape
            dst1=np.zeros((height,width,1),np.uint8)
            for i in range(height):
                for j in range(width):
                    dst1[i,j]=255-img[i,j]
        
            if SHOW_PHOTO:
                dst1_res = cv.resize(dst1, dsize=None,fx=1,fy=1)
                cv.imshow("dst1_res", dst1_res)
            # 形态学计算，开操作1
            kernel1 = cv.getStructuringElement(cv.MORPH_RECT, (4, 3))
            res_open1 = cv.morphologyEx(dst1, cv.MORPH_OPEN, kernel1)
            if SHOW_PHOTO:
                res_open1_res = cv.resize(res_open1, dsize=None,fx=1,fy=1)
                cv.imshow("res_open1_res", res_open1_res)
        
            # 形态学计算，开操作2
            kernel2=cv.getStructuringElement(cv.MORPH_RECT,(3,4))
            res_open2=cv.morphologyEx(res_open1,cv.MORPH_OPEN,kernel2) #开操作
            if SHOW_PHOTO:
                res_open2_res = cv.resize(res_open2,dsize=None,fx=1,fy=1)    
                cv.imshow("res_open2_res",res_open2_res)  # 显示图片
        
            # 形态学计算，闭操作
            kernel3=cv.getStructuringElement(cv.MORPH_RECT,(5,5))
            res_close1=cv.morphologyEx(res_open2,cv.MORPH_CLOSE,kernel3) # 闭操作
            if SHOW_PHOTO:
                res_close1_res = cv.resize(res_close1, dsize=None,fx=1,fy=1)
                cv.imshow("res_close1_res", res_close1_res)
        
            img = res_close1
            height,width=img.shape
            dst2=np.zeros((height,width,1),np.uint8)
            for i in range(height):
                for j in range(width):
                    dst2[i,j]=255-img[i,j]
        
            if SHOW_PHOTO:
                dst2_res = cv.resize(dst2, dsize=None,fx=1,fy=1)
                cv.imshow("dst2_res", dst2_res)
        
            img2 = dst2
            '''
        
            result1 = tool.image_to_string(Image.fromarray(im_gray_bright), lang='jpn+eng', builder=pyocr.builders.TextBuilder(tesseract_layout=11))
            result2 = tool.image_to_string(Image.fromarray(im_gray_bright), lang='jpn+eng', builder=pyocr.builders.WordBoxBuilder(tesseract_layout=6))
            # print(result1)
            # print(result2)
            # cv.imwrite('output.jpg',image)
        
            result3 = np.array(result2)
            for box in result3:
                cv.rectangle(im, box.position[0], box.position[1], 2, 4)
            # cv.imwrite('output.png',im)

            if SHOW_PHOTO:
                im_out = cv.imread('output.png')
                im_out_res = cv.resize(im_out, dsize=None,fx=1,fy=1)
                cv.imshow('im_out_res',im_out_res)
                cv.waitKey(0)
        
            # verification
            # KP-NO  XX-XXXXX,  two-numbers, KP-NO before
            # 品名 
        
            # print(result1)
            # result1.replace('\\n', ',',flags=4)
            print("データ %d  "%k,end="")
            # print(" %d"%k)
            if k == 0 :
                tmp_1 = re.search('(\w\w-\d\d\d\d\d)', result1) #去掉换行符, flags=re.DOTALL   
                if tmp_1 != None:
                # if len(tmp_1):
                    print(tmp_1.group(1))
                    names['x%s'%k].append(str(tmp_1.group(1)))
                    # print(x0)
            if k == 1:
                tmp_2 = re.search('(\d\d\d\d\d\d\d-\d)', result1) 
                # print(tmp_2)
                if tmp_2 != None:
                    print(tmp_2.group(1))
                    names['x%s'%k].append(str(tmp_2.group(1)))
                else:
                    print("-")
                    names['x%s'%k].append("-")
                    # print(x1)
            if k == 2:
                tmp_3 = re.search('(abc)', result1, flags=re.DOTALL) 
                # print(tmp_2)
                if tmp_3 != None:
                    print(tmp_3.group(1))
                    names['x%s'%k].append(str(tmp_3.group(1)))
                else:
                    print("-")
                    names['x%s'%k].append("-")
                    # print(x1)
            if k == 3:
                tmp_4 = re.search('(abc)', result1) 
                # print(tmp_2)
                if tmp_4 != None:
                    print(tmp_4.group(1))
                    names['x%s'%k].append(str(tmp_4.group(1)))
                else:
                    print("-")
                    names['x%s'%k].append("-")
                    # print(x1)
            if k == 4:
                tmp_5 = re.search('(abc)', result1) 
                # print(tmp_2)
                if tmp_5 != None:
                    print(tmp_5.group(1))
                    names['x%s'%k].append(str(tmp_5.group(1)))
                else:
                    print("-")
                    names['x%s'%k].append("-")
                    # print(x1)

            
        counts += 1
    # print(x0)
    # print(x1)
    if NUM_OF_ROI==1:
        df = pd.DataFrame({'x0':x0})
    if NUM_OF_ROI==2:
        df = pd.DataFrame({'x0':x0, 'x1':x1})
    if NUM_OF_ROI==3:
        df = pd.DataFrame({'x0':x0, 'x1':x1, 'x2':x2})
    if NUM_OF_ROI==4:
        df = pd.DataFrame({'x0':x0, 'x1':x1,'x2':x2, 'x3':x3})
    if NUM_OF_ROI==5:
        df = pd.DataFrame({'x0':x0, 'x1':x1,'x2':x2, 'x3':x3,'x4':x4})


    # df = pd.DataFrame(x1)
    print("")
    print("csvファイルに書き込み内容は以下となります。")
    print(df)
    df.to_csv('Data_from_photos.csv')

else:
    print("There is no *.jpg file!")
print("")
print("--------処理完了--------")
cv.destroyAllWindows()

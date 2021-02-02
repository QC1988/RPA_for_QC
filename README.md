# RPA_for_QC  
a series of programs to develop the performance for QC  

Step 1  
create models for every function  

1.Model to join excel files to one
  OK
----------------------------------------------
2.Model to make graphs to explain the data on excel file  
3.Model to send emails automatically according to messages  
4.Model to find the excel files appointed and copy to the folder_path  


5.Model to insert excel files to one excel file as new sheets   
  import os
import shutil
import glob


# # 1.找到名称为RMA的excel文件
# 1.1.接受输入的excel文件名称
# 1.2.在指定目录（及子目录）下搜索文件
# 1.3.打开文件
# 1.4.保存文件路径

# # 2.将excel中特定列的值取出，存到list中
# 2.1.特定RMA.No关键字
# 2.2.选定RMA.No下面的数字，存到list中
# 2.3.加一个判定，是否只包含需要的数字

# 3.循环，根据list中的值，将符合值的文件复制到一个地方
# 4.将复制好的文件插入RMA中，并编号1,2，3.。。（也可以用RMA里面的编号）
# 5.插入的同时删除掉每个sheet的特定字符


file_name = input("Please input the file name:")
# print(file_name)
folder_path = 'C:\\Users\\Qichang Ql\\Desktop\\target'
file_path = folder_path + "\\" + file_name
print(file_path)
# print(folder_path)
# print(glob.glob('**/*.xlsx', recursive=True))
    # shutil.copy()
    
    
-------------------------------------------------------    

step 2  
integrate models to make flow under RPA  
1. 

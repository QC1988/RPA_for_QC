# delete the NG reports before today
#!/usr/bin/python
# -*- coding: <utf-8> -*-


import datetime
import os
import glob
import re
from typing import Pattern


year = datetime.date.today().year
month = datetime.date.today().month
if month <=9:
	month = "0" + str(month)
day = datetime.date.today().day
date_today = str(year) + str(month) + str(day)

print("Today is %s."%date_today)


local_path = "C:\\Users\\035203557\\Desktop\\kaizen_space\\RPA\\6.NEP_UpdateData"
file_name = "OMRON UPS INSPECTION REPORT-"
reg = "*"
local_path_files_name = local_path + "\\" + file_name + reg
NEP_inspection_files_path_name = glob.glob(local_path_files_name, recursive=True)


local_path = "C:\\Users\\035203557\\Desktop\\kaizen_space\\RPA\\6.NEP_UpdateData"
file_name = "NGReport_"
reg = "*"
local_path_files_name = local_path + "\\" + file_name + reg
NEP_NG_files_path_name = glob.glob(local_path_files_name, recursive=True)


# print(NEP_inspection_files_path_name)


yearmonthday = []


for i in NEP_inspection_files_path_name:
	pattern1 = re.compile(r'REPORT-\d\d\d\d\d\d\d\d')
	pattern2 = re.compile(r'\d\d\d\d\d\d\d\d')
	tmp1 = pattern1.search(i)
	if tmp1:
		tmp2 = pattern2.search(tmp1.group())
		
	for j in NEP_NG_files_path_name:
		if tmp2.group() in j:
			print("%s will not be delete."%j)
			continue
		else:
			print("%s will be delete"%j)
			os.remove(j)

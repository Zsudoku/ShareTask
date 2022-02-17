import pyautogui
import xlrd
import os
import pandas as pd
import re
import sys

def ChangeFlieName():
    rbData = xlrd.open_workbook("data.xls")
    wsData = rbData.sheet_by_index(0)
    roadLis = [] #存放路径
    daysLis = [] #存放天数
    floatRoadLis = [] #流通股存放路径
    sharkesNum = wsData.nrows-1
    oldFile_Name = "1"
    newFileName = "2"
    for i in range (1,sharkesNum+1):
        roadLis.append(wsData.cell(i,3).value)
        daysLis.append(wsData.cell(i,4).value)
        j = 1
        #while j < daysLis[i-1] + 1:
        oldFile_Name = str(roadLis[i-1]) + "\\" + str(j) + ".xls"
        try:
            rbData = pd.read_csv(oldFile_Name,encoding='gb18030')
        except FileNotFoundError:
            print("不存在文件",oldFile_Name)
            continue
        string = str(rbData.tail(0))
        lis = []
        lis.append(re.findall(r"\d+\.?\d*",string))
        lis = sum(lis,[])
        print(lis[0]+'_'+lis[1])
        newFileName = roadLis[i-1] + "\\" + lis[0]+'_'+lis[1] + ".xls"
        try:
            os.rename(oldFile_Name,newFileName)
        except FileExistsError:
            print("已存在文件",newFileName)
        j += 1 

if __name__ == '__main__':
    b = pyautogui.confirm(text='要开始修改文件名程序么？', title='请求框', buttons=['开始','取消'])
    if(b == '取消'):
        pyautogui.alert(text='程序运行结束！', title='请求框', button='OK')
        sys.exit()
    print('程序开始')
    #批量修改文件名
    ChangeFlieName()
    pyautogui.alert(text='程序运行结束！', title='请求框', button='OK')
#
# Created by Zsudoku on 2022/1/26.
#
import pyautogui
import time
import xlrd
import xlwt
import pyperclip
from xlutils.copy import copy
from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import os
import pandas as pd
import re
import sys

#去重
def delList(L):
    L1 = []
    for i in L:
        if i not in L1:
            L1.append(i)
    return L1

def getFloatShares(shares_name,load_name,shares_code):
    data_lis = []
    stock_lis = []
    data_lis2 = []
    stock_lis2 = []
    load_name = load_name + "\\流通股\\"
    stockName = shares_name
    stockCode = shares_code
    if(type(stockCode)== float):
        stockCode = int(stockCode)
    opt = Options()
    opt.add_experimental_option("excludeSwitches",['enable-automation','enable-logging'])
    opt.add_argument('log-level=3')
    opt.add_argument("--headless")
    opt.add_argument("--disbale-gpu")
    web = Chrome(options=opt)
    #web = Chrome()
    time.sleep(5)
    web.get("https://so.eastmoney.com/web/s?keyword=")
    time.sleep(5)
    #搜索
    web.find_element_by_xpath('//*[@id="search_key"]').send_keys(stockCode,Keys.ENTER)
    time.sleep(2)

    try:
        web.find_element_by_class_name("exstock_t_l") .click()
    except BaseException as msg:
        print("该股票不符条件",stockName)
        return
    else:
        time.sleep(2)
    #web.find_element_by_xpath('//*[@id="app"]/div[3]/div[1]/div[3]/div[1]/a[1]').click()
    #time.sleep(2)

    web.switch_to.window(web.window_handles[-1])

    #财务指标
    try:
        web.find_element_by_xpath('//*[@id="zjl_box"]/div[1]/ul/li[2]/h3/a').click()
    except BaseException as msg:
        print("该股票不符条件",stockName)
        return
    else:
        time.sleep(2)


    time.sleep(5)
    web.switch_to.window(web.window_handles[-1])
    time.sleep(1)
    #股本结构
    try:
        web.find_element_by_xpath('//*[@id="CapitalStockStructure"]/a').click() 
    except BaseException as msg:
        return -1
    #web.find_element_by_xpath('//*[@id="CapitalStockStructure"]/a').click() 
    time.sleep(5)

    #日期
    try:
        data= web.find_element_by_xpath('//*[@id="lngbbd_Table"]/tbody/tr[1]').text  
    except BaseException as msg:
        return -1
    data_lis.append(data)

    time.sleep(1)

    #流通股
    floatStock = web.find_element_by_xpath('//*[@id="lngbbd_Table"]/tbody/tr[19]').text  
    stock_lis.append(floatStock)
    time.sleep(1)

    for i in range(0,50):
        try:
            web.find_element_by_xpath('//*[@id="Table2Right"]/img').click()
            time.sleep(1)
        except BaseException as msg:
            break
        else:
            #日期 
            try:
                data= web.find_element_by_xpath('//*[@id="lngbbd_Table"]/tbody/tr[1]').text  
                data_lis.append(data)
                time.sleep(1)
            except BaseException as msg:
                break
            #流通股
            try:
                floatStock = web.find_element_by_xpath('//*[@id="lngbbd_Table"]/tbody/tr[19]').text  
                stock_lis.append(floatStock)
                time.sleep(1)
            except BaseException as msg:
                break
            stock_lis.append(floatStock)
            time.sleep(1)
            
        time.sleep(1)
    print("结束:",shares_name)
    if(os.path.exists(load_name) == False):
        os.makedirs(r"%s"%(load_name))
    fileName= load_name + stockName + '.xls'
    if(os.path.isfile(fileName)==False):
        workbook = xlwt.Workbook(encoding='utf-8')       #新建工作簿
        sheet1 = workbook.add_sheet("sheet1")          #新建sheet
        workbook.save(r"%s"%(fileName))   #保存

    rb = xlrd.open_workbook(fileName)    #打开.xls文件
    wb = copy(rb)                          #利用xlutils.copy下的copy函数复制
    ws = wb.get_sheet(0)  #获取表单0

    for i in range(len(data_lis)):
        data_lis2.append(data_lis[i].split())
    for i in range(len(stock_lis)):
        stock_lis2.append(stock_lis[i].split())
    for i in range(len(data_lis2)):
        try:
            del data_lis2[i][0]
        except BaseException as msg:
            break
    for i in range(len(stock_lis2)):
        try:
            del stock_lis2[i][0]
        except BaseException as msg:
            break
    data_lis2 = sum(data_lis2, [])#二维数组变一维数组
    stock_lis2 = sum(stock_lis2, [])
    # data_lis2 = delList(data_lis2)
    # stock_lis2 = delList(stock_lis2)
    print(data_lis2)
    print(stock_lis2)
    x = 0
    y = 0

    ws.write(0,0,"日期")
    ws.write(0,1,"流通股数量，单位：万股")

    x = 0
    y = 1

    for i in range(len(data_lis2)):
        ws.write(y,x,data_lis2[i])
        y += 1

    x = 1
    y = 1

    for i in range(len(stock_lis2)):
        ws.write(y,x,stock_lis2[i])
        y += 1
    wb.save(fileName) 

    data = pd.DataFrame(pd.read_excel(fileName))
    no_re_row = data.drop_duplicates()
    no_re_row.to_excel(fileName)
    time.sleep(1)
    web.quit()
    return 1


def FloatShare():
    rbData = xlrd.open_workbook("data.xls")
    wsData = rbData.sheet_by_index(0)
    codeLis = [] #存放股票代码
    nameLis = [] #存放股票搜索代码
    roadLis = [] #存放路径
    sharkesNum = wsData.nrows-1
    print("股票个数",sharkesNum)

    for i in range (1,sharkesNum+1):
        codeLis.append(wsData.cell(i,2).value)
        nameLis.append(wsData.cell(i,1).value)
        roadLis.append(wsData.cell(i,5).value)
        try:
            os.remove(roadLis[0] + "流通股\\" + codeLis[i-1] + ".xls")
        except BaseException as msg:
            print(roadLis[0] + "流通股\\" + codeLis[i-1] + ".xls")
    

    for i in range(len(codeLis)):
        while True:
            if(getFloatShares(codeLis[i],roadLis[i],nameLis[i]) == -1):
                getFloatShares(codeLis[i],roadLis[i],nameLis[i])
            elif(getFloatShares(codeLis[i],roadLis[i],nameLis[i]) == 1):
                break
        time.sleep(2)
    

if __name__ == '__main__':
    b = pyautogui.confirm(text='要开始运行爬取流通股程序么？', title='请求框', buttons=['开始','取消'])
    if(b == '取消'):
        pyautogui.alert(text='程序运行结束！', title='请求框', button='OK')
        sys.exit()
    print('程序开始')
    
    FloatShare()
    
    pyautogui.alert(text='程序运行结束！', title='请求框', button='OK')

    

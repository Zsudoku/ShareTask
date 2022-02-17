#
# Created by Zsudoku on 2022/1/14.
#
import pyautogui
import time
import xlrd
import pyperclip
from xlutils.copy import copy
import os
import sys
import shutil
#定义鼠标事件1
def mouseClick(clickTimes,lOrR,img,reTry):
    ti = 1
    while True:
        location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
        if location is not None:
            pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.15,duration=0.1,button=lOrR)
            return 1
            #break
        if(ti>20):
            return -1
        print("未找到匹配图片,0.1秒后重试 %s"%(img))
        ti += 1
        #time.sleep(0.1)
#判断前复权
def mouseClick2(clickTimes,lOrR,img,reTry):
    ii = 1
    while True:
        location=pyautogui.locateCenterOnScreen(img,confidence=0.95)
        if location is not None:
            print("更换为前复权")
            pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.15,duration=0.1,button=lOrR)
            time.sleep(0.5)
            break
        else:
            print("已是前复权")
            pyautogui.press('esc')
            time.sleep(0.5)
            break
#定义鼠标事件3
def mouseClick3(clickTimes,lOrR,img,reTry):
    if reTry == 1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.7)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.15,duration=0.1,button=lOrR)
                break
            print("未找到匹配图片,0.1秒后重试 %s"%(img))
def writeFile():
    rbData = xlrd.open_workbook("data.xls")
    wsData = rbData.sheet_by_index(0)

    rb = xlrd.open_workbook("shares/cmd.xls")    #打开.xls文件
    wb = copy(rb)                          #利用xlutils.copy下的copy函数复制
    ws = wb.get_sheet(0)  #获取表单0

    codeLis = [] #存放股票名称
    daysLis = [] #存放获取股票的天数
    roadLis = [] #存放路径
    numLis = [] #存放数字代码
    sharkesNum = wsData.nrows-1
    print("股票个数",sharkesNum)
    for i in range (1,sharkesNum):
        numLis.append(wsData.cell(i,1).value)
        codeLis.append(wsData.cell(i,2).value)
        daysLis.append(wsData.cell(i,4).value)
        roadLis.append(wsData.cell(i,3).value)
    ws.write(3,1,numLis[0])
    ws.write(1,1,daysLis[0])
    ws.write(24,1,roadLis[0])
    # ws.write(39,1,roadLis[0])
    # ws.write(50,1,roadLis[0])
    wb.save('shares/cmd.xls') 

#获取粘贴板里的内容   
def getCopyTxet():
    #os.system("echo off | clip")
    pyautogui.hotkey('ctrl','c')
    time.sleep(0.2)
    copytxet = pyperclip.paste()
    #os.system("echo off | clip")
    time.sleep(0.2)
    return str(copytxet)
#任务
def mainWork():
    imgResult = 0
    j = 1
    rbData = xlrd.open_workbook("data.xls")
    wsData = rbData.sheet_by_index(0)
    codeLis = [] #存放股票代码
    daysLis = [] #存放获取股票的天数
    roadLis = [] #存放路径
    numLis = [] #存放数字代码
    shareFileName = "1"
    sharkesNum = wsData.nrows-1
    path = '1'
    codeName = 0 
    print("股票个数",sharkesNum)

    for i in range (1,sharkesNum+1):
        numLis.append(wsData.cell(i,1).value)
        codeLis.append(wsData.cell(i,2).value)
        daysLis.append(wsData.cell(i,4).value)
        roadLis.append(wsData.cell(i,3).value)
        #numLis[i-1] = int(numLis[i-1])
    print("进入j")
    while j<sharkesNum+1:
        rb = xlrd.open_workbook("shares/cmd.xls")    #打开.xls文件
        wb = copy(rb)                          #利用xlutils.copy下的copy函数复制
        ws = wb.get_sheet(0)  #获取表单0
        ws.write(3,1,numLis[j-1])
        ws.write(1,1,daysLis[j-1])
        ws.write(27,1,roadLis[j-1])
        ws.write(29,1,roadLis[j-1])
        wb.save('shares/cmd.xls') 
        wb = xlrd.open_workbook("shares/cmd.xls")
        days = 0
        sheet1 = wb.sheet_by_index(0)
        img = sheet1
        i = 1
        downFileName = 1
        pageUpNum = 0
        time.sleep(0.5)
        pyautogui.press('esc')
        time.sleep(0.5)
        pyautogui.press('esc')
        time.sleep(0.5)
        pyautogui.press('esc')
        time.sleep(0.5)
        pyautogui.press('esc')
        time.sleep(0.5)
        pyautogui.press('esc')
        while i < sheet1.nrows:
            #取本行指令的操作类型
            cmdType = sheet1.row(i)[0]
            if cmdType.value == 1.0:
                #取图片名称
                img = sheet1.row(i)[1].value
                reTry = 1
                imgResult = mouseClick(1,"left",img,reTry)
                if(imgResult == -1):
                    j -= 1
                    break
                time.sleep(0.1)
                #print("单击左键",img)
            #2代表双击左键
            elif cmdType.value == 2.0:
                #取图片名称
                img = sheet1.row(i)[1].value
                #取重试次数
                reTry = 1
                mouseClick(2,"left",img,reTry)
                #print("双击左键",img)
            #3代表右键
            elif cmdType.value == 3.0:
                #取图片名称
                img = sheet1.row(i)[1].value
                #取重试次数
                reTry = 1
                mouseClick(1,"right",img,reTry)
                #print("右键",img) 
            #4代表输入
            elif cmdType.value == 4.0:
                inputValue = sheet1.row(i)[1].value
                pyautogui.press(inputValue)
                time.sleep(0.5)
                #print("键入:",inputValue)                                        
            #5代表等待
            elif cmdType.value == 5.0:
                #取图片名称
                waitTime = float(sheet1.row(i)[1].value)
                time.sleep(waitTime)
                #print("等待",waitTime,"秒")
            #6代表滚轮
            elif cmdType.value == 6.0:
                #取图片名称
                scroll = sheet1.row(i)[1].value
                pyautogui.scroll(int(scroll))
                #print("滚轮滑动",int(scroll),"距离")   
            #7代表输入
            elif cmdType.value == 7.0:
                
                pyautogui.press('1')
                time.sleep(0.1)
                inputValue = sheet1.row(i)[1].value
                if(type(inputValue)== float):
                    inputValue = int(inputValue)
                codeName = inputValue
                pyperclip.copy(inputValue)
                pyautogui.hotkey('ctrl','v')
                pyautogui.press('home')
                pyautogui.press('right')
                pyautogui.press('backspace')
                time.sleep(0.1)
                #print("键入:",inputValue)  
            #8代表写入路径
            elif cmdType.value == 8.0:
                inputValue = sheet1.row(i)[1].value
                if(os.path.exists(inputValue) == False):
                    os.makedirs(r"%s"%(inputValue))
                pyperclip.copy(inputValue)
                pyautogui.hotkey('ctrl','v')
                pyperclip.copy("\\"+str(downFileName)+".xls")
                pyautogui.hotkey('ctrl','v')
                #time.sleep(0.1)
                #print("写入路径:",inputValue+"\\"+str(downFileName)+".xls") 
                downFileName += 1
            #9代表清空输入框
            elif cmdType.value == 9.0:
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                pyautogui.press('backspace')
                #print("输入框已清空") 
            #10代表天数
            elif cmdType.value == 10.0:
                inputValue = sheet1.row(i)[1].value
                days = int(inputValue) - 3
                turns = 0
                #print("已获取天数:",inputValue)
            #11代表pageUp所在行数输入
            elif cmdType.value == 11.0:
                pageUpNum = i
                inputValue = sheet1.row(i)[1].value
                pyautogui.press(inputValue)
                #
                #print("键入:",inputValue)    
            #13代表最后一天的路径
            elif cmdType.value == 13.0:
                inputValue = sheet1.row(i)[1].value
                path = sheet1.row(i)[1].value
                if(os.path.exists(inputValue) == False):
                    os.makedirs(r"%s"%(inputValue))
                else:
                    shutil.rmtree(r"%s"%(inputValue))
                    time.sleep(2)
                    os.makedirs(r"%s"%(inputValue))
                time.sleep(0.1)
                pyperclip.copy(inputValue)
                pyautogui.hotkey('ctrl','v')
                time.sleep(0.1)
                pyperclip.copy("\\"+str(downFileName)+".xls")
                pyautogui.hotkey('ctrl','v')
                time.sleep(0.2)
                shareFileName = inputValue+"\\"+str(downFileName)+".xls"
                # time.sleep(0.5)
                # print("写入路径:",inputValue+"\\"+str(downFileName)+".xls") 
                #downFileName += 1
            #14代表进入日线图 判断天数是否大于250
            elif cmdType.value == 14.0:
                if(days>250):
                    for a in range(0,12):
                        pyautogui.press('down')
                        time.sleep(0.1)
                        print("向下移动%d次"%(a+1))
                time.sleep(0.1)
                inputValue = sheet1.row(i)[1].value
                pyautogui.press(inputValue)
                time.sleep(0.1)
                #print("键入:",inputValue)   
            #  15 判断前复权
            elif cmdType.value == 15.0:
                #取图片名称
                img = sheet1.row(i)[1].value
                reTry = 1
                mouseClick2(1,"left",img,reTry)
                #print("单击左键",img)
            #  16 floatEnd
            elif cmdType.value == 16.0:
                #取图片名称
                img = sheet1.row(i)[1].value
                reTry = 1
                if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                    reTry = sheet1.row(i)[2].value
                mouseClick3(1,"left",img,reTry)
                #print("单击左键",img)
            #17 判断是否存在目标文件
            elif cmdType.value == 17.0:
                #取图片名称
                nn = 0
                while True:
                    nn += 1
                    #print(shareFileName)
                    print("没有检测到文件",shareFileName," 当前次数为",nn)
                    time.sleep(0.1)
                    if (os.path.exists(shareFileName) == True):
                        print("已存在文件",shareFileName)
                        time.sleep(0.1)
                        break
                    elif(nn>20):
                        print("获取失败！重复获取今日数据!")
                        img1 = 'shares/cancel.png'
                        reTry = 1
                        #点击取消
                        imgResult = mouseClick(1,"left",img1,reTry)
                        if(imgResult == -1):
                            j -= 1
                            break
                        time.sleep(0.1)
                        #点击操作
                        img2 = 'shares/operate.png'
                        imgResult = mouseClick(1,"left",img2,reTry)
                        if(imgResult == -1):
                            j -= 1
                            break
                        time.sleep(0.1)
                        #点击明细数据导出
                        img3 = 'shares/dataOut.png'
                        imgResult = mouseClick(1,"left",img3,reTry)
                        if(imgResult == -1):
                            j -= 1
                            break
                        time.sleep(0.1)
                        #获取名称
                        shareFileName =  getCopyTxet()
                        time.sleep(0.1)
                        #点击导出
                        img4 = 'shares/leadOut.png'
                        imgResult = mouseClick(1,"left",img4,reTry)
                        if(imgResult == -1):
                            j -= 1
                            break
                        time.sleep(0.1)
                        nn = 0
                #print("单击左键",img)
            #18 获取当前文件名称
            elif cmdType.value == 18.0:
                # pyautogui.hotkey('ctrl','c')
                # time.sleep(0.1)
                shareFileName =  getCopyTxet()
                #time.sleep(0.1)
            #19 判断是否存在目标文件
            elif cmdType.value == 19.0:
                #取图片名称
                nn = 0
                while True:
                    nn += 1
                    #print(shareFileName)
                    print("没有检测到文件",shareFileName," 当前次数为",nn)
                    time.sleep(0.1)
                    if (os.path.exists(shareFileName) == True):
                        print("已存在文件",shareFileName)
                        time.sleep(0.1)
                        break
                    elif(nn>20):
                        print("获取失败！重复获取今日数据!")
                        img1 = 'shares/cancel.png'
                        reTry = 1
                        #点击取消
                        imgResult = mouseClick(1,"left",img1,reTry)
                        if(imgResult == -1):
                            j -= 1
                            break
                        time.sleep(0.1)
                        #输入3
                        pyautogui.press(3)
                        time.sleep(0.5)
                        #输入4
                        pyautogui.press(4)
                        #获取名称
                        inputValue = sheet1.row(i)[1].value
                        time.sleep(0.1)
                        shareFileName = inputValue+"\\"+str(1)+".xls"
                        time.sleep(0.1)
                        #输入名称
                        pyperclip.copy(shareFileName)
                        time.sleep(0.1)
                        pyautogui.hotkey('ctrl','v')
                        time.sleep(0.1)
                        #点击导出
                        img4 = 'shares/leadOut.png'
                        imgResult = mouseClick(1,"left",img4,reTry)
                        if(imgResult == -1):
                            j -= 1
                            break
                        time.sleep(0.1)
                        nn = 0
            #检测 复权 图片
            if cmdType.value == 20.0:
                #取图片名称
                img = sheet1.row(i)[1].value
                reTry = 1
                t20i = 0
                while True:
                    t20i += 1
                    imgResult = mouseClick(1,"left",img,reTry)
                    if(imgResult == -1):
                        pyautogui.press('f5')
                    elif(imgResult == 1 or t20i > 20):
                        break
                time.sleep(0.1)
                #print("单击左键",img)        
            i += 1
            if(days == 0 and i == sheet1.nrows):
                i = i + 100
            elif(days == -1 and i == 45):
                i = i + 100
            elif(days == -2 and i == 35):
                i = i + 100
            elif(i == sheet1.nrows and turns < days):
                i = pageUpNum
                turns += 1
            elif(i == sheet1.nrows and turns == days):
                #文件夹下文件个数。
                print(codeName,"的已下载文件数量：",len([lists for lists in os.listdir(path) if os.path.isfile(os.path.join(path, lists))]))
                if(len([lists for lists in os.listdir(path) if os.path.isfile(os.path.join(path, lists))]) != days+3 ):
                    j -= 1
                    break
        j += 1



if __name__ == '__main__':
    b = pyautogui.confirm(text='要开始程序么？在开始前，请确保当前处于xxxxxx的自选页面', title='请求框', buttons=['开始','取消'])
    if(b == '取消'):
        pyautogui.alert(text='程序运行结束！', title='请求框', button='OK')
        sys.exit()
    print('程序开始')
    #软件自动化
    mainWork()
    #批量修改文件名
    #ChangeFlieName()
    pyautogui.alert(text='程序运行结束！', title='请求框', button='OK')

    

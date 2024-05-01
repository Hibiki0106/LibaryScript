import pyautogui
import pandas as pd
import pyperclip
import time
import cv2
import keyboard
import os

'''
import mss
from PIL import Image
import pytesseract
import re
'''

def pause_program(event):
    if event.name == 'p':  # 当按下 'p' 键时暂停程序
        print("程序已暂停，按下 'p' 键继续执行...")
        keyboard.wait('p')  # 等待用户按下 'p' 键继续执行
        print("程序继续执行...")

# 监听键盘事件


def getDataOfFrame() :
    #dataframe = pd.read_excel("D:\Backup\Desktop\圖書\簡易按鍵精靈1\借書.xlsx", sheet_name='工作表1', usecols=range(1, 6) )
    datalist = dataframe.iloc[:, i].tolist()  
    pyperclip.copy(datalist[j])
    time.sleep(1)
    return datalist


def getChrome() :
    try :
        img = cv2.imread(r"chrome.png") 
        region = pyautogui.locateOnScreen( img, confidence=1 )
    except pyautogui.ImageNotFoundException :
        region = pyautogui.locateOnScreen( img, confidence=0.9 )
    return region




def getIdRegion():
    try :
        img = cv2.imread(r"idRegion.png") 
        region = pyautogui.locateOnScreen( img, confidence=0.7 )
    except pyautogui.ImageNotFoundException :
        print("find ID region again")
        img = cv2.imread(r"idRegion2.png")
        region = pyautogui.locateOnScreen( img, confidence=0.6 )
    return region


def doRecentRecords() :
    recent = f.readline().strip()
    find = False
    a, b = 0, 0
    print( "recent is : " + recent )
    while a < 5 :
        b = 0
        while b < len(dataframe.iloc[ :, a ] ) :
            if str( dataframe.iloc[ b, a ] ).strip() == recent :
                print ( "find!" )
                find = True
                break
            else : 
                b+=1
                
        if find : break
        a+=1
        
    if find == True : 
        print("do last times" )
        print( a )
        print( b )
        return a, b
    else : 
        return 0, 0



def detectHandling() : # 檢測處理中
    time.sleep(3)
    try :
        img = cv2.imread(r"handling.png") 
        region = pyautogui.locateOnScreen( img, confidence=0.6 )
        if region :
            print("handling now, please wait")
            detectHandling()
        else :
            print("handle complete")
    except pyautogui.ImageNotFoundException :
        print("handle complete")





def detectHadBorrowed() : # 檢測是否已經借過
    try :
        img = cv2.imread(r"borrowed.png") 
        region = pyautogui.locateOnScreen( img, confidence=0.7 )
        if region :
            print("had been borrowed")
            center = pyautogui.center( region )
            X, Y = center
            pyautogui.click( X, Y )
            detectHandling()
        else :
            pass
    except pyautogui.ImageNotFoundException :
        pass
        
        

        
def detectCD() : # 檢測是否有cd 有的話就還掉
    try :
        img = cv2.imread(r"returnCD.png") 
        region = pyautogui.locateOnScreen( img, confidence=0.8 )
        if region :
            print("had cd")
            img = cv2.imread(r"yes.png") 
            region = pyautogui.locateOnScreen( img, confidence=0.7 )
            center = pyautogui.center( region )
            X, Y = center
            pyautogui.click( X, Y )
            detectHandling()
        else :
            pass
    except pyautogui.ImageNotFoundException :
        pass
    



datalist = []
i, j = 0, 0
current_directory = os.getcwd()
file_path = os.path.join(current_directory, "借書.xlsx")

dataframe = pd.read_excel(file_path, sheet_name='工作表1', usecols=range(1, 6) )
print(dataframe)

f = open('returnRecord.txt','r+')
i, j = doRecentRecords()
#keyboard.on_press(pause_program)
#keyboard.wait('p')
datalist = getDataOfFrame()

chromeRegion = getChrome()
chromeCenter = pyautogui.center( chromeRegion )
chromeX, chromeY = chromeCenter
pyautogui.click( chromeX, chromeY )
pyautogui.press( 'F2' )



while i < 5 :
    
    if j == 0 :
        idRegion = getIdRegion()
        idCenter = pyautogui.center( idRegion )
        idX, idY = idCenter
        pyautogui.click( idX, idY )
        pyautogui.hotkey( 'ctrl', 'v' )
        pyautogui.press( 'enter' )
        detectHandling()
        pyautogui.press( 'F2' )
        j+=1
        
    
    print( datalist[j] )
    while j < len( datalist ) :
        datalist = getDataOfFrame()
        s = str(datalist[j])
        while  ( len(s) == 0 or s == "nan" ) and j < len( datalist )-1:
            j+=1
                
            datalist = getDataOfFrame()
            s = str(datalist[j])
            
        if len(s) == 0 or s == "nan" : break
        
        pyautogui.hotkey( 'ctrl', 'v' )
        pyautogui.press( 'enter' )
        detectHandling() #檢測是否還在處理
        #detectHadBorrowed() #檢測是否有借過
        detectCD() #檢測是否有cd
        
        f.seek(0)
        f.write( s + '\n' )
        f.flush()
        print( s )
        j+=1
        
    i+=1
    j = 0

f = open('returnRecord.txt','w')
f.write( '' )
f.close()

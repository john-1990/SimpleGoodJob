# -*- coding: utf-8 -*-

import os
import cv2
import numpy
import pytesseract
from PIL import Image
import difflib
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import *
import win32gui
import sys
import shutil
import win32api,win32con
import tkinter
import tkinter.messagebox

def string_similar(s1, s2):
    return difflib.SequenceMatcher(None, s1, s2).quick_ratio()

def takeSecond(elem):
    return elem[0]

global iCount
iCount = 1
pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files (x86)/Tesseract-OCR/tesseract.exe'
tessdata_dir_config = '--tessdata-dir "C:/Program Files (x86)/Tesseract-OCR/tessdata"'

#创建图片保存文件夹
curDir = os.path.abspath(os.path.dirname(__file__))
picDir = curDir + "\\Pic\\"
if os.path.exists(picDir):
    shutil.rmtree(picDir)
os.mkdir(picDir)


zsQue = []
zsAns = []
zsAnsItem = []
#读取Word，把所有的题目都处理，保存成对应的格式，用于比对
f = open("word.txt","r")
line = f.readline()
line = line.replace(" ","")
line = line[:-1]

bQue = True
while line:
    line = line.replace(" ","")
    line = line[:-1]
    if line != "":
        if bQue:
            zsQue.append(line)
            bQue = False
        else:
            zsAnsItem.append(line)
    else:
        if len(zsAnsItem) != 0:
            zsAns.append(zsAnsItem)
            zsAnsItem.clear()
            bQue = True

    line = f.readline()
zsAns.append(zsAnsItem)
f.close()

#print(len(zsQue))
#print(len(zsAns))
#获取投影屏幕HWND，并筛选范围
#hwnd = win32gui.FindWindow("CHWindow", "")
hwnd = win32gui.FindWindow("CHWindow","")
app = QApplication(sys.argv)
screen = QApplication.primaryScreen()

root=tkinter.Tk()
root.title('GUI')               #标题
root.geometry('300x300')        #窗体大小
root.resizable(False, False)    #固定窗体

def but():
    global iCount
    imgName = picDir + str(iCount)
    img = screen.grabWindow(hwnd).toImage()
    img.save(imgName + ".jpg")

    cvImg = Image.open(imgName + ".jpg")
    cropped = cvImg.crop((43, 111, 370, 397))
    strLines = pytesseract.image_to_string(cropped,lang='chi_sim')
    cropped.save(imgName + ".png")
    code = strLines.replace(" ","")
    code = code.replace("\n","")
    print(code)

    iCount += 1
    #开始比较
    fValues = []
    for i in range(len(zsQue)):
        fVal = string_similar(strLines,zsQue[i])
        fValues.append([fVal,i])

    fValues.sort(key=takeSecond,reverse=True)
        
    for i in range(3):
        #iIndex = fValues[i]
        #print(iIndex[1])
        result = fValues[i]
        strLine = zsQue[result[1]] + "\n"
        ansItem = zsAns[result[1]]
        for j in range(len(ansItem)):
            strLine = strLine  + ansItem[j] + "\n"
        bValue = a=tkinter.messagebox.askokcancel('提示', strLine)
        print(bValue)
        if bValue == True:
            break

tkinter.Button(root, text='Next',command=but).pack()
root.mainloop()





'''
cap = cv2.VideoCapture(0)
i = 1
while True:
    ret, frame = cap.read()
    cv2.imshow("capture", frame)

    key = cv2.waitKey(1)
    if key & 0xFF == ord('c'):
        cv2.imwrite(folder_path + str(i) + '.jpg', frame)
        #添加OCR识别，但是笔记本拍摄的图片分辨率太低，识别不了，采用投屏的方法
        i += 1
    elif key & 0xFF == ord('q'):
        break
        
cap.release()
cv2.destroyAllWindows()
'''

import pyautogui
from pywinauto.application import Application
import psutil
import time
import cv2
import numpy as np
import pyperclip
import win32com.client
import pathlib
pid_dic={}

#change
people=2 
#文本書入框位址
#第一個保戶位置
this_path=str(pathlib.Path(__file__).parent.absolute()) #可能不用改
#change

def get_pid():
    pids = psutil.pids()
    for pid in pids:
        p = psutil.Process(pid)
        pid_dic[p.name()]=pid

def open_line():
    pid=pid_dic["LINEAPP.exe"]
    app=Application(backend='uia').connect(process=pid)
    win=app["LINE"]
    if win.exists():
        win.maximize()
        print(pid)
    else:
        print("window not exist")


def imagesearch(image, precision=0.8):
    im = pyautogui.screenshot()
    img_rgb = np.array(im)
    img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
    template = cv2.imread(image, 0)
    w,h=template.shape[::-1]

    res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    if max_val < precision:
        return (-1,-1)
    return (max_loc[0]+w/2, max_loc[1]+h/2) #返回圖片座標(center)
def paste(foo):
  pyperclip.copy(foo)
  pyautogui.hotkey('ctrl', 'v')


people=people-1
get_pid()
open_line()


time.sleep(0.5)
friend=imagesearch("friend_gray.png")
pyautogui.click(friend)
pyautogui.click(friend[0]+200,friend[1])                #移動到文本書入框
paste("保戶")


# word內容複製
wordapp=win32com.client.Dispatch('Word.Application')
doc=wordapp.Documents.Open(this_path+r"\content.docx")
doc.Content.Copy()
doc.Close()
#  

time.sleep(0.5)
pyautogui.click(friend[0]+200,friend[1]+180)              #第一個保戶位置


while people>0:
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('down')
    time.sleep(0.15)
    people=people-1

pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('enter')
pyautogui.hotkey('down')





#size=pyautogui.size() #(2880,1800) #解析度


# print(pyautogui.position())

import logging
import os
from datetime import datetime, timedelta
import pyautogui as pg
import time
import pyperclip
import pandas as pd

class e:
    def __init__(self):
        self.errflag
        #errflag values
        # 0 : error encountered
        # 1 : Successful
        # 2 : error encountered, continue to next iteration
        self.errcontext

def assignvalues(df,i):
    # default values
    action = ""
    image = ""
    xoffset = 0
    yoffset = 0
    prenewtext = ""
    keypress = "Enter"
    max_timer = 60
    clicktype = 1
    log = ""
    iferrorgoto=0

    if df['action'][i] == df['action'][i]: action = df['action'][i]
    if df['image'][i] == df['image'][i]: image = df['image'][i]
    if df['xoffset'][i] == df['xoffset'][i]: xoffset = df['xoffset'][i]
    if df['yoffset'][i] == df['yoffset'][i]: yoffset = df['yoffset'][i]
    if df['newtext'][i] == df['newtext'][i]:
        str = df['newtext'][i]
        for j in str.split():
            if j.replace('_', '').isalpha() :
                prenewtext = prenewtext + " globalvar[\"" + j + "\"]"
            else :
                prenewtext = prenewtext + " " + j

    if df['keypress'][i] == df['keypress'][i]: keypress = df['keypress'][i]
    if df['max_timer'][i] == df['max_timer'][i]: max_timer = df['max_timer'][i]
    if df['clicktype'][i] == df['clicktype'][i]: clicktype = df['clicktype'][i]
    if df['log'][i] == df['log'][i] : log = df['log'][i]
    if df['iferrorgoto'][i] == df['iferrorgoto'][i]:
        iferrorgoto = int(df['iferrorgoto'][i])
        iferrorgoto = df.values.transpose()[0].tolist().index(int(iferrorgoto)) - 1

    # print(action, image, xoffset, yoffset, prenewtext, keypress, max_timer, clicktype, log, iferrorgoto)
    return (action, image, xoffset, yoffset, prenewtext, keypress, max_timer, clicktype, log, iferrorgoto)


def clicknwrite(image="", xoffset=0, yoffset=0, newtext="", keypress="", max_timer=60, log="", oldtext=None):
    pg.useImageNotFoundException()
    pyperclip.copy('')
    delaytime = 2
    e.errflag= 1
    e.errcontext = ""
    print(image)
    img=image.split(",")
    if not (image == ""):  # does not require mouse position
        if not (os.path.exists(img[0])) :
            e.errflag = 0
            e.errcontext = "ERROR : Reference image, " + img[0] + ", not found"
            return e
        ts = datetime.today() + timedelta(seconds=max_timer)
        x = y = 0
        while ts > datetime.today():
            for i in img:
                try:
                    x, y = pg.locateCenterOnScreen(i, confidence=0.8)
                    break
                except pg.ImageNotFoundException:
                    pass
                if not (x==0 or y==0):
                    break
            if not (x == 0 or y == 0):
                break
        if x == 0 or y == 0:
            # logging.info("ERROR : Image "+image+" not found")
            e.errflag = 0
            e.errcontext = "ERROR : Image, " + i + ", not found on screen"
            return e
        else:
            # slowdown mouse moevement
            x1, y1 = pg.position()
            t = ((x1 - x) ** 2 + (y1 - y) ** 2) ** 0.5 / 600

            # move to image location with offset
            pg.moveTo(x + xoffset, y + yoffset, t)
            time.sleep(delaytime)
            # copy old text in clipborad
            pg.doubleClick()
            pg.hotkey('ctrl', 'a')
            time.sleep(0.3)
            pg.hotkey('ctrl', 'c')
            time.sleep(0.3)
            pg.hotkey('Backspace')
            time.sleep(0.3)

    # type new text
    if newtext != "":
        #pg.hotkey('shift')
        pg.typewrite(newtext,0.2)
        time.sleep(delaytime)

    # key press
    if keypress != "":
        pg.press(keypress, interval=0.2)

    #if logging and pyperclip.paste()=="":
        # if log:
        e.errflag = 1
       #e.errcontext = "SUCCESFUL: old text :" + pyperclip.paste() + "; new text :" + newtext
       # print("SUCCESFUL: old text :" + pyperclip.paste() + " new text :" + newtext)
       # logging.info("SUCCESFUL: old text :" + pyperclip.paste() + "; new text :" + newtext)
       # return e
    return e


def findnclick(image, xoffset=0, yoffset=0, clicktype=1, max_timer=60):
    pg.useImageNotFoundException()
    delaytime = 3
    e.errflag = 1
    e.errcontext = ""
    img = image.split(",")
    print(img)
    if not (os.path.exists(img[0])):
        e.errflag = 0
        e.errcontext = "ERROR : Reference image, " + img[0] + ", not found"
        return e
    ts = datetime.today() + timedelta(seconds=max_timer)
    x = y = 0
    while ts > datetime.today():
        for i in img:
            try:
                x, y = pg.locateCenterOnScreen(i, confidence=0.8)
                break
            except pg.ImageNotFoundException :
                pass
            if not (x==0 or y==0):
                break
        if not (x == 0 or y == 0) :
            break
    if x == 0 or y == 0:
        logging.info("ERROR : Image " + image + " not found")
        print("ERROR : Image " + image + " not found")
        e.errflag = 0
        e.errcontext = "ERROR : Image, " + i + ", not found on screen"
        return e
    else:
        x1, y1 = pg.position()
        t = ((x1 - x) ** 2 + (y1 - y) ** 2) ** 0.5 / 700
        pg.moveTo(x + xoffset, y + yoffset, t)
        time.sleep(delaytime)
        if clicktype == 1:
            pg.click()
        elif clicktype == 2:
            pg.doubleClick()
        elif clicktype == 3:
            pg.click(); pg.press('Enter')
        elif clicktype == 4:
            pg.rightClick()
    time.sleep(delaytime)
    e.errflag = 1
    e.errcontext = "SUCCESS : Image, " + i + ", found on screen"
    return e


def RPAsteps(dataframe: object, globalvar: object) -> object :
    # global globalvar
    i=0
    errorflag=1
    errorcontext=""
    while i<len(dataframe.index):
        try:
            (action, image, xoffset, yoffset, prenewtext, keypress, max_timer, clicktype, log, iferrorgoto) = assignvalues(dataframe,i)
            #print(action, image, xoffset, yoffset, prenewtext, keypress, max_timer, clicktype, log,iferrorgoto)
            #convert text into formula
            if prenewtext != "":
                newtext = eval(prenewtext)
            if action.strip() == "clicknwrite":
                e = clicknwrite(image=image, xoffset=xoffset, yoffset=yoffset, newtext=newtext,
                                  keypress=keypress, max_timer=max_timer, log=log)
                e.errflag=e.errflag * errorflag
                if e.errflag==0:
                    if iferrorgoto==0:
                        if errorflag==0 : e.errcontext=""
                        return e
                    else:
                        i=iferrorgoto
                        errorflag=0
            elif action.strip() == "findnclick" :
                e=findnclick(image=image, xoffset=xoffset, yoffset=yoffset, max_timer=max_timer, clicktype=clicktype)
                e.errflag = e.errflag * errorflag
                if e.errflag==0:
                    if iferrorgoto==0:
                        if errorflag == 0 : e.errcontext = ""
                        return e
                    else:
                        i=iferrorgoto
                        errorflag = 0
            elif action.strip() == "exit" :
                e.errflag = 1
                e.errcontext = "Exit from iteration"
                return e
            else:
                e.errflag = 0
                e.errcontext = "ERROR : RPA action not defined in excel"
                return e
        except Exception as error:
            #print(error)
            e.errflag = 0
            e.errcontext = "ERROR : " + "some issue"
            return e
        i=i+1
    return e

import pandas as pd
import pyautogui as pg
import webbrowser as wb
import time
import pyautogui


df = pd.read_excel('data.xlsx')


no1 = df.iloc[0,0]
msg1 = df.iloc[0,1]

no2 = df.iloc[1,0]
msg2 = df.iloc[1,1]

no2 = df.iloc[2,0]
msg2 = df.iloc[2,1]

no2 = df.iloc[3,0]
msg2 = df.iloc[3,1]

no2 = df.iloc[4,0]
msg2 = df.iloc[4,1]


number = no1
my_msg = msg1

wb.open('https://web.whatsapp.com/')

def sendmessage():
    time.sleep(10)
    print(pyautogui.position())

    # click on search bar
    pyautogui.click(383,256)
    pyautogui.typewrite(number)


    time.sleep(5)

    #click on person 
    time.sleep(5)
    pg.press('enter')

    time.sleep(5)
    pg.hotkey('ctrl', 'w')
    


sendmessage()
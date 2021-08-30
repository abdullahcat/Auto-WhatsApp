import pyautogui as pg
import pandas as pd
import webbrowser as wb
import time
import pandas as pd



df = pd.read_excel("data.xlsx")
data_dict = df.to_dict('list')
print(df)

tel = data_dict['Telefon']
message = data_dict['Message']

zipp = zip(tel,message)


first = True
for tel,message in zipp:

    time.sleep(5)
    wb.open("https://web.whatsapp.com/send?phone="+tel+"&text="+message)

    if first:
        time.sleep(5)
        first=False

    width,height = pg.size()
    pg.click(width/2,height/2)

    time.sleep(7)
    pg.press('enter')

    time.sleep(7)
    pg.hotkey('ctrl', 'w')
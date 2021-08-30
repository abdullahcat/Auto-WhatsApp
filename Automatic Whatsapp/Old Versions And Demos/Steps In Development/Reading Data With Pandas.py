
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from fontTools.ttLib import TTFont
import time, os, datetime, webbrowser, pandas, pyautogui

#Special Font
SF = TTFont('SFCompact.ttf')

#DataBase Options
df = pandas.read_excel("data.xlsx")
data_dict = df.to_dict('list')

tel = data_dict['Telefon']
Name = data_dict['Name']
message = data_dict['Message']

zipp = zip(tel,message)

#Main Window Options
root = Tk()
root.title('Auto WhatsApp')
root.geometry("550x350")
root.resizable (False, False)
root.configure(bg="#232830")


#Creating Interface
interface = ttk.Notebook(root)
interface.pack()

#Interface Options and Creating Tabs
main_frame = Frame(interface, width=550, height=450, bg="#232830")
second_frame = Frame(interface, width=550, height=450, bg="#232830")

def hide():
    interface.forget(1)

main_frame.pack(fill="both", expand=1,anchor='center')
second_frame.pack( fill="both",expand=1,anchor='center')

#Tab Options
interface.add(main_frame, text="Home")
interface.add(second_frame, text="New")

#Main Frame Tabs Items
timelabel = Label(main_frame,bg='#232830',fg='#bb86f0')

def time():
    dt = datetime.datetime.now().strftime("%H:%M:%S")
    timelabel.config(text=dt)
    timelabel.after(1000, time)

timelabel.pack(pady=55)
timelabel.config(font=("SF", 25))
time()

global label1
label1 = Label(main_frame,text="You Have No Message To Send.",bg='#232830',fg='white',font='SF')
label1.pack()

global functionlabel2
functionlabel2 = Label(main_frame, text='*Please click help before starting.*',bg='#232830',fg='#67b3e6')
functionlabel2.pack(pady=50)
functionlabel2.config(font=("SF", 10))

line = Label(main_frame,text="__________________________________________________",bg='#232830',fg='white',font='SF')
line.pack(pady='1000')

#Creating Menu
wmenu = Menu(main_frame)
root.config(menu=wmenu)

#Menu Commands
def gomessage():
    interface.add(second_frame, text="New")
    interface.select(1)

def gomain():
    interface.select(0)

def gowebsite():
    webbrowser.open_new_tab("https://github.com/abdullahcat/Automatic-Messages-With-WhatsApp")

#Menu Items
file_menu = Menu(wmenu,tearoff=0)
wmenu.add_cascade(label='File', menu=file_menu)
file_menu.add_command(label='Exit',command=root.quit)

new_menu = Menu(wmenu,tearoff=0)
wmenu.add_cascade(label='Edit', menu=new_menu)
new_menu.add_command(label='New Message',command=gomessage)

help_menu = Menu(wmenu,tearoff=0)
wmenu.add_cascade(label='Help', menu=help_menu)
help_menu.add_command(label='User Manual',command=gowebsite)
help_menu.add_separator()
help_menu.add_command(label='About Auto WhatsApp',command=gowebsite)
help_menu.add_command(label='Github Repo',command=gowebsite)

#OUR COMMANDS
#Message Send Command Using PyAutoGUI
def sendbulk():
    


    for tel,message in zipp:
    
       time.sleep(5)
       webbrowser.open("https://web.whatsapp.com/send?phone="+tel+"&text="+message)

    if True:
        time.sleep(5)
        first=False

    width,height = pyautogui.size()
    pyautogui.click(width/2,height/2)

    time.sleep(7)
    pyautogui.press('enter')

    time.sleep(7)
    pyautogui.hotkey('ctrl', 'w')

#Creating Tab Command
 


#Second Frame Tabs Items
who = Label(second_frame,text="Who do you want to send message?", font='SF',fg='white',bg='#232830')
who.pack()

#ListBox Scrollbar
scrollbar_frame = Frame(second_frame)
list_scrollbar =Scrollbar(scrollbar_frame, orient=VERTICAL,bg="black")

#Listbox
contacts = Listbox(second_frame,height=3,bg="black",fg="white",font="SF",yscrollcommand=list_scrollbar.set,selectmode=MULTIPLE)
contacts.pack()

##Configure Scrollbar
#list_scrollbar.config(command= contacts.yview)
#list_scrollbar.pack(side=RIGHT, fill=Y)
#scrollbar_frame.pack()

#Items In Listbox
#contacts.insert(END, "Abdullah")
#contacts.insert(1, "Ece")
#contacts.insert(0, "Mete")
#Adding List Of Items

my_contacts = ["Defne","Abdullah","Daft Punk","Giorgio","Defne","Ahmet","Deniz","Alex"]
for person in my_contacts:
    contacts.insert(0, person)

#Functions For ListBox
def select():
    result = ''
    for item in contacts.curselection():
        result = result + str(contacts.get(item)) + '\n'
     
    print(result)
    
    label1.config(text="Waiting for the send message to contacts that you chose.")
    functionlabel2.config(text='')
    print(alarmH,alarmM)
    #functionlabel2.config(text="Waiting for the send message to" + my_contacts.get(select))
    while(1 == 1):
       if(alarmH == datetime.datetime.now().hour and
       alarmM == datetime.datetime.now().minute) :
           gomessage()
       
       break
    gomain()
    hide()
    


#def selectAll():
#    print(my_contacts.curselection())


when = Label(second_frame,text="When do you want to send message?", font='SF',fg='white',bg='#232830')
when.pack()

ghour = Spinbox(second_frame,from_=0,to=23,wrap=True,width=5,state="readonly")
ghour.pack(pady=10)
gminute = Spinbox(second_frame,from_=0,to=59,wrap=True,width=5)
gminute.pack()
    
alarmH = int(ghour.get())
alarmM = int(gminute.get())



CancelButton = Button(second_frame, text="Cancel", font='SF',bg='#eb4034',fg='white',width=8,command=gomain)
CancelButton.pack(pady=10)

SaveButton = Button(second_frame, text="Save",font='SF',bg='#21856a',fg='white',width=8,command=select)
SaveButton.pack()

root.mainloop()



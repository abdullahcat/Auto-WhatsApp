import webbrowser
import time
import os
from string import ascii_letters 
import schedule


def send_whatsapp_msg(name, message):
    webbrowser.open_new_tab(f'https://web.whatsapp.com/send?phone=+90{name}&text={message}')
    return schedule.CancelJob  # do this if you want to send this message only once
    # or just exit the program entirely if you don't want to run any more tasks
    # os._exit(0)
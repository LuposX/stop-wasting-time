'''
Author: Lupos
Started: 26.04.2019
Lang: Phyton
Description: A Programm that warns u when u when u are to long on youtube
             and closes youtube after some time.

Ideas: maybe popup size dynamic
'''
import win32com.client
import json
import os
from win32api import GetSystemMetrics
from playsound import playsound
import time
import psutil

import threading
from _thread import start_new_thread

import tkinter as tk
from tkinter import ttk

# boolean for timer
checker = False
counter = 0
timer = 0
p = 1 # for ewhen a popup should show up

# get height and width from monitor
ws = GetSystemMetrics(0)
hs = GetSystemMetrics(1)
w = 400
h = 80
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)

# function for fingin if programm is running
def checkIfProcessRunning(processName):
    '''
    Check if there is any running process that contains the given name processName.
    '''
    # Iterate over the all the running process
    for proc in psutil.process_iter():
        try:
            # Check if process name contains the given name string.
            if processName.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False;

# function for popup
def popupmsg(msg):
    LARGE_FONT = ("Verdana", 12)
    NORM_FONT = ("Verdana", 10)
    SMALL_FONT = ("Verdana", 8)
    popup = tk.Tk()
    popup.wm_title("Warning!!")
    label = ttk.Label(popup, text=msg, font=NORM_FONT)
    label.pack(side="top", fill="x", pady=10)
    B1 = ttk.Button(popup, text="Okay", command = popup.destroy)
    B1.pack()
    popup.geometry('%dx%d+%d+%d' % (w, h, x, y))
    popup.attributes("-topmost", True)
    popup.mainloop()

#function for playing sound
def playingsound(name_of_audio):
    playsound(name_of_audio)

# opens the settings.jsonb file and reads some settings
filename = "settings.json"

if filename:
    # checks that json file is not corrupted
    try:
        with open(filename, 'r') as f:
            # loads the file in python
            datastore = json.load(f)
            # sets the timeout time
            timeout_time = datastore['timeout_time']
            # sets the website filter
            website_filter = datastore['website_filter']
            # sets frequency/how often we see a popup
            time_frequency_popup = datastore['time_frequency_popup']
            # set when firefox gets killed
            time_kill_firefox = datastore['time_kill_firefox']
            # path/name of audio track when firefox gets killed
            name_of_audio = datastore['name_of_audio']
    except:
        print("json format is corrupted")
        # default settings
        timeout_time = 30
        website_filter = ["youtube"]
        time_kill_firefox = 1200
        name_of_audio = "just_do_it.wav"

if __name__ == "__main__":
    # open a handler to AutoItX3
    # AutoItX3 is a script language for automation purpose
    autoit = win32com.client.Dispatch("AutoItX3.Control")
    autoit.ClipPut("")
    if autoit.WinExists("[CLASS:MozillaWindowClass]"):
        autoit.WinActivate("[CLASS:MozillaWindowClass]")
    elif autoit.WinExists("[CLASS:IEFrame]"):
        autoit.WinActivate("[CLASS:IEFrame]")
    elif autoit.WinExists("[CLASS:Chrome_WidgetWin_1]"):
         autoit.WinActivate("[CLASS:Chrome_WidgetWin_1]")

    # every 2 seconds
    while(True):
        #counter gets each iteration reseted
        counter = 0
        # copies address in clipboard
        autoit.Sleep(100)

        if checkIfProcessRunning('firefox'):
            autoit.send("!d")
            autoit.send("^c")
            adressBar = autoit.ClipGet()
        else:
            adressBar = "a"

        # dynamic filter gets website_filter from json file
        for i in range(len(website_filter)):
            if website_filter[i] in adressBar:
                counter += 1
                print("true")
            else:
                print("false")

        # when at least 1 filter is detected than start timer
        print("counter: ", counter)
        if counter > 0:
            timer += timeout_time

        print("Timer: ", timer)
        # only a new thread every % variable seconds. Every time_frequency_popup seconds one popup
        if timer > p * time_frequency_popup and timer != 0:
            p += 1
            msg_string = "You're wasting your time since: " + str(timer) + "s."
            start_new_thread(popupmsg, (msg_string,))
            # popupmsg(msg_string)

        # time_kill_firefox time when firefox gets killed in settings.json
        if counter > 0 and timer > time_kill_firefox:
            # trys to kill firefox if open
            try:
                os.system("taskkill /im firefox.exe /f")
                print("killing firefox")
                timer = 0
                counter = 0
                p = 0
                start_new_thread(playsound, (name_of_audio,))
            except:
                print("firefox is already dead")

        time.sleep(timeout_time)
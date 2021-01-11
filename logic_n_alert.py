#
#   Module responsible for logic and exeptions handeling
#

from tkinter import *
from tkinter import filedialog, simpledialog, messagebox
from tkinter.ttk import *
from applications import application_urlop, application_reward, application_one, application_hdk, application_boots
from docx_lib import docx_urlop, docx_reward, docx_one, docx_hdk, docx_boots
import requests
import os
from sys import exit
import platform

# application choosing
def application_choose(choose, window):                                                                                 # subapplication choosing a formula
    if(choose == "urlop okolicznościowy"):
        application_urlop(window)
    elif(choose == "urlop nagrodowy"):
        application_reward(window)
    elif(choose == "przepustka jednorazowa"):
        application_one(window)
    elif(choose == "HDK(nie działa)"):
        application_hdk(window)
    elif(choose == "buty wojskowe"):
        application_boots(window)
    else:
        messagebox.showinfo("NoneAppFonund", "Nie wybrano odpowiedniej aplikacji.")
        sys.exit()

def docx_choose(app):                                                                                                   # subapplication choosing a document
    if(app == application_urlop()):
        docx_urlop()
    elif(app == application_reward()):
        docx_reward()
    elif(app == application_one()):
        docx_one()
    elif(app == application_hdk()):
        docx_hdk()
    elif(app == application_boots()):
        docx_boots()

# updater configuration
def update_me(win, pro, pro_label):                                                                                     # main function that update program(web sync)
        # download file
    if platform.system() == "Windows":                                                                                  # Windows system recognition

        res = requests.get("https://github.com/Jamro90/Generator/raw/master/generator.exe")                             # making request on github

    elif platform.system() == "Linux":                                                                                  # Linux system recognition

        res = requests.get("https://github.com/Jamro90/Generator/raw/master/generator")                                 # making request on github

    elif platform.system() == "Darwin":                                                                                 # Mac system recognition
        pass
    try:
        file = open("generator", "wb")                                                                                  # making a new program
        print("update completed")

    except:
        print("Program working...")
        os.remove("generator")                                                                                          # removing existed program
        file = open("generator", "wb")                                                                                  # making a new

    for chunk in res.iter_content(chunk_size = 4096):                                                                   # loop that magazine chunks

        file.write(chunk)
        pro["value"] = 100
        win.update_idletasks()

def restart_program():
    print("Restarting...")
    exit()

def progress():                                                                                                         # progress subapplication
    # progress of update configuration
    progress = Tk()
    progress.geometry("700x200")                                                                                        # window resolution
    progress.title("Update")                                                                                            # window title
    proFrame = Frame(progress)                                                                                          # frame

    # bar workplace
    pro_title = Label(proFrame, text = "Update status", font = ("arial", 20))                                           # main label "Update status"
    pro_label = Label(proFrame, text = "100%", font = ("arial", 10))                                                    # progress label
    pro = Progressbar(proFrame, orient = "horizontal", mode = "determinate", length = 300)                              # Progresbar
    cancel_button = Button(proFrame, text = "zamknij", command = restart_program)                                       # button "zamknij"
#    update_button = Button(proFrame, text = "update", command = update_me)
    update_me(progress, pro, pro_label)
        # display
    pro_title.grid(column = 3, row = 1, padx = 5, pady = 10)                                                            # position of label "Update status"
    pro_label.grid(column = 4, row = 5, padx = 5, pady = 10)                                                            # position progress status label
    pro.grid(column = 3, row = 5, padx = 5, pady = 10)                                                                  # position of Progressbar
    cancel_button.grid(column = 1, row = 10, padx = 5, pady = 10)                                                       # position of button "zamknij"
    #update_button.grid(column = 4, row = 10, padx = 5, pady = 10)

    proFrame.pack()                                                                                                     # progress frame
    progress.mainloop()                                                                                                 # mainloop of subapplication

def update(var):
    if var == "1":
        progress()
        restart_program()
    else:
        pass

    # show specific application form

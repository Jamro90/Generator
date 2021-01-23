#
#   Module responisble for GUI and style of application
#

from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog, simpledialog, messagebox
from tool_lib import toolbar_settings
from logic_n_alert import application_choose, docx_choose, update
from date_lib import date_read
from applications import application_urlop, application_reward, application_hdk, application_boots

# font conofigurations
version = "1.2.1"                                                                                                             # program version
font_size20 = 20                                                                                                            # font size config (20)
font_size10 = 10                                                                                                            # font size config (10)

def window_start(title = "Generator ver " + version, geometry = "450x250", font_name = "arial"):
    win = Tk()
    win.title(title)                                                                    							        # main window title
    win.geometry(geometry)

    # GUI navigator
    start_label = Label(win, text = title, font = (font_name, font_size20))
    null_label = Label(win, text = "\t", font = (font_name, font_size10))
    # Update checker
    check_value = IntVar()
    def updater():
        update(str(check_value.get()))
        print(str(check_value.get()))

        # next
        win.destroy()
        window_main()

    # shortcut
    def updater_short(event = ""):
        updater()
    win.bind("<Return>", updater_short)

    check = Checkbutton(win, onvalue = 1, offvalue = 0, variable = check_value, state = NORMAL)
    check_label = Label(win, text = "Update", font = (font_name, font_size10))
    update_button = Button(win, text = "Start", command = updater)

    # display
    null_label.grid(column = 0, row = 0, padx = 5, pady = 10)
    start_label.grid(column = 10, row = 1, padx = 5, pady = 10)
    check_label.grid(column = 10, row = 5, padx = 5, pady = 10)
    check.grid(column = 15, row = 5, padx = 5, pady = 10)
    update_button.grid(column = 15, row = 15, padx = 5, pady = 10)

    win.mainloop()

def window_form(title = "Generator", geometry = "900x1100+800+0", font_name = "arial"):

        # general config
    win = Tk()
    win.title(title)                                                                    							        # main window title
    win.geometry(geometry)                                                                                                  # main window geometry

    mainFrame = Frame(win)
        # I/O objects
    toolbar_settings(win, option_var.get())                                                                                                   # toolbar
    application_header = Label(mainFrame, text = "Wniosek o " + option_var.get(), font = (font_name, font_size20))

    scroll = Scrollbar(win)

        # dislpay
    application_header.pack()                                                                                               # label application
    application_choose(option_var.get(), mainFrame)                                                                         # aplication choose

    scroll.pack(side = "right", fill = Y)
    mainFrame.pack()
    win.mainloop()

def window_main(title = "Generator ver " + version, geometry = "700x300+50+0", font_name = "arial"):

    # general config
    win = Tk()										                                                                        # main window object
    win.title(title)	       						                                                                        # main window title
    win.geometry(geometry)		    				                                                                        # main window geometry
    frameWin = Frame(win)							                                                                        # main frame of main window

    # shortcut
    def next(event = ""):
        window_form()
    win.bind("<Return>", next)

    def delete(event = ""):
        win.destroy()
    win.bind("<Escape>", delete)

    # I/O objects
    header = Label(frameWin, text = title, font = (font_name, font_size20))                                                 # frame title
    option_label = Label(frameWin, text = "wniosek", font = (font_name, font_size10))                                       # combobox label
    global option_var
    option_var = StringVar()                                                                                                # combobox value storage
    option = Combobox(frameWin, width = 35, textvariable = option_var)                                                      # cobobox object
    option["value"] = ("urlop okolicznościowy", "urlop nagrodowy", "przepustka jednorazowa", "HDK(nie działa)", "buty wojskowe", "zmianę służby na kompanii")        # combobox values
    option.current(0)                                                                                                       # current value

    next_button = Button(frameWin, text = "Dalej", command = window_form)                                                   # button that shows main window
    exit_button = Button(frameWin, text = "Wyjdź", command = win.destroy)                                                   # exit button

    # display
    frameWin.pack()																									    	# main frame of main window
    header.grid(column = 1, row = 0, padx = 25, pady = 25)	                											    # show header
    option_label.grid(column = 0, row = 1, padx = 25, pady = 25)	                                                        # show combobox label
    option.grid(column = 1, row = 1, pady = 25)	                                                                            # show combobox
    option.current()                                                                                                        # show current value of combobox
    next_button.grid(column = 3, row = 3, pady = 25)	                                                                    # show button "Dalej"
    exit_button.grid(column = 2, row = 3, padx = 25, pady = 25)	                                                            # show button "Wyjdź"

    win.mainloop()																										    # program loop


# -----|MODULES|----- #

from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from datetime import datetime

# -----|DATA CONVERT|----- #

def date_read():

	date = datetime.now()						# import date from local host

	if(float(date.day) < 10):					# add 0 befor day less than 10
		day = "0" + str(date.day)
	else:
		day = str(date.day)
	if(int(date.month) < 10):					# add 0 befor month less than 10
		month = "0" + str(date.month)
	else:
		month = str(date.month)

	date_right = str(day + "." + month + "." + str(date.year) + "r.")	# date formula
	return date_right

# -----|WINDOW'S SETTINGS|----- #

win = Tk()								# main window object
win.title("Generator")							# main window title
win.geometry("1200x850")						# main window geometry
frameWin = Frame(win)							# main frame of main window

# -----|TOOLBAR'S CONTROLS|----- #

def fileNew():
	print("nie działa :(")

def fileOpen():								# function that open file
	response = filedialog.askopenfile(initialdir = "/", title = "select", filetypes = (("documents", "*.doc"), ("all files", "*.*")))

def fileSave():								# function that saves
	response = filedialog.asksaveasfile(initialdir = "/", title = "select", mode = "w", defaultextension = ".doc")

def copyText():								# WORK IN PROGRESS

	pass

def pasteText():							# WORK IN PROGRESS

	pass

def cutText():								# WORK IN PROGRESS

	pass


# -----|GUI APPLICATION|----- #

screen_title = Label(frameWin, text = "Generator", font = ("arial", 27))

date = str(date_read())
date_label = Label(frameWin, text = date, font = ("arial", 10))
date_label2 = Label(frameWin, text = "Data", font = ("arial", 10))

place = StringVar()
place_entry = Entry(frameWin, textvariable = place, width = 41, font = ("arial", 10))
place_take = place_entry.get()
place_label = Label(frameWin, text = "Miejscowość: ", font = ("arial", 10))

name = StringVar()
name_entry = Entry(frameWin, textvariable = name, width = 41, font = ("arial", 10))
name_take = name_entry.get()
name_label = Label(frameWin, text = "Imię i nazwisko: ", font = ("arial", 10))

level = StringVar()
level_option = Combobox(frameWin, width = 35, textvariable = level)
level_option["values"] = ("szeregowy", "starszy szereregowy", "kapral", "starszy kapral", "plutonowy", "sierżant", "starszy sierżant")
level_label = Label(frameWin, text = "Stopień: ", font = ("arial", 10))

toWho = StringVar()
toWho_option = Combobox(frameWin, width = 35, textvariable = toWho)
toWho_option["values"] = ("porucznik", "generał","major")
toWho_label = Label(frameWin, text = "Do kogo: ", font = ("arial", 10))

direct = StringVar()
direct_entry = Entry(frameWin, textvariable = direct, width = 41, font = ("arial", 10))
direct_take = direct_entry.get()
direct_label = Label(frameWin, text = "batalion/kompania/pluton", font = ("arial", 10))

title = StringVar()
title_entry = Entry(frameWin, textvariable = title, width = 41, font = ("arial", 10))
title_take = title_entry.get()
title_label = Label(frameWin, text = "Nagłówek: ", font = ("arial", 10))

topic = StringVar()
topic_entry = Entry(frameWin, textvariable = topic, width = 41, font = ("arial", 10))
topic_take = topic_entry.get()
topic_label = Label(frameWin, text = "Dotyczy: ", font = ("arial", 10))

add = StringVar()
add_entry = Entry(frameWin, textvariable = add, width = 41, font = ("arial", 10))
add_take = add_entry.get()
add_label = Label(frameWin, text = "załącznik: ", font = ("arial", 10))
add_button = Button(frameWin, text = "Dodaj", command = None)

def acceptData():
	name.set("")
	place.set("")
	direct.set("")
	title.set("")
	topic.set("")
	add.set("")
accept_button = Button(frameWin, text = "Akceptuj", command = acceptData)

# -----|TOOLBAR'S SETTINGS|----- #

menu_obj = Menu(win)							# menu object

file = Menu(menu_obj, tearoff = 0)					# 'File' option (variable)
menu_obj.add_cascade(label = "File", menu = file)			# 'File' option creation
file.add_command(label = "New", command = fileNew)			# 'New' command
file.add_command(label = "Open", command = fileOpen)			# 'Open' command
file.add_command(label = "Save", command = fileSave)			# 'Save' command
file.add_separator()
file.add_command(label = "Exit", command = win.destroy)			# 'Exit' command

edit = Menu(menu_obj, tearoff = 0)					# 'Edit' option (variable)
menu_obj.add_cascade(label = "Edit", menu = edit)			# 'Edit' option creation
edit.add_command(label = "Copy", command = copyText)			# 'Copy' command
edit.add_command(label = "Cut", command = cutText)			# 'Cut' command
edit.add_command(label = "Paste", command = pasteText)			# 'Paste' command
edit.add_command(label = "Select all", command = None)			# 'Select all' command

help = Menu(menu_obj, tearoff = 0)					# 'Help' option (variable)
menu_obj.add_cascade(label = "Help", menu = help)			# 'Help' option creation
help.add_command(label = "About", commend = None)			# 'About' command
help.add_separator()
help.add_command(label = "Manual", command = None)			# 'Manual' command

# -----|FILE MANAGMENT|----- #



# -----|DISPLAY|----- #

screen_title.grid(column = 3, row = 0, pady = 25)

date_label2.grid(column = 1, row = 1, pady = 5)				# actual label of date
date_label.grid(column = 2, row = 1, pady = 5)				# label of date

place_entry.grid(column = 4, row = 1, pady = 5)				# space for place
place_label.grid(column = 3, row = 1, pady = 5)				# label of place

name_entry.grid(column = 2, row = 2, pady = 5)				# space for name
name_label.grid(column = 1, row = 2, padx = 10, pady = 5)		# label of name

level_label.grid(column = 3, row = 2, pady = 5)				# label of level
level_option.grid(column = 4, row = 2, pady = 5)			# level choose
level_option.current()							# default value is 1st-one

toWho_label.grid(column = 1, row = 3, pady = 5)				# label of person
toWho_option.grid(column = 2, row = 3, padx = 10, pady = 5)		# person choose
toWho_option.current()							# default option is 1st-one

direct_entry.grid(column = 4, row = 3, pady = 5)			# space for direct
direct_label.grid(column = 3, row = 3, padx = 10, pady = 5)		# label of direct

title_entry.grid(column = 2, row = 4, pady = 5)				# space for title
title_label.grid(column = 1, row = 4, pady = 5)				# label of title

topic_entry.grid(column = 4, row = 4, pady = 5)				# space for topic
topic_label.grid(column = 3, row = 4, pady = 5)				# label of topic

add_entry.grid(column = 2, row = 5, pady = 5)				# space for addjustment
add_label.grid(column = 1, row = 5, pady = 5)				# label of addjustment
add_button.grid(column = 2, row = 6, pady = 5)				# button that adds next addjustment

accept_button.grid(column = 5, row = 6, pady = 5)			# button 'Akceptuj'

frameWin.pack()								# main frame of main window
win.config(menu = menu_obj)						# show toolbar
win.mainloop()								# program loop



#
#   Module responisible for tool_bar settings and implimentation
#

from tkinter import *
from tkinter import messagebox
from tkinter.ttk import *
from applications import application_urlop, application_reward, application_one, application_hdk, application_boots, application_dutyChange_kmp
from applications import docx_command_urlop, docx_command_reward, docx_command_one, docx_command_hdk, docx_command_boots, docx_command_dutyChange_kmp
from docx_lib import docx_urlop, docx_reward, docx_one, docx_hdk, docx_boots, docx_dutyChange_kmp
from logic_n_alert import application_choose

#
font_name = "arial"
font_size20 = 20

def toolbar_settings(window, choose):

    # toolbar functions
    def fileNew(event = ""):                                                                                                # new formula

        if choose == "urlop okolicznościowy":

            window.destroy()
            # new window configuration
            window1 = Tk()
            window1.geometry("900x1100+800+0")
            window1.title("Generator")

            application_header = Label(window1, text = "Wniosek o urlop okolicznościowy", font = (font_name, font_size20)).pack()
            scroll = Scrollbar(window1)
            scroll.pack(side = "right", fill = Y)
            toolbar_settings(window1, choose)
            application_urlop(window1)

        elif choose == "urlop nagrodowy":

            window.destroy()
            # new window configuration
            window1 = Tk()
            window1.geometry("900x1100+800+0")
            window1.title("Generator")

            application_header = Label(window1, text = "Wniosek o urlop nagrodowy", font = (font_name, font_size20)).pack()
            scroll = Scrollbar(window1)
            scroll.pack(side = "right", fill = Y)
            toolbar_settings(window1, choose)
            application_reward(window1)

        elif choose == "przepustka jednorazowa":

            window.destroy()
            # new window configuration
            window1 = Tk()
            window1.geometry("900x1100+800+0")
            window1.title("Generator")

            application_header = Label(window1, text = "Wniosek o przepustkę jednorazową", font = (font_name, font_size20)).pack()
            scroll = Scrollbar(window1)
            scroll.pack(side = "right", fill = Y)
            toolbar_settings(window1, choose)
            application_one(window1)

        elif choose == "HDK":

            window.destroy()
            # new window configuration
            window1 = Tk()
            window1.geometry("900x1100+800+0")
            window1.title("Generator")

            application_header = Label(window1, text = "Wniosek o HDK", font = (font_name, font_size20)).pack()
            scroll = Scrollbar(window1)
            scroll.pack(side = "right", fill = Y)
            toolbar_settings(window1, choose)
            application_hdk(window1)

        elif choose == "buty wojskowe":

            window.destroy()
            # new window configuration
            window1 = Tk()
            window1.geometry("900x1100+800+0")
            window1.title("Generator")

            application_header = Label(window1, text = "Wniosek o buty", font = (font_name, font_size20)).pack()
            scroll = Scrollbar(window1)
            scroll.pack(side = "right", fill = Y)
            toolbar_settings(window1, choose)
            application_boots(window1)
            
        elif choose == "zmianę służby na kompanii":

            window.destroy()
            # new window configuration
            window1 = Tk()
            window1.geometry("900x1100+800+0")
            window1.title("Generator")

            application_header = Label(window1, text = "Wniosek o zmianę służby na kompanii", font = (font_name, font_size20)).pack()
            scroll = Scrollbar(window1)
            scroll.pack(side = "right", fill = Y)
            toolbar_settings(window1, choose)
            application_dutyChange_kmp(window1)

    def fileSave(event = ""):                                                                                               # save/print formula
        if choose == "urlop okolicznościowy":
            docx_command_urlop()
        elif choose == "urlop nagrodowy":
            docx_command_reward()
        elif choose == "przepustka jednorazowa":
            docx_command_one()
        elif choose == "HDK(nie działa)":
            docx_command_hdk()
        elif choose == "buty wojskowe":
            docx_command_boots()
        elif choose == "zmianę służby na kompanii":
            docx_command_dutyChange_kmp()
        else:
            messagebox.showinfo("NoneAppFonund", "Nie wybrano odpowiedniej aplikacji.")
            sys.exit()

    def copyText(event = ""):                                                                                               # copy text
        messagebox.showinfo("CopyFile", "not working")
        '''global selected
        if event:
            window.clipboard_get()
        if my_text.selection_get():
            selected = my_text.selection_get()
            window.clipboard_clear()
            window.clipboard_append(selected)'''

    def cutText(event = ""):                                                                                                # cut text
        messagebox.showinfo("CutFile", "not working")
        '''global selected
        global my_text
        if event:
            selected = window.clipboard_get()
        else:
            if my_text.selection_get():
                selected = my_text.selection_get()
                my_text.delete("sel.first", "sel.last")
                window.clipboard_clear()
                window_clipboard_append(selected)'''

    def pasteText(event = ""):                                                                                              # paste text
        messagebox.showinfo("PasteFile", "not working")
        '''global selected
        global my_text
        if event:
            selected = window.clipboard_get()
        else:
            if selected:
                position = my_text.index(INSERT)
                my_text.insert(position, selected)'''

    # shortcuts
    def about(event = ""):                                                                                                  # information about program
        messagebox.showinfo("about", "not working")

    def manual(event = ""):                                                                                                 # show manual
        messagebox.showinfo("Manual", "not working")

    def delete(event = ""):                                                                                             # exit shortcut
        window.destroy()
    window.bind("<Escape>", delete)

    def newfile(event = ""):                                                                                            # new file shortcut
        fileNew()
    window.bind("<Control-n>", newfile)

    def savefile(event = ""):                                                                                           # save file shortcut
        fileSave()
    window.bind("<Control-s>", savefile)

    def showabout(event = ""):                                                                                          # help shortcut
        messagebox.showinfo("about", "not working")
    window.bind("<Control-h>", showabout)

    def showmanual(event = ""):                                                                                         # manual shortcut
        messagebox.showinfo("Manual", "not working")
    window.bind("<Control-m>", showmanual)

    # toolbar configuration
    menu_obj = Menu(window)																								# menu object

    file = Menu(menu_obj, tearoff = 0)																					# 'Plik' option (variable)
    menu_obj.add_cascade(label = "Plik", menu = file)																	# 'Plik' option creation
    file.add_command(label = "Nowy  Ctrl+N", command = lambda : fileNew(False))				     						# 'Nowy' command
    file.add_command(label = "Drukuj Ctrl+S", command = lambda : fileSave(False))										# 'Drukuj' command
    file.add_separator()
    file.add_command(label = "Wyjdź  ESC", command = window.destroy)							     					# 'Wyjdź' command

    edit = Menu(menu_obj, tearoff = 0)																					# 'Zmień' option (variable)
    menu_obj.add_cascade(label = "Zmień", menu = edit)																	# 'Zmień' option creation
    edit.add_command(label = "Kopiuj Ctrl+C", command = lambda : copyText(False))										# 'Kopiuj' command
    edit.add_command(label = "Wytnij Ctrl+X", command = lambda : cutText(False))		   								# 'Wytnij' command
    edit.add_command(label = "Wklej  Ctrl+V", command = lambda : pasteText(False))				       					# 'Wklej' command

    help = Menu(menu_obj, tearoff = 0)																					# 'Pomoc' option (variable)
    menu_obj.add_cascade(label = "Pomoc", menu = help)																	# 'Pomoc' option creation
    help.add_command(label = "O programie Ctrl+H", command = lambda : about(False))										# 'O programie' command
    help.add_separator()
    help.add_command(label = "Instrukcja     Ctrl+M", command = lambda : manual(False))	    						  	# 'Instrukcja' command


    window.config(menu = menu_obj)																						# show toolbar

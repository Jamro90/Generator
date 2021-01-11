from tkinter import *
from tkinter import messagebox
win = Tk()

ent = StringVar()
entry = Entry(win, textvariable = ent).pack()


print(ent.get())

def click_me():
    if ((str(ent.get())[-1] == ".") or (str(ent.get())[-1] == " ")) :
        messagebox.showinfo("działa!", str(ent.get())[0:-1]+"1")
        print(ent.get()[0:-1]+"1")
    else:
        messagebox.showinfo("działa", str(ent.get()))
        print(ent.get())
but = Button(win, text = "click", command = click_me).pack()

win.mainloop()

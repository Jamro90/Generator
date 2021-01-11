from tkinter import *
from tkinter.ttk import Progressbar
root = Tk()
root.geometry("700x500")

pro = Progressbar(root, orient = "horizontal", mode = "determinate", length = 300)


pro.pack(pady = 100)
root.mainloop()

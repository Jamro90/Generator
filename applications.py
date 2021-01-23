#
#   Module responsible for handeling form of applications
#
from tkinter import StringVar, Label, Entry, Frame, Button, Radiobutton, IntVar
from tkinter.ttk import Combobox, Checkbutton
from date_lib import *
from docx_lib import docx_urlop, docx_reward, docx_one, docx_hdk, docx_boots, docx_dutyChange_kmp
from tkinter import filedialog, simpledialog, messagebox

# variables

rektor = "pułkownik Wachulak"

# alerts
def dateChangeButton():
    box = messagebox.showinfo("Err:404","I'll be back!")

def timeTravel(day1, month1, year1, day2, month2, year2):
    if(((day1 > day2) and (month1 == month2) and (year1 == year2)) or ((month1 > month2) and (year1 == year2)) or ((month2 > month1) and (year1 > year2)) or (year1 > year2)):
        answer = messagebox.showerror("Error:TimeTravel", "Nie możesz podawać daty w przeszłość!")
        print(1/0)
    else:
        pass

def timeLimit(day1, month1, year1, day2, month2, year2): # ---> WORK IN PROGRESS <--- #
    if((((day2 - day1) > 3) and (month1 == month2) and (year1 == year2)) or (((day1 == 29 and day2 > 1) and (day1 == 30 and day2 > 2)) and (month1 == 4 or month1 == 6 or month1 == 9 or month1 == 11) and (year1 == year2)) or (((day1 == 30 and day2 > 1) or (day1 == 31 and day2 > 2)) and (month1 == 1 or month1 == 3 or month1 == 5 or month1 == 7 or month1 == 8 or month1 == 10) and (year1 == year2))):
        answer = messagebox.showerror("Error:TimeLimit", "Przepustki jednorazowe są wydawane do 72 godzin!")
        print(1/0)
    else:
        pass

def timeOrder(day1, month1, year1, day2, month2, year2):
    if(isinstance(int(day1), int) and isinstance(int(day2), int) and isinstance(int(month1), int) and isinstance(int(month2), int) and isinstance(int(year1), int) and isinstance(int(year2), int)):
        pass
    else:
        answer = messagebox.showerror("Error:TimeOrder", "Daty podaje się jako liczby naturalne!")
        print(1/0)

def checkCalendar(day, month, year):
    x = int(year)%4
    if((int(day) > 31) and (int(month) == 1 or int(month) == 3 or int(month) == 5 or int(month) == 7 or int(month) == 8 or int(month) == 10 or int(month) == 12)):
        answer = messagebox.showerror("Error:CheckClendar", "Te miesiące nie mają więcej niż 31 dni!")
        print(1/0)
    if((int(day) > 30) and (int(month) == 4 or int(month) == 6 or int(month) == 9 or int(month) == 11)):
        answer = messagebox.showerror("Error:CheckCalendar", "Te miesiące nie mają więcej niż 30 dni!")
        print(1/0)
    if((int(day) > 29) and (int(month) == 2) and (x == 0)):
        answer = messagebox.showerror("Error:CheckCalendar", "Luty ma w tym roku 29 dni!")
        print(1/0)
    if((int(day) > 28) and (int(month) == 2) and (x != 0)):
        answer = messagebox.showerror("Error:CheckCalendar", "Luty ma 28 dni!")
        print(1/0)
    else:
        pass
# code chuncks
def default_application(window, width_size = 40, font_name = "arial", font_size = 12):
    # default settings
    frameDefault = Frame(window)
    nullLabel = Label(frameDefault, text = "  ")
    # frame settings
    global level
    level_var = StringVar()                                                                                                     # level variable
    level_label = Label(frameDefault, text = "stopień", font = (font_name, font_size))                                          # label "stopień"
    level = Combobox(frameDefault, width = width_size + 4, textvariable = level_var)                                            # cobobox object (level)
    level["value"] = ("szeregowy", "starszy szeregowy", "kapral", "starszy kapral", "plutonowy", "sierżant", "starszy sierżant")# combobox values (level)
    level.current(0)
    # if statment for transform level in short

    date_label = Label(frameDefault, text = "Data:", font = (font_name, font_size))                                             # label "Data:"
    date_output = Label(frameDefault, text = date_read(), font = (font_name, font_size))                                        # label of output of local host date
    date_button = Button(frameDefault, text = "...", font = (font_name, font_size), command = dateChangeButton)                 # date change button

    global name_var
    name_var = StringVar()                                                                                                      # name variable
    name_label = Label(frameDefault, text = "Imię: ", font = (font_name, font_size))                                            # label "Imię:"
    global name_entry
    name_entry = Entry(frameDefault, width = width_size, textvariable = name_var, font = (font_name, font_size))                # name entry pole

    global surname_var
    surname_var = StringVar()                                                                                                   # surname variable
    surname_label = Label(frameDefault, text = "Nazwisko: ", font = (font_name, font_size))                                     # label "Nazwisko:"
    global surname_entry
    surname_entry = Entry(frameDefault, width = width_size, textvariable = surname_var, font = (font_name, font_size))          # surname entry pole

    global group
    group = StringVar()                                                                                                         # group variable
    group_label = Label(frameDefault, text = "Grupa: ", font = (font_name, font_size))                                          # label "Grupa:"
    global group_entry
    group_entry = Entry(frameDefault, width = width_size, textvariable = group, font = (font_name, font_size))                  # group entry pole

    global where
    where = StringVar()                                                                                                         # where variable
    where_label = Label(frameDefault, text = "Miejscowość: ", font = (font_name, font_size))                                    # label "Miejscowość: "
    global where_entry
    where_entry = Entry(frameDefault, width = width_size, textvariable = where, font = (font_name, font_size))                  # where entry pole

    # display
    nullLabel.grid(column = 0, row = 0, padx = 70)
    level_label.grid(column = 1, row = 1, padx = 10, pady = 5)                                                                  # level label position
    level.grid(column = 2, row = 1, padx = 10, pady = 5)                                                                        # level bar position
    name_label.grid(column = 1, row = 2, padx = 10, pady = 5)                                                                   # name label position
    name_entry.grid(column = 2, row = 2, padx = 10, pady = 5)                                                                   # name entry position
    surname_label.grid(column = 1, row = 3, padx = 10, pady = 5)                                                                # surname label position
    surname_entry.grid(column = 2, row = 3, padx = 10, pady = 5)                                                                # surname entry position
    where_label.grid(column = 1, row = 4, padx = 10, pady = 5)                                                                  # where label position
    where_entry.grid(column = 2, row = 4, padx = 10, pady = 5)                                                                  # where entry position
    date_label.grid(column = 1, row = 5, padx = 10, pady = 5)                                                                   # date label position
    date_output.grid(column = 2, row = 5, padx = 10, pady = 5)                                                                  # date label output position
    date_button.grid(column = 3, row = 5, padx = 10, pady = 5)                                                                  # date button position
    group_label.grid(column = 1, row = 6, padx = 10, pady = 5)                                                                  # group label position
    group_entry.grid(column = 2, row = 6, padx = 10, pady = 5)                                                                  # group entry position
    frameDefault.pack(side = "top")                                         # frame position

def levelCheck(level):
    if level == "szeregowy":
        level_get = "szer. " 		                 																	# adding a run
    if level == "starszy szeregowy":
        level_get = "st. szer. "					                      												# adding a run
    if level == "kapral":
        level_get = "kpr. " 											                 								# adding a run
    if level == "starszy kapral":
        level_get = "st. kpr. " 														              					# adding a run
    if level == "plutonowy":
        level_get = "plut. " 																		                 	# adding a run
    if level == "sierżant":
        level_get = "sier. " 					                 														# adding a run
    if level == "starszy sierżant":
        level_get = "st. sier. " 								                   										# adding a run
    return level_get

def destination_application(window, width_size = 10, font_name = "arial", font_size = 12):
    # settings for lader's destinetion
    frameDestination = Frame(window)
    # frame settings
    global plat_var
    plat_var = StringVar()                                                                                                      # level variable
    plat_label = Label(frameDestination, text = "pluton: ", font = (font_name, font_size))                                      # label "pluton:"
    global plat
    plat = Combobox(frameDestination, width = width_size + 4, textvariable = plat_var)                                          # cobobox object (plat)
    plat["value"] = ("1", "2", "3", "4", "5", "6")                                                                              # combobox values (plat)
    plat.current(0)

    global kmp_var
    kmp_var = StringVar()                                                                                                       # level variable
    kmp_label = Label(frameDestination, text = "kompania: ", font = (font_name, font_size))                                     # label "kompania:"
    global kmp
    kmp = Combobox(frameDestination, width = width_size + 4, textvariable = kmp_var)                                            # cobobox object (kmp)
    kmp["value"] = ("1", "2", "3", "4", "5", "6", "7", "8", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20")   # combobox values (kmp)
    kmp.current(0)

    # display
    plat_label.grid(column = 1, row = 1, padx = 10, pady = 5)                                                                   # plat label position
    plat.grid(column = 2, row = 1, padx = 10, pady = 5)                                                                         # plat combobx position
    plat.current()
    kmp_label.grid(column = 1, row = 2, padx = 10, pady = 5)                                                                    # kmp label position
    kmp.grid(column = 2, row = 2, padx = 10, pady = 5)                                                                          # kmp combobox position
    kmp.current()
    frameDestination.pack()                                     # frame position

def date_application_single(window, text, width_size = 10, font_name = "arial", font_size = 12):
    # settings for date range input
    frameData = Frame(window)                                                                                                   # indemendent frame for date
    #frame settings
    date_text_label = Label(frameData, text = text, font = (font_name, font_size))                                              # label text

    day_vars = StringVar()                                                                                                      # day variable
    day_labels = Label(frameData, text = "dzień", font = (font_name, font_size))                                                # label "dzień"
    global days
    days = Combobox(frameData, width = width_size + 4, textvariable = day_vars)                                                 # cobobox object (day)
    days["value"] = ("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31")   # combobox values (day)
    days.current(0)

    month_vars = StringVar()                                                                                                    # month variable
    month_labels = Label(frameData, text = "miesiąc", font = (font_name, font_size))                                            # label "miesiąc"
    global months
    months = Combobox(frameData, width = width_size + 4, textvariable = month_vars)                                             # cobobox object (month)
    months["value"] = ("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")                                  # combobox values (month)
    months.current(0)

    year_vars = StringVar()                                                                                                     # year variable
    year_labels = Label(frameData, text = "rok", font = (font_name, font_size))                                                 # label "rok"
    global years
    years = Combobox(frameData, width = width_size + 4, textvariable = year_vars)                                               # cobobox object (next year)
    years["value"] = (str(datetime.now().year), str(datetime.now().year + 1))                                                   # combobox values (next year)
    years.current(0)

    # display
    date_text_label.grid(column = 5, row = 6, padx = 10, pady = 5)                                                              # date(text) label position
    day_labels.grid(column = 4, row = 7, padx = 10, pady = 5)                                                                   # day label position
    days.grid(column = 4, row = 8, padx = 10, pady = 5)                                                                         # day combobox position
    days.current()
    month_labels.grid(column = 5, row = 7, padx = 10, pady = 5)                                                                 # month label position
    months.grid(column = 5, row = 8, padx = 10, pady = 5)                                                                       # month combobox position
    months.current()
    year_labels.grid(column = 6, row = 7, padx = 10, pady = 5)                                                                  # year label position
    years.grid(column = 6, row = 8, padx = 10, pady = 5)                                                                        # year entry position
    frameData.pack()

def date_application(window, width_size = 10, font_name = "arial", font_size = 12):
    # settings for date range input
    frameData = Frame(window)                                                                                                   # indemendent frame for date
    #frame settings
    date_text_label = Label(frameData, text = "Okres", font = (font_name, font_size))                                           # label "data: "

    text_label = Label(frameData, text = "\tod:\t", font = (font_name, font_size))                                              # label "od"

    global day_var
    day_var = StringVar()                                                                                                       # day variable
    day_label = Label(frameData, text = "dzień", font = (font_name, font_size))                                                 # label "dzień"
    global day
    day = Combobox(frameData, width = width_size + 4, textvariable = day_var)                                                   # cobobox object (day)
    day["value"] = ("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31")   # combobox values (day)
    day.current(0)

    global month_var
    month_var = StringVar()                                                                                                     # month variable
    month_label = Label(frameData, text = "miesiąc", font = (font_name, font_size))                                             # label "miesiąc"
    global month
    month = Combobox(frameData, width = width_size + 4, textvariable = month_var)                                               # cobobox object (month)
    month["value"] = ("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")                                   # combobox values (month)
    month.current(0)

    global year_var
    year_var = StringVar()                                                                                                      # year variable
    year_label = Label(frameData, text = "rok", font = (font_name, font_size))                                                  # label "rok"
    global year
    year = Combobox(frameData, width = width_size + 4, textvariable = year_var)                                                 # cobobox object (next year)
    year["value"] = (str(datetime.now().year), str(datetime.now().year + 1))                                                    # combobox values (next year)
    year.current(0)

    text_label2 = Label(frameData, text = "\tdo:\t", font = (font_name, font_size))                                             # label "do"

    global day_var2
    day_var2 = StringVar()                                                                                                      # next day variable
    day_label2 = Label(frameData, text = "dzień", font = (font_name, font_size))                                                # label "dzień"
    global day2
    day2 = Combobox(frameData, width = width_size + 4, textvariable = day_var2)                                                 # cobobox object (next day)
    day2["value"] = ("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31")   # combobox values (next day)
    day2.current(0)

    global month_var2
    month_var2 = StringVar()                                                                                                    # next month variable
    month_label2 = Label(frameData, text = "miesiąc", font = (font_name, font_size))                                            # label "miesiąc"
    global month2
    month2 = Combobox(frameData, width = width_size + 4, textvariable = month_var2)                                             # cobobox object (next month)
    month2["value"] = ("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")                                  # combobox values (next month)
    month2.current(0)

    global year_var2
    year_var2 = StringVar()                                                                                                     # next year variable
    year_label2 = Label(frameData, text = "rok", font = (font_name, font_size))                                                 # label "rok"
    global year2
    year2 = Combobox(frameData, width = width_size + 4, textvariable = year_var2)                                               # cobobox object (next year)
    year2["value"] = (str(datetime.now().year), str(datetime.now().year + 1))                                                   # combobox values (next year)
    year2.current(0)

    # display
    date_text_label.grid(column = 5, row = 6, padx = 10, pady = 5)                                                              # date(text) label position
    text_label.grid(column = 3, row = 7, padx = 10, pady = 5)                                                                   # text("od") label position
    day_label.grid(column = 4, row = 7, padx = 10, pady = 5)                                                                    # day label position
    day.grid(column = 4, row = 8, padx = 10, pady = 5)                                                                          # day combobox position
    day.current()
    month_label.grid(column = 5, row = 7, padx = 10, pady = 5)                                                                  # month label position
    month.grid(column = 5, row = 8, padx = 10, pady = 5)                                                                        # month combobox position
    month.current()
    year_label.grid(column = 6, row = 7, padx = 10, pady = 5)                                                                   # year label position
    year.grid(column = 6, row = 8, padx = 10, pady = 5)                                                                         # year entry position
    text_label2.grid(column = 3, row = 9, padx = 10, pady = 5)                                                                  # text("do") label position
    day_label2.grid(column = 4, row = 9, padx = 10, pady = 5)                                                                   # next day label position
    day2.grid(column = 4, row = 10, padx = 10, pady = 5)                                                                        # next day entry position
    month_label2.grid(column = 5, row = 9, padx = 10, pady = 5)                                                                 # next month label position
    month2.grid(column = 5, row = 10, padx = 10, pady = 5)                                                                      # next month entry position
    year_label2.grid(column = 6, row = 9, padx = 10, pady = 5)                                                                  # next year label position
    year2.grid(column = 6, row = 10, padx = 10, pady = 5)                                                                       # next year entry position
    frameData.pack()                                            # frame position

def extra_information(window, width_size = 40, font_name = "arial", font_size = 12):
    # settings for additional informations

    frameInfo = Frame(window)                                                                                                   # frame for extra informations

    # frame settings
    nullLabel = Label(frameInfo, text = " ", font = (font_name, font_size))
    punishment_label = Label(frameInfo, text = "Kary dyscyplinarne: ", font = (font_name, font_size))

    global punishment
    punish_stat = StringVar()                                                                                                   # punishment variable
    punishment = Combobox(frameInfo, width = width_size + 4, textvariable = punish_stat)                                        # cobobox object (punishment)
    punishment["value"] = ("brak", "nagana", "5/30", "10/30")                                                                   # combobox values (punishment)
    punishment.current(0)

    back = StringVar()                                                                                                          # back variable
    back_label = Label(frameInfo, text = "Posiadam zaległości z: ", font = (font_name, font_size))                              # label "Posiadam zaległości z: "
    global back_entry
    back_entry = Entry(frameInfo, width = width_size, textvariable = back, font = (font_name, font_size))                       # back entry pole
    hint_label = Label(frameInfo, text = "(przedmiot i semestr)")                                                               # label "(przedmiot i semestr)"

    # display
    punishment_label.grid(column = 1, row = 1, padx = 10, pady = 5)                                                             # punishment label position
    punishment.grid(column = 2, row = 1, padx = 10, pady = 5)                                                                   # punishment Combobox position
    nullLabel.grid(column = 3, row = 1, padx = 0, pady = 5)
    back_label.grid(column = 1, row = 13, padx = 30, pady = 5)                                                                  # back label position
    back_entry.grid(column = 2, row = 13, padx = 5, pady = 5)                                                                   # back entry position
    hint_label.grid(column = 1, row = 14, padx = 5, pady = 1)                                                                   # hint label position
    frameInfo.pack()

def add_application(window, width_size = 40, font_name = "arial", font_size = 12):

    frameAdd = Frame(window)

    addLabel = Label(frameAdd, text = "\t" + 8 * " " + "Załączniki:", font = (font_name, font_size))
    add1_label = Label(frameAdd, text = "1)", font = (font_name, font_size))
    add2_label = Label(frameAdd, text = "2)", font = (font_name, font_size))
    add3_label = Label(frameAdd, text = "3)", font = (font_name, font_size))

    add1 = StringVar()
    global add1_entry
    add1_entry = Entry(frameAdd, width = width_size, textvariable = add1, font = (font_name, font_size))

    add2 = StringVar()
    global add2_entry
    add2_entry = Entry(frameAdd, width = width_size, textvariable = add2, font = (font_name, font_size))

    add3 = StringVar()
    global add3_entry
    add3_entry = Entry(frameAdd, width = width_size, textvariable = add3, font = (font_name, font_size))

    # display
    addLabel.grid(column = 0, row = 1, padx = 5, pady = 10)
    add1_label.grid(column = 1, row = 1, padx = 5, pady = 5)
    add1_entry.grid(column = 2, row = 1, padx = 5, pady = 5)
    add2_label.grid(column = 1, row = 2, padx = 5, pady = 5)
    add2_entry.grid(column = 2, row = 2, padx = 5, pady = 5)
    add3_label.grid(column = 1, row = 3, padx = 5, pady = 5)
    add3_entry.grid(column = 2, row = 3, padx = 5, pady = 5)
    frameAdd.pack()
# code repearing text values

def caseAndDotter(sentence):                                                                                                    # function making first letter uncapital & erasing dot at the end
    if((str(sentence))[-1] == "."):
        print(sentence[0:-1])
        sentence = sentence[0:-1]
    if((str(sentence[0]) == str(sentence[0]).upper())):
        sentence = str(sentence[0]).lower() + sentence[1:]
    return sentence

def dotter(sentence):                                                                                                           # function erasing dot at the end
    try:
        if((str(sentence))[-1] == "."):
            sentence = sentence[0:-1]
    except:
        messagebox.showinfo("DataMissing", "Brakuje danych w motywacji.")

def firstCapital(sentence):                                                                                                     # function making first letter capital
    try:                                                                                                     # function making first letter as capital
        if((str(sentence[0]) == str(sentence[0]).lower())):
            sentence = str(sentence[0]).upper() + sentence[1:]
    except:
        messagebox.showinfo("DataMissing", "Brakuje danych w imię, miejscowość lub miejsce.")
    return sentence

# applications

def docx_command_urlop():                                                                                                       # function printing "urlop okolicznościowy"
    # errors
    timeTravel(day.get(), month.get(), year.get(), day2.get(), month2.get(), year2.get())
    checkCalendar(day.get(), month.get(), year.get())
    checkCalendar(day2.get(), month2.get(), year2.get())
    timeOrder(day.get(), month.get(), year.get(), day2.get(), month2.get(), year2.get())

    level_take = levelCheck(level.get())
    docx_urlop(level_take + "pchor.", firstCapital(name_entry.get()), surname_entry.get(), firstCapital(where_entry.get()), date_read(), group_entry.get() + ", ",  plat.get() + "pl" + "/" + kmp.get() + "kp", rektor, day.get() + "." + month.get() + "." + year.get() + " r.", day2.get() + "." + month2.get() + "." + year2.get() + " r.", firstCapital(place_entry.get()), punishment.get(), back_entry.get(), caseAndDotter(mot_entry.get()), add1_entry.get(), add2_entry.get(), add3_entry.get())

def docx_command_reward():                                                                                                      # function printing "urlop nagrodowy"
        # errors
    timeTravel(day.get(), month.get(), year.get(), day2.get(), month2.get(), year2.get())
    timeTravel(days.get(), months.get(), years.get(), day.get(), month.get(), year.get())
    checkCalendar(day.get(), month.get(), year.get())
    checkCalendar(day2.get(), month2.get(), year2.get())
    checkCalendar(days.get(), months.get(), years.get())
    timeOrder(day.get(), month.get(), year.get(), day2.get(), month2.get(), year2.get())

    level_take = levelCheck(level.get())
    docx_reward(level_take + "pchor. ", firstCapital(name_entry.get()), surname_entry.get(), firstCapital(where_entry.get()), date_read(), group_entry.get() + ", ", plat.get(), kmp.get(), rektor, day.get() + "." + month.get() + "." + year.get() + " r.", day2.get() + "." + month2.get() + "." + year2.get() + " r.", dotter(firstCapital(place_entry.get())), punishment.get(), back_entry.get(), nr_entry.get(), days.get() + "." + months.get() + "." + years.get(), firstCapital(add1_entry.get()), firstCapital(add2_entry.get()), firstCapital(add3_entry.get()))

def docx_command_one():                                                                                                         # function printing "przepustka jednorazowa"
    # Error handeling
    timeTravel(day.get(), month.get(), year.get(), day2.get(), month2.get(), year2.get())
    checkCalendar(day.get(), month.get(), year.get())
    checkCalendar(day2.get(), month2.get(), year2.get())
    timeOrder(day.get(), month.get(), year.get(), day2.get(), month2.get(), year2.get())

    level_take = levelCheck(level.get())
    docx_one(level_take + "pchor.", firstCapital(name_entry.get()), surname_entry.get(), firstCapital(where_entry.get()), date_read(), group_entry.get() + ", ",  plat.get() + "pl" + "/" + kmp.get() + "kp", kmp.get(), day.get() + "." + month.get() + "." + year.get() + " r.", day2.get() + "." + month2.get() + "." + year2.get() + " r.", dotter(firstCapital(place_entry.get())), punishment.get(), back_entry.get(), caseAndDotter(mot_entry.get()), firstCapital(add1_entry.get()), firstCapital(add2_entry.get()), firstCapital(add3_entry.get()))

def docx_command_hdk():	                                                                                                        # function printing "HDK"
    level_take = levelCheck(level.get())
    docx_hdk(level_take + "pchor.", name_entry.get(), surname_entry.get(), where_entry.get(), date_read(), group_entry.get(), plat.get() + "pl" + "/" + kmp.get() + "kp", rektor, day.get() + "." + month.get() + "." + year.get() + " r.", day2.get() + "." + month2.get() + "." + year2.get() + " r.", days.get() + "." + months.get() + "." + years.get() + "r.", place_entry.get(), "1.  Potwierdzenie oddania krwi.")

def docx_command_boots():                                                                                                       # function printing "buty wojskowe"
        # Error handeling
    checkCalendar(days.get(), months.get(), years.get())

    level_take = levelCheck(level.get())
    docx_boots(level_take + "pchor.", firstCapital(name_entry.get()), surname_entry.get(), firstCapital(where_entry.get()), date_read(), group_entry.get() + ", ",  plat.get() + "pl" + "/" + kmp.get() + "kp", kmp.get(), days.get() + "." + months.get() + "." + years.get() + " r.")

def docx_command_dutyChange_kmp():                                                                                              # function printing "zmianę służby na kompanii"
    checkCalendar(days.get(), months.get(), years.get())

    level_take = levelCheck(level.get())
    level_take2 = levelCheck(level2.get())
    docx_dutyChange_kmp(level_take + "pchor. ", firstCapital(name_entry.get()), surname_entry.get(), rang.get(), level_take2 + "pchor. ", firstCapital(name2_entry.get()), surname2_entry.get(), firstCapital(where_entry.get()), date_read(), days.get() + "." + months.get() + "." + years.get() + " r.", group_entry.get(), plat.get(), kmp.get(), caseAndDotter(mot_entry.get()))
    
def application_urlop(window, width_size = 40, font_name = "arial", font_size = 12):                                            # application for "urlop okolicznościowy"

    # additional settings for application
    frameUrlop = Frame(window)
    global place_entry
    place = StringVar()                                                                                                         # place variable
    place_label = Label(frameUrlop, text = "miejsce docelowe(miasto): ", font = (font_name, font_size))                         # label "gdzie: "
    place_entry = Entry(frameUrlop, width = width_size, textvariable = place, font = (font_name, font_size))                    # place entry pole

    global mot_entry
    mot = StringVar()                                                                                                           # mot variable
    mot_label = Label(frameUrlop, text = "Wniosek swój motywuje: ", font = (font_name, font_size))                              # label "motywacja: "
    mot_entry = Entry(frameUrlop, width = width_size, textvariable = mot, font = (font_name, font_size))                        # mot entry pole

    # display
    default_application(window)                                                                                                 # default application
    place_label.grid(column = 0, row = 1, padx = 10, pady = 5)                                                                  # place label position
    place_entry.grid(column = 1, row = 1, padx = 10, pady = 5)                                                                  # place entry position
    mot_label.grid(column = 0, row = 2, padx = 10, pady = 5)                                                                    # motivation label position
    mot_entry.grid(column = 1, row = 2, padx = 10, pady = 5)                                                                    # motivation entry position
    frameUrlop.pack()                                                                                                           # frame position
    destination_application(window)                                                                                             # destination application
    date_application(window)                                                                                                    # date application

    extra_information(window)                                                                                                   # extra info application
    add_application(window)

    bnt = Button(window, text = "Drukuj!", command = docx_command_urlop)                                                        # button "Drukuj!"
    bnt.pack()                                                                                                                  # button position

def application_reward(window, width_size = 40, font_name = "arial", font_size = 12):                                           # application for "urlop nagrodowy"

    frameReward = Frame(window)
    # display
    global place_entry
    place = StringVar()                                                                                                         # place variable
    place_label = Label(frameReward, text = "miejsce docelowe(miasto): ", font = (font_name, font_size))                        # label "gdzie: "
    place_entry = Entry(frameReward, width = width_size, textvariable = place, font = (font_name, font_size))                   # place entry pole

    default_application(window)                                                                                                 # default application
    global nr_entry
    nrLabel = Label(frameReward, text = "numer rozkazu:", font = (font_name, font_size))                                        # label "numer rozkazu:"
    nr = StringVar()
    nr_entry = Entry(frameReward, width = width_size, textvariable = nr, font = (font_name, font_size))                         # nr entry

    place_label.grid(column = 1, row = 1, padx = 10, pady = 5)                                                                  # place label position
    place_entry.grid(column = 2, row = 1, padx = 10, pady = 5)                                                                  # place entry position
    nrLabel.grid(column = 1, row = 2, padx = 5, pady = 10)
    nr_entry.grid(column = 2, row = 2, padx = 5, pady = 10)
    date_application_single(window, "Data rozkazu:")
    frameReward.pack()
    destination_application(window)                                                                                             # destination application
    date_application(window)                                                                                                    # date application
    extra_information(window)                                                                                                   # extra info application
    add_application(window)

    bnt = Button(window, text = "Drukuj!", command = docx_command_reward)                                                       # button "Drukuj!"
    bnt.pack()

def application_one(window, width_size = 40, font_name = "arial", font_size = 12):                                              # application for "przpeustka jednorazowa"

    frameOne = Frame(window)

    global place_entry
    place = StringVar()                                                                                                         # place variable
    place_label = Label(frameOne, text = "miejsce docelowe(miasto): ", font = (font_name, font_size))                           # label "gdzie: "
    place_entry = Entry(frameOne, width = width_size, textvariable = place, font = (font_name, font_size))                      # place entry pole

    global mot_entry
    mot = StringVar()                                                                                                           # mot variable
    mot_label = Label(frameOne, text = "Wniosek swój motywuje: ", font = (font_name, font_size))                                # label "motywacja: "
    mot_entry = Entry(frameOne, width = width_size, textvariable = mot, font = (font_name, font_size))                          # mot entry pole

    # display
    default_application(window)                                                                                                 # default application

    place_label.grid(column = 0, row = 1, padx = 10, pady = 5)                                                                  # place label position
    place_entry.grid(column = 1, row = 1, padx = 10, pady = 5)                                                                  # place entry position
    mot_label.grid(column = 0, row = 2, padx = 10, pady = 5)                                                                    # motivation label position
    mot_entry.grid(column = 1, row = 2, padx = 10, pady = 5)                                                                    # motivation entry position

    frameOne.pack()
    destination_application(window)                                                                                             # destination application
    date_application(window)                                                                                                    # date application
    extra_information(window)                                                                                                   # extra info application
    add_application(window)

    bnt = Button(window, text = "Drukuj!", command = docx_command_one)                                                          # button "Drukuj!"
    bnt.pack()                                                                                                                  # button position

def application_hdk(window, width_size = 40, font_name = "arial", font_size = 12):                                              # application for "książeczka wojskowa"

    frameHdk = Frame(window)
    # display
    default_application(window)                                                                                                 # default application

    global place_entry
    place = StringVar()                                                                                                         # place variable
    place_label = Label(frameHdk, text = "miejsce docelowe(miasto): ", font = (font_name, font_size))                           # label "gdzie: "
    place_entry = Entry(frameHdk, width = width_size, textvariable = place, font = (font_name, font_size))                      # place entry pole

    place_label.grid(column = 1, row = 0, padx = 5, pady = 10)
    place_entry.grid(column = 2, row = 0, padx = 5, pady = 10)

    destination_application(window)                                                                                             # destination application
    date_application(window)                                                                                                    # date application
    date_application_single(window, "Data donacji krwi:")
    extra_information(window)                                                                                                   # extra info application

    bnt = Button(window, text = "Drukuj!", command = docx_command_hdk)                                                          # button "Drukuj!"

    frameHdk.pack()
    bnt.pack()

def application_boots(window, width_size = 40, font_name = "arial", font_size = 12):                                            # application for "buty wojskowe"

    frameBoots = Frame(window)
    # display
    default_application(window)                                                                                                 # default application
    frameBoots.pack()
    destination_application(window)                                                                                             # destination application
    date_application_single(window, "Data zniszczenia obuwia:")                                                                 # date application

    bnt = Button(window, text = "Drukuj!", command = docx_command_boots)                                                        # button "Drukuj!"
    bnt.pack()

def application_dutyChange_kmp(window, width_size = 40, font_name = "arial", font_size = 12):                                   # application for "zmianę służby na kompanii"

    # additional settings for application
    frameDutyChange = Frame(window)

    global mot_entry
    mot = StringVar()                                                                                                           # mot variable
    mot_label = Label(frameDutyChange, text = "Wniosek swój motywuje: ", font = (font_name, font_size))                         # label "motywacja: "
    mot_entry = Entry(frameDutyChange, width = width_size, textvariable = mot, font = (font_name, font_size))                   # mot entry pole

    # next person data
    global level2
    level2_var = StringVar()                                                                                                     # level2 variable
    level2_label = Label(frameDutyChange, text = "stopień", font = (font_name, font_size))                                       # label2 "stopień"
    level2 = Combobox(frameDutyChange, width = width_size + 4, textvariable = level2_var)                                        # cobobox object (level2)
    level2["value"] = ("szeregowy", "starszy szeregowy", "kapral", "starszy kapral", "plutonowy", "sierżant", "starszy sierżant")# combobox values (level2)
    level2.current(0)
    
    change_label = Label(frameDutyChange, text = "Dane zmiennika")                                                              # label "Dane zmiennika"
    
    global name2_entry
    name2 = StringVar()                                                                                                         # name2 variable
    name2_label = Label(frameDutyChange, text = "Imię: ", font = (font_name, font_size))                                        # label "Imię: "
    name2_entry = Entry(frameDutyChange, width = width_size, textvariable = name2, font = (font_name, font_size))               # name entry pole
    
    global surname2_entry
    surname2 = StringVar()                                                                                                      # surname2 variable
    surname2_label = Label(frameDutyChange, text = "Nazwisko: ", font = (font_name, font_size))                                 # label "Nazwisko: "
    surname2_entry = Entry(frameDutyChange, width = width_size, textvariable = surname2, font = (font_name, font_size))         # surname2 entry pole

    global rang
    rang_var = StringVar()                                                                                                      # surname2 variable
    rang_label = Label(frameDutyChange, text = "Funkcja: ", font = (font_name, font_size))                                      # label "Nazwisko: "
    rang = Combobox(frameDutyChange, width = width_size + 4, textvariable = rang_var)                                           # cobobox object (level2)
    rang["value"] = ("Podoficer dyżurny", "I Dyżurny", "II Dyżurny")                                                            # combobox values (level2)
    rang.current(0)
    
    # display
    default_application(window)                                                                                                 # default application
    
    rang_label.grid(column = 0, row = 1, padx = 10, pady = 5)                                                                   # rang label position
    rang.grid(column = 1, row = 1, padx = 10, pady = 5)                                                                         # rang combobox position

    mot_label.grid(column = 0, row = 2, padx = 10, pady = 5)                                                                    # motivation label position
    mot_entry.grid(column = 1, row = 2, padx = 10, pady = 5)                                                                    # motivation entry position
 
    destination_application(window)                                                                                             # destination application
    date_application_single(window, "Data pełnienia służby", 10, font_name, font_size)
    change_label.grid(column = 1, row = 3, padx = 10, pady = 5)                                                                 # change label position
    
    level2_label.grid(column = 0, row = 5, padx = 10, pady = 5)                                                                 # level2 label position
    level2.grid(column = 1, row = 5, padx = 10, pady = 5)                                                                       # level2 label position
    name2_label.grid(column = 0, row = 15, padx = 10, pady = 5)                                                                 # name2 label position
    name2_entry.grid(column = 1, row = 15, padx = 10, pady = 5)                                                                 # name2 entry position
    
    surname2_label.grid(column = 0, row = 25, padx = 10, pady = 5)                                                              # surname2 label position
    surname2_entry.grid(column = 1, row = 25, padx = 10, pady = 5)                                                              # surname2 entry position

    frameDutyChange.pack()                                                                                                      # frame position

   
    bnt = Button(window, text = "Drukuj!", command = docx_command_dutyChange_kmp)                                               # button "Drukuj!"
    bnt.pack()                                                                                                                  # button position

#
#   Module responisible for making a date and animate changing process
#

from datetime import datetime

def date_read():                                # reading the date from local desktop
    date = datetime.now()						# import date from local host

    if(float(date.day) < 10):					# add 0 befor day less than 10
        day = "0" + str(date.day)
    else:
        day = str(date.day)
    if(int(date.month) < 10):					# add 0 befor month less than 10
        month = "0" + str(date.month)
    else:
        month = str(date.month)

    date_right = str(day + "." + month + "." + str(date.year) + " r.")	# date formula
    return date_right

def date_show():                                     # present a date in gui
    year_header = datetime.now().year
    months = ["Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec", "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"]
    week = ["pn", "wt", "śr", "cz", "pt", "so", "nd"]


def date_change():                                   # logic for changing date manualy
    pass

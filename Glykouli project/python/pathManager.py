import datetime

def getMonthString(number):
    if(number == 1):
        return "January"
    if(number == 2):
        return "February"
    if(number == 3):
        return "March"
    if(number == 4):
        return "April"
    if(number == 5):
        return "May"
    if(number == 6):
        return "June"
    if(number == 7):
        return "July"
    if(number == 8):
        return "August"
    if(number == 9):
        return "September"
    if(number == 10):
        return "October"
    if(number == 11):
        return "November"
    if(number == 12):
        return "December"

def getWeekDay(day):
    if(day == 0 or day == 7):
        return "Monday"
    if(day == 1):
        return "Tuesday"
    if(day == 2):
        return "Wednesday"
    if(day == 3):
        return "Thursday"
    if(day == 4):
        return "Friday"
    if(day == 5):
        return "Saturday"
    if(day == 6):
        return "Sunday"

def getNextMonth(month):
    if(month == 12):
        return 1
    else:
        return month + 1

def getPreviousMonth(month):
    if month == 1:
        return 12
    else:
        return month - 1

def getPreviousMonthPath():
    path = "excel/"
    month = datetime.date.today().month
    prevMonth = getPreviousMonth(month)
    date = datetime.date.today() - datetime.timedelta(days=30)
    return path + getMonthString(prevMonth) + str(date.year) + ".xlsx"

def getNextMonthPath():
    path = "excel/"
    month = datetime.date.today().month
    nextMonth = getNextMonth(month)
    date = datetime.date.today() + datetime.timedelta(days=30)
    return path + getMonthString(nextMonth) + str(date.year) + ".xlsx"

def getThisMonthPath():
    path = "excel/"
    month = datetime.date.today().month
    date = datetime.date.today()
    return path + getMonthString(month) + str(date.year) + ".xlsx"



import PyPDF2
import re
import datetime
import tkinter
from PyPDF2 import PdfReader
from icalendar import Calendar, Event
from tkinter import filedialog
from tkinter import messagebox as mbox

#to adjust this for any individual put in their last name as it is on schedule
Emp_Name = "CAVANNA"

root = tkinter.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()
text = str()

#makes code to split string by space delimiter
def split_string(string):
    list_string = string.split('\n')
    return list_string

# Function to convert string to datetime
def convert_time(date_time):
    format = '%m/%d/%Y%I:%M%p'  # The format
    datetime_str = datetime.datetime.strptime(date_time, format)
    return datetime_str

#initiates vars
cal = Calendar()
numOfDays = 0

#creates .ics files
def create_ics(start, end, task):
    event = Event()
    event.add('summary', 'REI Work: ' + task)
    event.add('dtstart', start)
    event.add('dtend', end)
    event.add('description', 'test test')
    cal.add_component(event)
    f = open('work_schedule.ics', 'wb')
    f.write(cal.to_ical())
    f.close()

#try:
    #opens file and extracts all text
    reader = PdfReader(file_path)
    number_of_pages = len(reader.pages)
    for i in range(0,number_of_pages):
        page = reader.pages[i]
        temp_text = page.extract_text()
        text = text +  temp_text

    #get dates
    ind  = text.index("Employee")
    ind2 = text.index("Weekly")
    dates = text[ind+9:ind2-1]
    #splits it up by returns (\n)
    dates_list = split_string(dates)

    #gets out my scedule
    ind  = text.index(Emp_Name)
    index_comma = text.index(", ",ind+15)
    index_page = text.index("Page",ind+15)
    if index_comma < index_page: ind2 = index_comma
    else: ind2 = index_page
    schedule = text[ind:ind2]

    #splits it up by returns (\n)
    strii = split_string(schedule)

    #reformats am/pm
    indx = 0
    for i in strii:
        if i == "A-" or i == "A": strii[indx] = "am"
        if i == "P-" or i == "P": strii[indx] = "pm"
        indx +=1

    indx = 0 #this is the index for the date
    iID = 0

    #finds the day schedules
    for i in strii:
        if indx == len(dates_list): indx = 0
        iID += 1
        if len(i) == 0: continue
        elif i == " ":  #checks if it is a day when i am not scheduled
                indx += 1 #adds a day to keep them lined up
                continue
                #makes sure its a time and that it is the first time of day
        elif i[0].isdigit() == True and\
                 i[3].isdigit() == True and\
                 iID+1 < len(strii) and\
                 strii[iID+1][0].isdigit() == True:
                indx += 1 #adds date index
                #creates date/time combinations
                startTm = convert_time(dates_list[indx-1][4:]\
                                       +strii[iID-1]  +strii[iID])
                endTm   = convert_time(dates_list[indx-1][4:]\
                                       +strii[iID+1]+strii[iID+2])
                task = strii[iID + 3]
                create_ics(startTm, endTm, task)
                numOfDays += 1

    complete_message = "You are scheduled to work "+str(numOfDays)+" days succkaahh!"
    mbox.showinfo("Heyyy beautiful",complete_message)

#this does error handeling, prob when you click a bad file
#except:mbox.showinfo("You idiot","A fucking error has occured you dipshit")





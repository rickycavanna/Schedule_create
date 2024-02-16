##import PyPDF2
##import pdfminer
import openpyxl
import datetime
from datetime import datetime
import tkinter as tk
##import csv
##from PyPDF2 import PdfReader
from icalendar import Calendar, Event
from tkinter import filedialog
from tkinter import messagebox as mbox
import pandas as pd
import numpy as np

#to adjust this for any individual put in their last name as it is on schedule
Emp_Name = "CAVANNA"
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

def load_xlsx(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = [list(row) for row in sheet.iter_rows(values_only=True)]
    return pd.DataFrame(data)


#initiates vars
cal = Calendar()
numOfDays = 0

def open_file():
    file_path = filedialog.askopenfilename(title="Select CSV File", filetypes=[("Excel files", "*.xlsx")])

    if file_path:
        if file_path.endswith('.xlsx'):
            Sch_data = load_xlsx(file_path)
        else:
            print("Unsupported file format.")
            return

    # Initialize the row index
    row = 0

    # Loop to drop rows until a non-None value is encountered in the first column
    while row < len(Sch_data) and Sch_data.iloc[row, 0] is None:
        Sch_data = Sch_data.drop(Sch_data.index[row]).reset_index(drop=True)

    # this expands the value of column down if it is none
    for i in range(1, len(Sch_data)):
        if Sch_data.iloc[i, 0] is None:
            Sch_data.iloc[i, 0] = Sch_data.iloc[i - 1, 0]

    # Identify rows containing 'Employee' or name in column 0
    employee_rows = Sch_data.iloc[:, 0].str.lower().eq('employee')
    comma_rows = Sch_data.iloc[:, 0].str.lower().str.contains(Emp_Name.lower())

    Sch_data = Sch_data[employee_rows | comma_rows]
    Sch_data = Sch_data.reset_index(drop=True)
    schdule = pd.DataFrame(columns=Sch_data.columns)

    #removes all the rows except the employee in question and dates
    emp_ind = Sch_data[Sch_data[0].str.lower().str.contains(Emp_Name.lower())].index
    tempp = min(emp_ind) - 1
    emp_ind = emp_ind.append(pd.Index([tempp]))

    Sch_data = Sch_data[Sch_data.index.isin(emp_ind)]

    #removes all columns where value is just "none"
    columns_to_drop = [col for col in Sch_data.columns if Sch_data[col][1:].isna().all()]
    Sch_data = Sch_data.drop(columns=columns_to_drop)
    Sch_data = Sch_data.reset_index(drop=True)

    # Print the entire DataFrame
    print(Sch_data)

##    # Write the DataFrame to a CSV file
##    csv_file_path = "C:/Users/Ecava/OneDrive/Desktop/ProgrammingStuff/rei_schedule/output.csv"
##    Sch_data.to_csv(csv_file_path, index=False)

# Create the main tkinter window
root = tk.Tk()
root.title("File Reader")

# Create and configure the Open File button
open_button = tk.Button(root, text="Open Schedule File", command=open_file)
open_button.pack(pady=20)

# Run the tkinter main loop
root.mainloop()





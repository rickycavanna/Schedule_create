#this takes the REI schedule and creates ICS files from it based on last name
import openpyxl
from datetime import datetime
import tkinter as tk
from icalendar import Calendar, Event
from tkinter import filedialog, messagebox as mbox, simpledialog as sd
import pandas as pd


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
    #Load data from an Excel file and return it as a DataFrame
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = [list(row) for row in sheet.iter_rows(values_only=True)]
    return pd.DataFrame(data)

#creates .ics files
def create_ics(df):
    cal = Calendar()

    for index, row in df.iterrows():
        event = Event()
        event.add('summary', 'REI Work: ' + row['Task'])
        event.add('dtstart', row['Start'].to_pydatetime())
        event.add('dtend', row['End'].to_pydatetime())
        event.add('description', 'test test')
        cal.add_component(event)

    return cal

def check_cols(df):
    sch_cols = []

    # Iterate over columns and check if the first value is in datetime format
    for col in df.columns:
        top_val = df[col].iloc[0]

        if isinstance(top_val, datetime):
            shift_start, shift_end, task = find_shift_times(df[col][1:])

            # Combine date part of 'Date' and time part of 'Shift Start' and 'Shift End'
            combined_start = pd.to_datetime(top_val.strftime('%Y-%m-%d') + ' ' + shift_start.strftime('%H:%M:%S'))
            combined_end = pd.to_datetime(top_val.strftime('%Y-%m-%d') + ' ' + shift_end.strftime('%H:%M:%S'))

            # Check if task is None or an empty string
            if task is None or task == '':
                # Use the first non-null value from the 'main_task' column
                task_com = str(df.iloc[:, 1].dropna().iloc[0])
            else:
                task_com = task

            # Assuming 'Shift Start', 'Shift End', 'Task' are column names in your DataFrame
            sch_cols.append([combined_start, combined_end, task_com])

    sch_df = pd.DataFrame(sch_cols, columns=['Start', 'End', 'Task'])
    return sch_df

def find_shift_times(column):
    # Extract shift time strings from the column
    shift_time_strs = [value for value in column if isinstance(value, str)]

    if not shift_time_strs:
        return None, None

    # Parse each shift time string and collect start and end times
    shift_times = [parse_shift_time(shift_time_str) for shift_time_str in shift_time_strs]

    # Filter out None values for start and end times
    tasks = [(start,end,task) for start,end,task in shift_times if task is not None]
    shift_times = [(start, end, task) for start, end, task in shift_times if start is not None and end is not None]
        
    if not shift_times:
        return None, None, None

    # Extract the earliest start time and latest end time
    earliest_start_time = min(shift_times, key=lambda x: x[0])[0]
    latest_end_time = max(shift_times, key=lambda x: x[1])[1]

    r_tasks = ', '.join(set(task for (_, _, task) in tasks))

    return earliest_start_time, latest_end_time, r_tasks


def parse_shift_time(shift_time_str):
    start_time = None
    end_time = None
    task = None

    # Split the shift time string into start and end time strings
    if ' - ' in shift_time_str:
        start_time_str, end_time_str = shift_time_str.split(' - ')

        try:
            # Parse the start and end times as datetime objects
            start_time = datetime.strptime(start_time_str, '%I:%M %p')
            end_time = datetime.strptime(end_time_str, '%I:%M %p')
        except ValueError:
            pass
    elif shift_time_str is not None:
        task = shift_time_str

    return start_time, end_time, task

def process_schedule(Sch_data, Emp_Name):
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

    ## below is the code that is causing problems when there is the wrong name
    
    #removes all the rows except the employee in question and dates
    emp_ind = Sch_data[Sch_data[0].str.lower().str.contains(Emp_Name.lower())].index
    tempp = min(emp_ind) - 1
    emp_ind = emp_ind.append(pd.Index([tempp]))

    Sch_data = Sch_data[Sch_data.index.isin(emp_ind)]

    #removes all columns where value is just "none"
    columns_to_drop = [col for col in Sch_data.columns if Sch_data[col][1:].isna().all()]
    Sch_data = Sch_data.drop(columns=columns_to_drop)
    Sch_data = Sch_data.reset_index(drop=True)

    return Sch_data

def save_ics_file(cal):
    with open('appointments.ics', 'wb') as f:
        f.write(cal.to_ical())
    mbox.showinfo("Success", "Appointments saved to 'appointments.ics'")

def open_file():
    #Open a file dialog to select an Excel file and process it

    Emp_Name = sd.askstring("Enter Employee Name", "Enter the Employee's Last Name:")

    if not Emp_Name:
        mbox.showerror("Error", "Last name cannot be empty.")
        return  # user clicked Cancel
    
    try:
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])

        if file_path:
            if file_path.endswith('.xlsx'):
                Sch_data = load_xlsx(file_path)
                Sch_data = process_schedule(Sch_data, Emp_Name)
                datetime_cols = check_cols(Sch_data)
                ical_calendar = create_ics(datetime_cols)
                save_ics_file(ical_calendar)

            else:
                raise TypeError("Unsupported Format", "Please select a valid Excel file, dickhead.")
        else:
            mbox.showinfo("Info", "No file selected. Exiting.")        
    except Exception as e:
        mbox.showerror("Error",f"An unexpected error occurred: {e}")#f"You did not upload the right file or some other error occured :(")

# Create and configure the Open File button
open_file()





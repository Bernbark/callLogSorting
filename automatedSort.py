from __future__ import print_function
import tkinter.tix as tix
import tkinter as tk
from tkinter import filedialog
import csv
import os
import getpass
from PIL import ImageTk, Image
from apscheduler.schedulers.blocking import BlockingScheduler

disconnected_filename = ''
voicemail_filename = ''
voicemail = 'Answering_Machine'
disconnected = 'Disconnected'
rows = []
household_id_found = False
household_id_filename = ''
call_logs_found = False
call_logs_filename = ''
automated_runtime_minutes = 0
automated_runtime_hour = 1
am_pm_runtime = "AM"


def cron_process():
    print ('periodic print')


scheduler = BlockingScheduler()
scheduler.add_job(cron_process, 'cron', day='*', hour='1', minute='57',second='0')
#scheduler.start()

"""
# If the number is disconnected then the script should, initially, 
# generate a Disconnected file in a seperate folder, 
# and if the file already exists append the new results to it.
"""

"""
Need folder to be made automatically 
"""

# Grabbing user name for creation of a Desktop folder
username = getpass.getuser()

# C:\Users\Kory Stennett\Desktop\CallReports is the format
path_str = "C:\\Users\\"+username+"\\Desktop\\CallReports"

# If the directory does not exist, we make it using the above path
if not os.path.isdir(path_str):
    os.makedirs(path_str)


# Check if CSV file exists and create if not, then add headers
# Headers are CallCompletedTimeStamp,PhoneNumberDialed,Response
# ExactBillingDurationInSeconds,RoundedBillingDurationInMinutes
def create_csv_files():
    # make sure we need a header by checking if we've already made this file


    if disconnected_filename == "":
        disconnected_file_path = path_str + "\\DisconnectedCallSheet.csv"
    else:
        disconnected_file_path = path_str + "\\"+disconnected_filename
        if disconnected_file_path == path_str+"\\"+".csv":
            disconnected_file_path = path_str + "\\DisconnectedCallSheet.csv"
    print(disconnected_file_path)
    file_exists = os.path.isfile(disconnected_file_path)
    try:
        with open(disconnected_file_path, 'a') as output_file:
            writer = csv.writer(output_file)
            # set header if file doesn't exist
            if not file_exists:
                writer.writerow(["Household ID","Phone Number","Call Timestamp"])
    except FileExistsError:
        pass
    print(disconnected_filename + " disconnected file name")

    if voicemail_filename == "":
        voicemail_filepath = path_str + "\\VoicemailCallSheet.csv"
    else:
        voicemail_filepath = path_str + "\\"+voicemail_filename
        if voicemail_filepath == path_str+"\\"+".csv":
            voicemail_filepath = path_str + "\\VoicemailCallSheet.csv"
    print(voicemail_filepath)
    file_exists = os.path.isfile(voicemail_filepath)
    try:
        with open(voicemail_filepath, 'a') as output_file:
            writer = csv.writer(output_file)
            # set header if file doesn't exist
            if not file_exists:
                writer.writerow(["Household ID","Phone Number","Call Timestamp"])
    except FileExistsError:
        pass
    print(voicemail_filename + " voicemail file name")


def getHouseholdID(household_ids,call_logs,file_path,target_word):
    id_rows = []
    count = 0
    id_index = 0
    with open(household_ids, 'r') as id_log:
        reader = csv.reader(id_log)
        d_reader = csv.DictReader(id_log)
        # setting the headers to a variable
        headers = d_reader.fieldnames
        for header in headers:
            if "ID" in header:
                break
            else:
                id_index+=1
        for line in reader:

            id_rows.append(line)

    phone_rows = []
    with open(call_logs, 'r') as call_log:
        # DictReader will help pull headers from spreadsheets
        d_reader = csv.DictReader(call_log)
        # setting the headers to a variable
        headers = d_reader.fieldnames
        reader = csv.reader(call_log)
        target_index = 0
        phone_index = 0
        date_index = 0
        for header in headers:
            if header == "Response":
                break
            else:
                target_index += 1
        for header in headers:
            if "Phone" in header:
                break
            else:
                phone_index += 1
        for header in headers:
            if "TimeStamp" in header:
                break
            else:
                date_index += 1
        date_list = []
        for line in reader:
            if target_word in line[target_index]:
                #print(line[date_index])



                rows.append((line[phone_index],line[date_index]))


        for entry in rows:
            for id in id_rows:
                if entry[0] == id[1]:
                    entry = (id[id_index],)+entry

                    with open(file_path, 'a+') as output_file:
                        writer = csv.writer(output_file)
                        writer.writerow(entry)

            count += 1

        rows.clear()
        id_rows.clear()

    #with open(file_path,'a') as output_file:

        #writer = csv.writer(output_file)
        #writer.writerows(phone_rows)
        """
        for row1, row2 in zip_longest(phone_rows, id_rows):
            try:
                if row1[1] == row2[1]:
                    writer.writerow(rows)
                    rows.clear()
                else:
                    pass
            except TypeError:
                pass
                """


# for use with the button to sort once files are selected
def get_all_household_id():
    global disconnected_filename
    disconnected_filename = onDisconnectedSaveEntry()

    global voicemail_filename
    voicemail_filename = onVoicemailSaveEntry()
    if voicemail_filename == "":
        voicemail_filename = "VoicemailCallSheet.csv"
    else:
        voicemail_filename = voicemail_filename + ".csv"
    if disconnected_filename == "":
        disconnected_filename = "DisconnectedCallSheet.csv"
    else:
        disconnected_filename = disconnected_filename + ".csv"
    create_csv_files()
    getHouseholdID(household_id_filename, call_logs_filename, path_str + "\\" + voicemail_filename, voicemail)
    getHouseholdID(household_id_filename, call_logs_filename, path_str + "\\" + disconnected_filename, disconnected)

def get_minutes():
    return minute_option_var.get()


def get_hour():
    night_or_day = am_pm_option_var.get()
    if "AM" in night_or_day:
        unfiltered_hour = hour_option_var.get()
        hour = int(unfiltered_hour.replace(":",""))
        if hour == 12:
            hour = 0
        print(hour)
        return hour
    else:
        unfiltered_hour = hour_option_var.get()
        hour = int(unfiltered_hour.replace(":", ""))
        if hour != 12:
            hour += 12
        print(hour)
        return hour


"""
Need to search for all successful answering machine calls and disconnected calls
"""

"""
# This function takes a string and searches for instances of that string in a
# csv file then fills another csv file with only the rows targeted
# Note that we are checking the 3rd column for "Answering_Machine" or "Disconnected"
def fill_call_sheet(target_word,filename):
    try:
        with open(filename) as spread_sheet:
            # DictReader will help pull headers from spreadsheets
            d_reader = csv.DictReader(spread_sheet)
            # setting the headers to a variable
            headers = d_reader.fieldnames
            # If it's not in the headers tell the user
            if "PhoneNumberDialed" not in headers:
                greeting.config(text="File Does Not Fit Specifications\n")
            # Otherwise run the code
            else:
                # grab the index where the phone number is
                phone_index = 0
                target_index = 0
                for header in headers:
                    if header == "Response":
                        break
                    else:
                        target_index+=1
                greeting.config(text="File Fits Specs, Sorting")
                reader = csv.reader(spread_sheet, delimiter=',')
                for line in reader:
                        if target_word in line[target_index]:
                            rows.append(line)
    except FileNotFoundError:
        greeting.config(text="File Not Found\n"
                             "Please Browse for Proper CSV File")
    if target_word == voicemail:
        try:
            with open(voicemail_file_path, 'a') as output_file:
                writer = csv.writer(output_file)
                writer.writerows(rows)
        except FileExistsError:
            pass
    elif target_word == disconnected:
        try:
            with open(disconnected_file_path, 'a') as output_file:
                writer = csv.writer(output_file)
                writer.writerows(rows)
        except FileExistsError:
            pass
    rows.clear()
"""

"""
# Fill all sheets for simplicity and single function for button use
def fill_all_sheets():
    # Set a value for filename based on the path that is browsed for by the user
    # Setting it up this way allows for a single button press to browse and then sort
    filename = browseFiles()
    # With the function defined we can fill our call sheets
    fill_call_sheet(disconnected, filename)
    fill_call_sheet(voicemail, filename)
"""

# Functionality to browse for files, starting with CSV files as the default type to find, uses built-in
# file explorer on Windows
def browseFiles():
    filename = filedialog.askopenfilename(initialdir = "/",
                                          title = "Select Unsorted Call Logs",
                                          filetypes = (("CSV files",
                                                        "*.csv*"),
                                                       ("Text files",
                                                        "*.txt*"),
                                                       ("all files",
                                                        "*.*")))
    try:
        with open(filename) as spread_sheet:
            # DictReader will help pull headers from spreadsheets
            d_reader = csv.DictReader(spread_sheet)
            # setting the headers to a variable
            headers = d_reader.fieldnames
            # If it's not in the headers tell the user
            if "PhoneNumberDialed" not in headers:
                greeting.config(text="File Does Not Fit Specifications\n")
                label_call_log_found.config(text = "Missing Unsorted\n"
                                                   "Call Logs Sheet",
                                            bg = colors['pink'])
            # Otherwise run the code
            else:
                # grab the index where the phone number is
                phone_index = 0
                target_index = 0
                for header in headers:
                    if header == "Response":
                        break
                    else:
                        target_index+=1

                greeting.config(text="File Fits Specifications")
                label_call_log_found.config(bg=colors['green'],
                                                text="File Found")

    except FileNotFoundError:
        greeting.config(text="File Not Found\n"
                             "Please Browse for Proper CSV File")

    label_file_explorer.configure(text="File Opened: "+filename)
    # make sure we're reaching to the correct scope, we're trying to change outer scope call_logs_found
    global call_logs_found
    call_logs_found = True

    if call_logs_found and household_id_found:
        button_sort.config(state="normal")
    else:
        button_sort.config(state="disabled")
    global call_logs_filename
    call_logs_filename = filename


def browseForHouseholdIDSheet():
    filename = filedialog.askopenfilename(initialdir = "/",
                                          title = "Select File With Household IDs",
                                          filetypes = (("CSV files",
                                                        "*.csv*"),
                                                       ("Text files",
                                                        "*.txt*"),
                                                       ("all files",
                                                        "*.*")))

    try:
        with open(filename) as spread_sheet:
            # DictReader will help pull headers from spreadsheets
            d_reader = csv.DictReader(spread_sheet)
            # setting the headers to a variable
            headers = d_reader.fieldnames
            # If it's not in the headers tell the user
            if "Household ID" not in headers:
                greeting.config(text="File Does Not Fit Specifications\n")
                label_household_id_found.config(text = "Missing Household\n"
                                                       "ID Sheet",
                                                bg = colors['pink'])
            # Otherwise run the code
            else:
                # alert user to success

                label_household_id_found.config(bg=colors['green'],
                                                text="File Found")
                greeting.config(text="File Fits Specifications")

    except FileNotFoundError:
        greeting.config(text="File Not Found\n"
                             "Please Browse for Proper CSV File")

    label_file_explorer.configure(text="File Opened: "+filename)
    global household_id_found
    household_id_found = True

    if call_logs_found and household_id_found:
        button_sort.config(state="normal")
    else:
        button_sort.config(state="disabled")
    global household_id_filename
    household_id_filename = filename


frame_width = 40
frame_height = 4

colors = {
    'salmon': '#FF7F7F',
    'dark grey': '#404040',
    'darker grey': "#878683",
    'dark purple': '#320484',
    'light purple': '#9574f7',
    'blue': '#3253ed',
    'honey': '#b1883c',
    'light blue': '#b7ccdf',
    'green': '#A3CFA7',
    'pink': '#F7DCEC'
}

minute_options = [
    "00",
    "15",
    "30",
    "45"
]

hour_options = [
    '1:',
    '2:',
    '3:',
    '4:',
    '5:',
    '6:',
    '7:',
    '8:',
    '9:',
    '10:',
    '11:',
    '12:'
]

am_pm_option = [
    'AM',
    'PM'
]

padding = 3

window = tix.Tk()
window.geometry("{0}x{1}+0+0".format(
            window.winfo_screenwidth()-padding, window.winfo_screenheight()-padding))
window.title("Call Sheet Sorter")
window.config(background=colors['dark grey'],
              bd=10)

tip = tix.Balloon(window)

company_logo = ImageTk.PhotoImage(Image.open("trgicon.PNG"))

info_frame = tk.Frame(window)
info_frame.config(bg=colors['dark grey'])
info_frame.pack(pady=5,padx=5)

topFrameBorderColor = tk.Frame(window,
                               width=frame_width,
                               height=frame_height
                               )
topFrameBorderColor.config(bg=colors['dark purple'])
topFrameBorderColor.pack(padx=4,pady=4,)

topFrame = tk.Frame(topFrameBorderColor)
topFrame.config(bg=colors['light blue'])
topFrame.pack(padx=5, pady=5)

greeting = tk.Label(
    topFrame,
    text="Hello, "+username+"!",
    font=("Helvetica", 25),
    foreground="white",
    background=colors['blue'],
    width=frame_width,
    height=frame_height
)
greeting.pack(padx=5,pady=5)

canvas = tk.Canvas(info_frame, width=450, height=80, bg=colors['light blue'])
canvas.create_image(15,7, anchor=tk.NW,image=company_logo)
canvas.pack(padx=5, pady=5)

middleFrame = tk.Frame(window,)
middleFrame.pack(padx=10,pady=10)

middleTopFrame = tk.Frame(middleFrame,
                          bg=colors['dark grey'])
middleTopFrame.pack(side=tk.TOP)
voicemail_label = tk.Label(middleTopFrame,
                              text="Enter File Name for Voicemail Calls File",
                            borderwidth=12)
voicemail_label.pack(side=tk.LEFT,padx=5)
voicemail_file_name_entry = tk.Entry(middleTopFrame,
                                   background="white",
                                   bd=4,
                                   fg="black")
voicemail_file_name_entry.insert(0,'VoicemailCallSheet')
voicemail_file_name_entry.pack(padx=5, pady=5, side=tk.LEFT)
middleBottomFrame = tk.Frame(middleFrame,
                             bg=colors['dark grey'])
middleBottomFrame.pack(side=tk.BOTTOM)
disconnected_label = tk.Label(middleBottomFrame,
                              text="Enter File Name for Disconnected Calls File")
disconnected_label.pack(side=tk.LEFT,padx=5,pady=10)
disconnected_file_name_entry = tk.Entry(middleBottomFrame,
                           background="white",
                           bd=4,
                           fg="black")
disconnected_file_name_entry.insert(0,'DisconnectedCallSheet')
disconnected_file_name_entry.pack(padx=5,pady=5,side=tk.LEFT)
tip.bind_widget(disconnected_file_name_entry,balloonmsg="Enter file name for voicemail log\n"
                                           "or leave blank for default name\n"
                                           "(Will be saved as .csv)")

bottomFrame = tk.Frame(window)
#bottomFrame.config(background=colors['blue'])
bottomFrame.config(borderwidth=4)
bottomFrame.config(bg=colors['dark purple'])
# Have to specifically call on tk to get the right frame type
bottomFrame.pack(side=tk.TOP)

bottomLeftFrame= tk.Frame(bottomFrame)
bottomLeftFrame.config(background=colors['blue'])
bottomLeftFrame.config(borderwidth=1)
bottomLeftFrame.pack(side=tk.LEFT)

bottomRightFrame = tk.Frame(bottomFrame)
bottomRightFrame.config(background=colors['blue'])
bottomRightFrame.config(borderwidth=1)
bottomRightFrame.pack(side=tk.RIGHT)


def onDisconnectedSaveEntry():

    user_given_name = disconnected_file_name_entry.get()
    return user_given_name


def onVoicemailSaveEntry():

    user_given_name = voicemail_file_name_entry.get()
    return user_given_name


label_file_explorer = tk.Label(topFrame,
                            text = "Browse For Household ID Sheet And\n"
                                   "Browse for Unsorted Call Logs\n"
                                    "If you do not enter file names"
                                    " default values will be used",
                            width = 70, height = 4,
                            font=("Helvetica", 15),
                            fg = "blue")
label_file_explorer.pack(padx=5, pady=5, side=tk.BOTTOM)

label_household_id_found = tk.Label(bottomLeftFrame,
                                    text = "Missing Household\n"
                                           "ID Sheet",
                                    width = 20,
                                    height = 10,
                                    bg = colors['pink'],
                                    )
label_household_id_found.pack(side=tk.BOTTOM, padx=30, pady=5)

button_browse_for_ids = tk.Button(bottomLeftFrame,
                        bg=colors['darker grey'],
                        borderwidth=10,
                        font=("Helvetica", 25),
                        text = "Browse For IDs",

                        command = browseForHouseholdIDSheet,
                        )
button_browse_for_ids.pack(side=tk.TOP,
                            padx=30,
                            pady=5)


button_browse_for_call_logs = tk.Button(bottomRightFrame,
                        bg=colors['darker grey'],
                        borderwidth=10,
                        text = "Browse For Call Logs",
                        font=("Helvetica", 25),
                        command = browseFiles)
button_browse_for_call_logs.pack(side=tk.TOP,
                            padx=30,
                            pady=5)

label_call_log_found = tk.Label(bottomRightFrame,
                                    text = "Missing Unsorted\n"
                                           "Call Logs Sheet",
                                    width = 20,
                                    height = 10,
                                    bg = colors['pink'])
label_call_log_found.pack(side=tk.BOTTOM,padx=30, pady=5)

button_sort = tk.Button(window,
                        text="Automated Sort",
                        borderwidth=10,
                        font=("Helvetica", 25),
                        state="disabled",
                        command = lambda : get_all_household_id())
button_sort.pack(padx=5,pady=3)
tip.bind_widget(button_sort,balloonmsg="Hit this to make 2 CSV files,\n"
                                       "one for calls that went to voicemail,\n"
                                       "and another for disconnected calls")

option_menus_frame = tk.Frame(window)
option_menus_frame.pack()

hour_option_var = tk.StringVar(option_menus_frame)
hour_option_var.set(hour_options[0])
hour_option_menu = tk.OptionMenu(option_menus_frame, hour_option_var, *hour_options)
hour_option_menu.pack(side=tk.LEFT)

test_hour = tk.Button(option_menus_frame,command=get_hour)
test_hour.pack()

minute_option_var = tk.StringVar(option_menus_frame)
minute_option_var.set(minute_options[0])
minute_option_menu = tk.OptionMenu(option_menus_frame, minute_option_var, *minute_options)
minute_option_menu.pack(side=tk.LEFT)

am_pm_option_var = tk.StringVar(option_menus_frame)
am_pm_option_var.set(am_pm_option[0])
am_pm_menu = tk.OptionMenu(option_menus_frame, am_pm_option_var, *am_pm_option)
am_pm_menu.pack(side=tk.LEFT)

window.mainloop()










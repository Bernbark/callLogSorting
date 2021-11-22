from __future__ import print_function
import tkinter.tix as tix
import tkinter as tk
from tkinter import filedialog
import csv
import os
import getpass
from PIL import ImageTk, Image

"""
This script allows a user to browse for two files to compare against, using phone numbers to match IDs to users and sort
them into 2 new files. Both files include a Household ID as well as a timestamp for when they were called.
The first file is for all users where the phone call reached their voicemail.
The second file is for all users where the phone call was disconnected.
"""

disconnected_filename = ''
voicemail_filename = ''
voicemail = 'Answering_Machine'
disconnected = 'Disconnected'
rows = []
household_id_found = False
household_id_filename = ''
call_logs_found = False
call_logs_filename = ''

# Grabbing user name for creation of a Desktop folder
username = getpass.getuser()

# C:\Users\UserName\Desktop\CallReports is the format
path_str = "C:\\Users\\" + username + "\\Desktop\\CallReports"
disconnected_path_str = "C:\\Users\\" + username + "\\Desktop\\CallReports\\DisconnectedCalls"
# If the directory does not exist, we make it using the above path
if not os.path.isdir(path_str):
    os.makedirs(path_str)
if not os.path.isdir(disconnected_path_str):
    os.makedirs(disconnected_path_str)

# Check if CSV file exists and create if not, then add headers
# Headers are CallCompletedTimeStamp,PhoneNumberDialed,Response
# ExactBillingDurationInSeconds,RoundedBillingDurationInMinutes
def create_csv_files():
    """
    This method will create 2 separate CSV files and apply headers to them if they don't exist yet
    :return:
    """

    # Alter the file name to be a default if nothing was chosen
    if disconnected_filename == "":
        disconnected_file_path = disconnected_path_str + "\\DisconnectedCallSheet.csv"
    # Otherwise use the file name provided by the user
    else:
        disconnected_file_path = disconnected_path_str + "\\" + disconnected_filename
        # Catchall in case things go south with the naming of files, force it back to default
        if disconnected_file_path == disconnected_path_str + "\\" + ".csv":
            disconnected_file_path = disconnected_path_str + "\\DisconnectedCallSheet.csv"
    # Create a bool to check if the file path exists
    file_exists = os.path.isfile(disconnected_file_path)
    try:
        # Either create and add headers, or do nothing, thus the 'a' setting
        with open(disconnected_file_path, 'a',newline='',) as output_file:
            writer = csv.writer(output_file)
            # set header if file doesn't exist
            if not file_exists:
                writer.writerow(["Household ID", "Call Timestamp"])
    # If the file exists then do nothing
    except FileExistsError:
        pass

    # This is the same as thing as above except for the voicemail file name
    if voicemail_filename == "":
        voicemail_filepath = path_str + "\\VoicemailCallSheet.csv"
    else:
        voicemail_filepath = path_str + "\\" + voicemail_filename
        if voicemail_filepath == path_str + "\\" + ".csv":
            voicemail_filepath = path_str + "\\VoicemailCallSheet.csv"
    file_exists = os.path.isfile(voicemail_filepath)
    try:
        with open(voicemail_filepath, 'a',newline='',) as output_file:
            writer = csv.writer(output_file)
            # set header if file doesn't exist
            if not file_exists:
                writer.writerow(["Household ID", "Call Timestamp"])
    except FileExistsError:
        pass


def get_household_id(household_ids,call_logs,file_path,target_word):
    """
    This method takes two files and searches and strips them of the appropriate information in order to make a
    comparison and sort based on whether the customer disconnected or a voicemail was left, and then puts that
    information into a csv file based on the given file path.
    :param household_ids: path to file chosen by user to compare household IDs
    :param call_logs: path to file chosen by user that contains unsorted call logs
    :param file_path: either path to disconnected or voicemail file which will have new records appended to it
    :param target_word: target word to be searched and sorted, in this case Disconnected or Answering_Machine
    :return: nothing
    """
    # holds the information from household IDs file
    id_rows = []
    # will be used throughout in order to track which column the id exists in the .csv file
    id_index = 0

    # Open the household IDs file and strip the header
    with open(household_ids, 'r',newline='',) as id_log:
        reader = csv.reader(id_log)
        d_reader = csv.DictReader(id_log)
        # setting the headers to a variable
        headers = d_reader.fieldnames
        # once the header for ID is found, break, otherwise keep increasing the id_index
        # this is useful if we know the headers ahead of time
        for header in headers:
            if "ID" in header:
                break
            else:
                id_index+=1
        # read the entire document, line by line, as entrys into id_rows
        for line in reader:
            id_rows.append(line)

    with open(call_logs, 'r',newline='',) as call_log:
        # DictReader will help pull headers from spreadsheets
        d_reader = csv.DictReader(call_log)
        # setting the headers to a variable
        headers = d_reader.fieldnames
        reader = csv.reader(call_log)
        # these indexes are used to give some fluidity to finding the proper column from which to strip information from
        # .csv files, we can call on id_rows[id_index] knowing that for every row in that index will hold info we want
        target_index = 0
        phone_index = 0
        date_index = 0
        # Messy way of finding headers, can probably be combined into a single loop
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
        # Depending on if we're checking for Disconnected or Answering_Machine target word, add the correct lines to
        # rows list, with only the correct information from each line, as a tuple
        for line in reader:
            if target_word in line[target_index]:
                rows.append((line[phone_index],line[date_index]))
        # Now that rows is formed
        for entry in rows:
            # For every id in id_rows
            for id in id_rows:
                # Check if the current entry matches the id fro the current id_row and form a new entry tuple with the
                # right info
                if entry[0] == id[1]:
                    # Since we made entry ourselves, we know that entry[1] will be the timestamp for that entry
                    newEntry = (id[id_index],entry[1])
                    # Finally write that new entry into the output file
                    with open(file_path, 'a+',newline='',) as output_file:
                        writer = csv.writer(output_file)
                        writer.writerow(entry)
        # Probably not necessary, something left over from testing
        rows.clear()
        # Same as above
        id_rows.clear()

# for use with the button to sort once files are selected
def get_all_household_id():
    """
        When user hits the sort button, this method will grab the file names provided by the user, or provide default
        names, and then create CSV files or append to them if they exist already, while also getting the ID and timestamp
        for both voicemail and disconnected calls
        """
    global disconnected_filename
    disconnected_filename = on_disconnected_save_entry()

    global voicemail_filename
    voicemail_filename = on_voicemail_save_entry()
    if voicemail_filename == "":
        voicemail_filename = "VoicemailCallSheet.csv"
    else:
        voicemail_filename = voicemail_filename + ".csv"
    if disconnected_filename == "":
        disconnected_filename = "DisconnectedCallSheet.csv"
    else:
        disconnected_filename = disconnected_filename + ".csv"
    create_csv_files()
    get_household_id(household_id_filename, call_logs_filename, path_str + "\\"+ voicemail_filename, voicemail)
    get_household_id(household_id_filename, call_logs_filename, disconnected_path_str + "\\"+disconnected_filename, disconnected)
    greeting.config(text="Sorting Complete!")


# Functionality to browse for files, starting with CSV files as the default type to find, uses built-in
# file explorer on Windows
def browse_for_call_logs():
    """
    Method which is used to browse for files, it is hooked up to the Browse for Call Logs button, so this one is used
    to search for and double check the validity of call logs files, uses built-in file explorer for Windows
    :return:
    """
    # Sets up browsing with the ability to search all files, .csv, or .txt files, defaulted to csv
    filename = filedialog.askopenfilename(initialdir = "/",
                                          title = "Select Unsorted Call Logs",
                                          filetypes = (("CSV files",
                                                        "*.csv*"),
                                                       ("Text files",
                                                        "*.txt*"),
                                                       ("all files",
                                                        "*.*")))
    try:
        # Read the browsed for file to check for validity
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
            # Otherwise alert the user that their file probably works
            else:
                greeting.config(text="File Fits Specifications")
                label_call_log_found.config(bg=colors['green'],
                                                text="File Found")

    except (FileNotFoundError, UnicodeDecodeError) as e:
        greeting.config(text="File Not Found\n"
                             "Please Browse for Proper CSV File")
    label_file_explorer.config(font=("Helvetica", 11))
    label_file_explorer.configure(text=filename)
    # make sure we're reaching to the correct scope, we're trying to change outer scope call_logs_found
    global call_logs_found
    call_logs_found = True

    if call_logs_found and household_id_found:
        button_sort.config(state="normal")
    else:
        button_sort.config(state="disabled")
    global call_logs_filename
    call_logs_filename = filename


def browse_for_household_id_sheet():
    """
    Method which is used to browse for files, it is hooked up to the Browse for Household ID button, so this one is used
    to search for and double check the validity of household ID files, uses built-in file explorer for Windows
    :return:
    """
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
    # Have to make sure the user isn't putting strange files in
    except (FileNotFoundError, UnicodeDecodeError) as e:
        greeting.config(text="File Not Found\n"
                             "Please Browse for Proper CSV File")
    # Display file name
    label_file_explorer.config(font=("Helvetica", 11))
    label_file_explorer.configure(text=filename)
    # This needs to be changed globally so that the second browse method is aware of the first's success
    global household_id_found
    household_id_found = True
    # If both browses were successful then we activate the sort button
    if call_logs_found and household_id_found:
        button_sort.config(state="normal")
    # Otherwise disable it
    else:
        button_sort.config(state="disabled")
    # This may not be the right way to do things, but because of how I adapted the UI, this seemed like the easiest way
    # to update the filename globally as opposed to returning a value
    global household_id_filename
    household_id_filename = filename

def on_disconnected_save_entry():

    user_given_name = disconnected_file_name_entry.get()
    return user_given_name


def on_voicemail_save_entry():

    user_given_name = voicemail_file_name_entry.get()
    return user_given_name


frame_width = 40
frame_height = 5
# Make a dictionary to track color keys with their matching hex code values
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
padding = 3
"""
        This section is the GUI, made from tkinter which is built into python, an attempt was made to make the UI
        user-friendly, not harsh on the eyes, and simple but with enough information to make it clear what is happening
        NOTE: the order in which objects are packed can greatly shift the UI in unexpected ways, careful moving things.
        Also, must refer to widgets in tkinter with the prefix tk. in order to use them (ex: tk.Frame)
"""

# Create the root window for the whole GUI
window = tix.Tk()
# Set the window to fit the user's screen
window.geometry("{0}x{1}+0+0".format(
            window.winfo_screenwidth()-padding, window.winfo_screenheight()-padding))
window.title("Call Sheet Sorter")
window.config(background=colors['dark grey'],
              bd=10)
# Create an image
company_logo = ImageTk.PhotoImage(Image.open("trgicon.PNG"))
# Create a frame to hold the company logo and potentially other information
info_frame = tk.Frame(window)
info_frame.config(bg=colors['dark grey'])
info_frame.pack(pady=5,padx=5)
# Practicing using frames to give a border to other frames
topFrameBorderColor = tk.Frame(window,
                               width=frame_width,
                               height=frame_height
                               )
topFrameBorderColor.config(bg=colors['dark purple'])
topFrameBorderColor.pack(padx=4,pady=4,)
# The top frame which holds the greeting and informational label
topFrame = tk.Frame(topFrameBorderColor)
topFrame.config(bg=colors['light blue'])
topFrame.pack(padx=5, pady=5)
# Greeting to the user and other various information such as success or failure of browsing/sort
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
# More information for the user, changes to display file names browsed for
label_file_explorer = tk.Label(topFrame,
                            text = "Browse For Household ID Sheet And\n"
                                   "Browse for Unsorted Call Logs\n"
                                    "If you do not enter file names"
                                    " default values will be used",
                            width = 70, height = 4,
                            font=("Helvetica", 13),
                            fg = "blue")
label_file_explorer.pack(padx=5, pady=5, side=tk.BOTTOM)
# Holds the company logo
canvas = tk.Canvas(info_frame, width=450, height=80, bg=colors['light blue'])
canvas.create_image(15,7, anchor=tk.NW,image=company_logo)
canvas.pack(padx=5, pady=5)
# Frame which holds the entry sections and their labels (where users can change file names)
middleFrame = tk.Frame(window,)
middleFrame.pack(padx=10,pady=10)
# Goes inside of middleFrame for organizing the placement of entry and label
middleTopFrame = tk.Frame(middleFrame,
                          bg=colors['dark grey'])
middleTopFrame.pack(side=tk.TOP)
# Information for user
voicemail_label = tk.Label(middleTopFrame,
                              text="Enter File Name for Voicemail Calls File",
                            borderwidth=12)
voicemail_label.pack(side=tk.LEFT,padx=5)
# Place for user to enter file name
voicemail_file_name_entry = tk.Entry(middleTopFrame,
                                   background="white",
                                   bd=4,
                                   fg="black")
voicemail_file_name_entry.insert(0,'VoicemailCallSheet')
voicemail_file_name_entry.pack(padx=5, pady=5, side=tk.LEFT)
# Goes inside of middleFrame for organizing the placement of entry and label
middleBottomFrame = tk.Frame(middleFrame,
                             bg=colors['dark grey'])
middleBottomFrame.pack(side=tk.BOTTOM)
# Information for user
disconnected_label = tk.Label(middleBottomFrame,
                              text="Enter File Name for Disconnected Calls File")
disconnected_label.pack(side=tk.LEFT,padx=5,pady=10)
# Place for user to enter file name
disconnected_file_name_entry = tk.Entry(middleBottomFrame,
                           background="white",
                           bd=4,
                           fg="black")
disconnected_file_name_entry.insert(0,'DisconnectedCallSheet')
disconnected_file_name_entry.pack(padx=5,pady=5,side=tk.LEFT)
# Holds Browse Buttons and Indicators
bottomFrame = tk.Frame(window)
bottomFrame.config(borderwidth=4)
bottomFrame.config(bg=colors['dark purple'])
bottomFrame.pack(side=tk.TOP)
# Organizes browse buttons
bottomLeftFrame= tk.Frame(bottomFrame)
bottomLeftFrame.config(background=colors['blue'])
bottomLeftFrame.config(borderwidth=1)
bottomLeftFrame.pack(side=tk.LEFT)
# Organizes browse buttons
bottomRightFrame = tk.Frame(bottomFrame)
bottomRightFrame.config(background=colors['blue'])
bottomRightFrame.config(borderwidth=1)
bottomRightFrame.pack(side=tk.RIGHT)
# Label indicating success or failure in browsing
label_household_id_found = tk.Label(bottomLeftFrame,
                                    text = "Missing Household\n"
                                           "ID Sheet",
                                    width = 20,
                                    height = 10,
                                    bg = colors['pink'],
                                    )
label_household_id_found.pack(side=tk.BOTTOM, padx=30, pady=5)
# Button to initiate file explorer and browse for files
button_browse_for_ids = tk.Button(bottomLeftFrame,
                        bg=colors['darker grey'],
                        borderwidth=10,
                        font=("Helvetica", 25),
                        text = "Browse For IDs",

                        command = browse_for_household_id_sheet,
                        )
button_browse_for_ids.pack(side=tk.TOP,
                            padx=30,
                            pady=5)

# Button to initiate file explorer and browse for files
button_browse_for_call_logs = tk.Button(bottomRightFrame,
                        bg=colors['darker grey'],
                        borderwidth=10,
                        text = "Browse For Call Logs",
                        font=("Helvetica", 25),
                        command = browse_for_call_logs)
button_browse_for_call_logs.pack(side=tk.TOP,
                            padx=30,
                            pady=5)
# Label indicating success or failure in browsing
label_call_log_found = tk.Label(bottomRightFrame,
                                    text = "Missing Unsorted\n"
                                           "Call Logs Sheet",
                                    width = 20,
                                    height = 10,
                                    bg = colors['pink'])
label_call_log_found.pack(side=tk.BOTTOM,padx=30, pady=5)
# Button used to sort manually
button_sort = tk.Button(bottomFrame,
                        text="Sort",
                        borderwidth=10,
                        font=("Helvetica", 25),
                        state="disabled",
                        command = lambda : get_all_household_id())
button_sort.pack(padx=5,pady=3,side=tk.TOP)


# Start the GUI
window.mainloop()










import tkinter
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl

def enter_data():
    accepted = accept_var.get()

    if accepted=="Accepted":
        # User Info
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()
        if firstname and lastname:

            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()

            # Course Info
            registration_status = reg_status_var.get()
            numcourses = numcourses_spinbox.get()
            numsemesters = numsemesters_spinbox.get()

            print("First name: ", firstname, "Last name: ", lastname)
            print("Title: ", title, "Age: ", age, "Nationality: ", nationality)
            print("Registration Status: ", registration_status)
            print("------------------------------------------")

            filepath = "C:/Documents/python/PythonHala/DataEntry_tkinter/dataEntry.xlsx"

            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["First Name", "Last Name", "Title", "Age", "Nationality",
                           "# Courses", "# Semesters", "Registration Status"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstname, lastname, title, age, nationality, numcourses,
                          numsemesters, registration_status])
            workbook.save(filepath)
        else:
            tk.messagebox.showwarning(title="Error", message="First name and last name are required")
    else:
        tk.messagebox.showwarning(title="Error", message="You have not accepted the terms")
# Create a root window i.e. the parent window for everything else
# window is the widget that will contain all of the other widgets
# think of it like the largest container
window = tk.Tk()
# for a title
window.title("Data Entry Form")


# layout mgr - to position widgets
# frame
frame = tk.Frame(window)
# frame = tk.Frame(bg='#89f0df')
frame.pack()

# Saving User Info
user_info_frame = tk.LabelFrame(frame, text="User Information")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)
# Label
first_name_label = tk.Label(user_info_frame,text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = tk.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0, column=1)
# Input
first_name_entry = tk.Entry(user_info_frame)
last_name_entry = tk.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

title_label = tk.Label(user_info_frame, text="Title" )
title_combobox = ttk.Combobox(user_info_frame, values=['', 'Mr.', 'Ms.', 'Dr.'])
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)

#spin box
age_label = tk.Label(user_info_frame, text="Age")
age_spinbox = tk.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

nationality_label = tk.Label(user_info_frame, text="Nationality")
nationality_combobox = ttk.Combobox(user_info_frame, values=["Africa", "Antartica", "Asia", "Europe", "North America", "South America"])
nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

#Saving Course Info
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

registered_label = tk.Label(courses_frame, text="Registration Status")
reg_status_var = tk.StringVar(value="Not Registered") # store value
registered_check = tk.Checkbutton(courses_frame, text="Currently Registered",
                                  variable=reg_status_var,
                                  onvalue="Registered", offvalue="Not Registered")
registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

numcourses_label = tk.Label(courses_frame, text="# Completed Courses")
numcourses_spinbox = tk.Spinbox(courses_frame, from_=0, to='infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)

numsemesters_label = tk.Label(courses_frame, text="# Semesters")
numsemesters_spinbox = tk.Spinbox(courses_frame, from_=0, to="infinity")
numsemesters_label.grid(row=0, column=2)
numsemesters_spinbox.grid(row=1, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Accept Terms
terms_frame = tk.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tk.StringVar(value="Not Accepted")
terms_check = tk.Checkbutton(terms_frame, text = "I accept the terms and conditions.",
                             variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

# Button
button = tk.Button(frame, text = "Enter data", command=enter_data)
button.grid(row=3, column=-0, sticky="news", padx=20, pady=10)


# To fun tkinter app, need to use window.mainloop
# will run an infinite loop as long the app is being executed
# will continue running until x out.
window.mainloop()



import tkinter
from tkinter import ttk
from tkinter import messagebox
import tkinter.messagebox
import os
import openpyxl

def enter_data():
    accepted = accept_var.get()

    if accepted== "Accepted":
        #information gathering
        FirstName = first_name_entry.get()
        LastName = last_name_entry.get()
        if FirstName and LastName:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()
            #course information function
            registration_status = reg_status_var.get()
            numcourses = numcourses_spinbox.get()
            numsemesters = numsemesters_spinbox.get() 

            print("FirstName:", FirstName, "Last Name:",LastName)
            print("title:",title,"age", age, "nationality", nationality)
            print("#courses", numcourses, "Semester:",numsemesters)
            print("Registration status", registration_status)
            print("-------------------------------------------")

            #creating the excel file with python.
            #filepath = "C:\Users\tmuze\AlxProject\data.xlsx"
            #os will help check availability of the excel file

            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["FirstName","LastName","Title","Age","Nationality","#Courses","#Semesters","Registration status"]
                #to appent the heading data to the table 
                sheet.append(heading)

            #data validation on input 
        else:
            tkinter.messagebox.showwarning(title="Error", message="First name and last name required")
    else:
        tkinter.messagebox.showwarning(title= "Error", message="You have not accepted the terms")

window = tkinter.Tk()
window.title("Data entry Form")

frame = tkinter.Frame(window)
frame.pack()
#write form frame user information
user_info_frame = tkinter.LabelFrame(frame, text="User Information")
user_info_frame.grid(row=0,column=0, padx=20, pady=10)
#form lables 
first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0, column=1)
#form entry
first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)
#combo box will import using ttk module 
title_label = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=["","Mr.", "Ms.", "Dr."])
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)

age_lable = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)
age_lable.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

nationality_label = tkinter.Label(user_info_frame, text="Nationality")
nationality_combobox = ttk.Combobox(user_info_frame, values=["Africa","Antarctica","Asia"])
nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)
#changing the widget size 
for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)
#saving course infor
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)
#registration infor
registered_label = tkinter.Label(courses_frame, text="Registration Status")
#function for the checker
reg_status_var = tkinter.StringVar(value="Not Registered")
registered_check = tkinter.Checkbutton(courses_frame, text="Current Registration",
                                       variable=reg_status_var, onvalue="Registered", offvalue="Not Registered")
registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)
#infor on courses available
numcourses_label = tkinter.Label(courses_frame, text="#Completed Courses")
numcourses_spinbox = tkinter.Spinbox(courses_frame,from_=0, to='Infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)
#infor on semester
numsemesters_label = tkinter.Label(courses_frame, text="# Semesters")
numsemesters_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='Infinity')
numsemesters_label.grid(row=0, column=2)
numsemesters_spinbox.grid(row=1, column=2)
# spacing for the courses widget
for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

#accepting terms and conditions 
terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)
accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text="Iaccept the terms and conditions",
                                  variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)
#inserting a button 
button= tkinter.Button(frame, text="Enter data", command= enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)





window.mainloop()
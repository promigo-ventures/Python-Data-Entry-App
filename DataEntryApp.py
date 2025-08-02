import tkinter.messagebox
from tkinter import *
from tkinter.ttk import Combobox
# import openpyxl
import os
import sqlite3
from tkinter.filedialog import asksaveasfilename
# from openpyxl.reader.excel import load_workbook
# from openpyxl.styles.builtins import title


def enter_data():
    terms = terms_status.get()
    if terms == "Accepted":
        firstname = firstname_entry.get()
        lastname = lastname_entry.get()
        if firstname and lastname:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()
            reg_student = reg_students.get()
            courses = courses_spinbox.get()
            semesters = semesters_spinbox.get()
            print(f"""
            FirstName: {firstname},
            LastName: {lastname},
            Title: {title},
            Age: {age},
            Nationality: {nationality},
            Registration Status: {reg_student} 
            Courses: {courses},
            Semesters: {semesters}
            """)

            # filepath = "/Users/apple/Documents/PycharmProjects/DataEntryApp/data.xlsx"

            # if not os.path.exists(filepath):
            #     workbook = openpyxl.Workbook()
            #     sheet = workbook.active
            #     heading = ["First Name", "Last Name", "Title", "Age", "Nationality", "Courses", "Semesters", "Registration Status"]
            #     sheet.append(heading)
            #     workbook.save(filepath)
            # workbook = load_workbook(filepath)
            # sheet = workbook.active
            # sheet.append([firstname, lastname, title, age, nationality, courses, semesters, reg_student])
            # workbook.save(filepath)

            # home_dir = os.path.expanduser("~")
            # db_path = os.path.join(home_dir, "Documents", "data.db")  # ✅ fixed path

            db_path = asksaveasfilename(
                defaultextension=".db",
                filetypes=[("SQLite Database", "*.db")],
                title = "Save database as..."
            )
            if not db_path:
                tkinter.messagebox.showwarning(title="Cancelled", message="No file selected. Data not saved.")
                return
            os.makedirs(os.path.dirname(db_path), exist_ok=True)
            conn = sqlite3.connect(db_path)
            table_create_query = '''
            CREATE TABLE IF NOT EXISTS Student_Data(firstname TEXT, lastname TEXT, title 
            TEXT, age INT, nationality TEXT, num_courses INT, num_semesters  INT, registration_status 
            TEXT)
            '''
            conn.execute(table_create_query)
            data_insert_query = '''INSERT INTO Student_Data(firstname, lastname, title, age, nationality, num_courses, num_semesters, registration_status) VALUES(?, ?, ?, ?, 
            ?, ?, ?, ?)
            '''
            data_insert_tuple = (firstname, lastname, title, age, nationality, courses, semesters, reg_student)
            cursor = conn.cursor()
            cursor.execute(data_insert_query, data_insert_tuple)
            conn.commit()
            print("Data inserted successfully.")
            conn.close()
            tkinter.messagebox.showinfo(title="Success", message=f"Data saved to: {db_path}")

        else:
            tkinter.messagebox.showwarning(title="Error", message="Firstname and lastname are required")
    else:
        tkinter.messagebox.showwarning(title="Error", message="You have not accepted the terms")

window = Tk()
window.title("Data Entry Form")
frame = Frame(window)
frame.pack()

user_info_frame = LabelFrame(frame, text="User Information", width="12", height="4")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

firstname_label = Label(user_info_frame, text="First Name", width="12", height="4")
firstname_label.grid(row=0, column=0)
lastname_label = Label(user_info_frame, text="Last Name", width="12", height="4")
lastname_label.grid(row=0, column=1)

firstname_entry = Entry(user_info_frame)
firstname_entry.grid(row=1, column=0)
lastname_entry = Entry(user_info_frame)
lastname_entry.grid(row=1, column=1)

title_label = Label(user_info_frame, text="Title")
title_label.grid(row=0, column=2)
title_combobox = Combobox(user_info_frame, values=["Mr", "Ms", "Mrs"])
title_combobox.grid(row=1, column=2)

age_label = Label(user_info_frame, text="Age")
age_label.grid(row=2, column=0)
age_spinbox = Spinbox(user_info_frame, from_=18, to=30)
age_spinbox.grid(row=3, column=0)

nationality_label = Label(user_info_frame, text="Nationality")
nationality_label.grid(row=2, column=1)
nationality_combobox = Combobox(user_info_frame, values=["Africa", "Asia", "Europe", "North America"])
nationality_combobox.grid(row=3, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

registration_label = LabelFrame(frame, text="Registration")
registration_label.grid(row=1, column=0, padx=20, pady=10, sticky="news")

reg_students =StringVar(value="Not Registered")
registration_status_label = Label(registration_label, text="Registration Status")
registration_status_label.grid(row=0, column=0)
registration_status_check = Checkbutton(registration_label, text="Currently Registered", variable=reg_students, onvalue="Registered", offvalue="Not Registered")
registration_status_check.grid(row=1, column=0)

courses_label = Label(registration_label, text="Completed Courses")
courses_label.grid(row=0, column=1)
courses_spinbox = Spinbox(registration_label, from_=0, to=20)
courses_spinbox.grid(row=1, column=1)

semesters_label = Label(registration_label, text="Semesters")
semesters_label.grid(row=0, column=2)
semesters_spinbox = Spinbox(registration_label, from_=1, to=5)
semesters_spinbox.grid(row=1, column=2)

for widget in registration_label.winfo_children():
    widget.grid_configure(padx=10, pady=5)

terms_status = StringVar(value="Not Accepted")
terms_label = LabelFrame(frame, text="Terms & Conditions")
terms_label.grid(row=3, column=0, sticky="news", padx=20, pady=10)

terms_check = Checkbutton(terms_label, text="I accept the terms and conditions.", variable=terms_status, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

for widget in terms_label.winfo_children():
    widget.grid_configure(padx=10, pady=5)

submit_button = Button(frame, text="Enter data", bg="#4CAF50", command=enter_data)
submit_button.grid(row=4, column=0, sticky="news", padx=20, pady=10)
window.mainloop()
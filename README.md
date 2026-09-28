# STUDENT MANAGEMENT SYSTEM Uing MySQL + Python + Tkinter

import tkinter as tk
from tkinter import messagebox, ttk
import mysql.connector
from mysql.connector import Error
from openpyxl import Workbook, load_workbook
import os


# =====================================================
# MYSQL PASSWORD
# =====================================================

MYSQL_PASSWORD = "YOUR_MYSQL_PASSWORD"


# =====================================================
# EXCEL FILE
# =====================================================

EXCEL_FILE = "students.xlsx"


# =====================================================
# DATABASE CONNECTION
# =====================================================

def connect_database():

    try:

        # Connect to MySQL server
        connection = mysql.connector.connect(
            host="localhost",
            user="root",
            password=MYSQL_PASSWORD
        )

        cursor = connection.cursor()

        # Create database
        cursor.execute(
            "CREATE DATABASE IF NOT EXISTS student_db"
        )

        cursor.close()
        connection.close()

        # Connect to student database
        connection = mysql.connector.connect(
            host="localhost",
            user="root",
            password=MYSQL_PASSWORD,
            database="student_db"
        )

        cursor = connection.cursor()

        # Create students table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS students (
                id INT PRIMARY KEY,
                name VARCHAR(100) NOT NULL,
                class_name VARCHAR(50) NOT NULL,
                email VARCHAR(100),
                phone VARCHAR(15)
            )
        """)

        connection.commit()

        cursor.close()
        connection.close()

        return True

    except Error as e:

        messagebox.showerror(
            "Database Error",
            str(e)
        )

        return False


# =====================================================
# CREATE EXCEL FILE
# =====================================================

def create_excel_file():

    if not os.path.exists(EXCEL_FILE):

        workbook = Workbook()

        sheet = workbook.active
        sheet.title = "Students"

        sheet.append([
            "Student ID",
            "Student Name",
            "Class",
            "Email",
            "Phone"
        ])

        workbook.save(EXCEL_FILE)


# =====================================================
# UPDATE EXCEL FILE FROM MYSQL
# =====================================================

def update_excel():

    try:

        connection = mysql.connector.connect(
            host="localhost",
            user="root",
            password=MYSQL_PASSWORD,
            database="student_db"
        )

        cursor = connection.cursor()

        cursor.execute(
            "SELECT * FROM students ORDER BY id"
        )

        students = cursor.fetchall()

        cursor.close()
        connection.close()

        # Create new Excel workbook
        workbook = Workbook()

        sheet = workbook.active
        sheet.title = "Students"

        # Header
        sheet.append([
            "Student ID",
            "Student Name",
            "Class",
            "Email",
            "Phone"
        ])

        # Student data
        for student in students:

            sheet.append([
                student[0],
                student[1],
                student[2],
                student[3],
                student[4]
            ])

        workbook.save(EXCEL_FILE)

    except Error as e:

        messagebox.showerror(
            "Excel Error",
            str(e)
        )


# =====================================================
# LOGIN WINDOW
# =====================================================

def create_login():

    login = tk.Tk()

    login.title(
        "Student Management System - Login"
    )

    login.geometry("400x300")

    login.resizable(False, False)

    # Title
    tk.Label(
        login,
        text="Student Management System",
        font=("Arial", 18, "bold")
    ).pack(pady=30)

    # Username
    tk.Label(
        login,
        text="Username"
    ).pack()

    username_entry = tk.Entry(
        login,
        width=30
    )

    username_entry.pack(pady=5)

    # Password
    tk.Label(
        login,
        text="Password"
    ).pack()

    password_entry = tk.Entry(
        login,
        width=30,
        show="*"
    )

    password_entry.pack(pady=5)

    # Login function
    def login_check():

        username = username_entry.get()
        password = password_entry.get()

        if username == "admin" and password == "1234":

            login.destroy()

            create_dashboard()

        else:

            messagebox.showerror(
                "Login Failed",
                "Invalid username or password"
            )

    # Login button
    tk.Button(
        login,
        text="Login",
        width=15,
        command=login_check
    ).pack(pady=20)

    login.mainloop()


# =====================================================
# DASHBOARD
# =====================================================

def create_dashboard():

    root = tk.Tk()

    root.title(
        "Student Management System"
    )

    root.geometry("1000x650")


    # =================================================
    # TITLE
    # =================================================

    tk.Label(
        root,
        text="Student Management System",
        font=("Arial", 22, "bold")
    ).pack(pady=15)


    # =================================================
    # INPUT FRAME
    # =================================================

    input_frame = tk.Frame(root)

    input_frame.pack(pady=10)


    # Student ID
    tk.Label(
        input_frame,
        text="Student ID"
    ).grid(
        row=0,
        column=0,
        padx=10,
        pady=5
    )

    id_entry = tk.Entry(
        input_frame,
        width=30
    )

    id_entry.grid(
        row=0,
        column=1,
        padx=10,
        pady=5
    )


    # Student Name
    tk.Label(
        input_frame,
        text="Student Name"
    ).grid(
        row=1,
        column=0,
        padx=10,
        pady=5
    )

    name_entry = tk.Entry(
        input_frame,
        width=30
    )

    name_entry.grid(
        row=1,
        column=1,
        padx=10,
        pady=5
    )


    # Class
    tk.Label(
        input_frame,
        text="Class"
    ).grid(
        row=2,
        column=0,
        padx=10,
        pady=5
    )

    class_entry = tk.Entry(
        input_frame,
        width=30
    )

    class_entry.grid(
        row=2,
        column=1,
        padx=10,
        pady=5
    )


    # Email
    tk.Label(
        input_frame,
        text="Email"
    ).grid(
        row=3,
        column=0,
        padx=10,
        pady=5
    )

    email_entry = tk.Entry(
        input_frame,
        width=30
    )

    email_entry.grid(
        row=3,
        column=1,
        padx=10,
        pady=5
    )


    # Phone
    tk.Label(
        input_frame,
        text="Phone"
    ).grid(
        row=4,
        column=0,
        padx=10,
        pady=5
    )

    phone_entry = tk.Entry(
        input_frame,
        width=30
    )

    phone_entry.grid(
        row=4,
        column=1,
        padx=10,
        pady=5
    )


    # =================================================
    # CLEAR FUNCTION
    # =================================================

    def clear_fields():

        id_entry.delete(0, tk.END)
        name_entry.delete(0, tk.END)
        class_entry.delete(0, tk.END)
        email_entry.delete(0, tk.END)
        phone_entry.delete(0, tk.END)


    # =================================================
    # ADD STUDENT
    # =================================================

    def add_student():

        if (
            id_entry.get() == ""
            or name_entry.get() == ""
            or class_entry.get() == ""
        ):

            messagebox.showwarning(
                "Warning",
                "Student ID, Name and Class are required"
            )

            return


        try:

            connection = mysql.connector.connect(
                host="localhost",
                user="root",
                password=MYSQL_PASSWORD,
                database="student_db"
            )

            cursor = connection.cursor()


            query = """
                INSERT INTO students
                (id, name, class_name, email, phone)
                VALUES (%s, %s, %s, %s, %s)
            """


            values = (
                int(id_entry.get()),
                name_entry.get(),
                class_entry.get(),
                email_entry.get(),
                phone_entry.get()
            )


            cursor.execute(
                query,
                values
            )

            connection.commit()


            cursor.close()
            connection.close()


            # Update Excel
            update_excel()


            messagebox.showinfo(
                "Success",
                "Student added successfully!"
            )


            clear_fields()

            view_students()


        except ValueError:

            messagebox.showerror(
                "Error",
                "Student ID must be a number"
            )


        except Error as e:

            messagebox.showerror(
                "Database Error",
                str(e)
            )


    # =================================================
    # VIEW STUDENTS
    # =================================================

    def view_students():

        for item in tree.get_children():

            tree.delete(item)


        try:

            connection = mysql.connector.connect(
                host="localhost",
                user="root",
                password=MYSQL_PASSWORD,
                database="student_db"
            )


            cursor = connection.cursor()


            cursor.execute(
                "SELECT * FROM students ORDER BY id"
            )


            rows = cursor.fetchall()


            for row in rows:

                tree.insert(
                    "",
                    tk.END,
                    values=row
                )


            cursor.close()
            connection.close()


        except Error as e:

            messagebox.showerror(
                "Database Error",
                str(e)
            )


    # =================================================
    # SEARCH STUDENT
    # =================================================

    def search_student():

        student_id = id_entry.get()


        if student_id == "":

            messagebox.showwarning(
                "Warning",
                "Enter Student ID"
            )

            return


        try:

            connection = mysql.connector.connect(
                host="localhost",
                user="root",
                password=MYSQL_PASSWORD,
                database="student_db"
            )


            cursor = connection.cursor()


            cursor.execute(
                "SELECT * FROM students WHERE id=%s",
                (int(student_id),)
            )


            row = cursor.fetchone()


            cursor.close()
            connection.close()


            if row:

                clear_fields()


                id_entry.insert(
                    0,
                    row[0]
                )

                name_entry.insert(
                    0,
                    row[1]
                )

                class_entry.insert(
                    0,
                    row[2]
                )

                email_entry.insert(
                    0,
                    row[3]
                )

                phone_entry.insert(
                    0,
                    row[4]
                )


            else:

                messagebox.showinfo(
                    "Search",
                    "Student not found"
                )


        except ValueError:

            messagebox.showerror(
                "Error",
                "Student ID must be a number"
            )


        except Error as e:

            messagebox.showerror(
                "Database Error",
                str(e)
            )


    # =================================================
    # UPDATE STUDENT
    # =================================================

    def update_student():

        if id_entry.get() == "":

            messagebox.showwarning(
                "Warning",
                "Enter Student ID"
            )

            return


        try:

            connection = mysql.connector.connect(
                host="localhost",
                user="root",
                password=MYSQL_PASSWORD,
                database="student_db"
            )


            cursor = connection.cursor()


            query = """
                UPDATE students
                SET
                    name=%s,
                    class_name=%s,
                    email=%s,
                    phone=%s
                WHERE id=%s
            """


            values = (
                name_entry.get(),
                class_entry.get(),
                email_entry.get(),
                phone_entry.get(),
                int(id_entry.get())
            )


            cursor.execute(
                query,
                values
            )


            connection.commit()


            if cursor.rowcount > 0:

                messagebox.showinfo(
                    "Success",
                    "Student updated successfully!"
                )

                update_excel()

            else:

                messagebox.showinfo(
                    "Update",
                    "Student not found"
                )


            cursor.close()
            connection.close()


            clear_fields()

            view_students()


        except ValueError:

            messagebox.showerror(
                "Error",
                "Student ID must be a number"
            )


        except Error as e:

            messagebox.showerror(
                "Database Error",
                str(e)
            )


    # =================================================
    # DELETE STUDENT
    # =================================================

    def delete_student():

        student_id = id_entry.get()


        if student_id == "":

            messagebox.showwarning(
                "Warning",
                "Enter Student ID"
            )

            return


        try:

            connection = mysql.connector.connect(
                host="localhost",
                user="root",
                password=MYSQL_PASSWORD,
                database="student_db"
            )


            cursor = connection.cursor()


            cursor.execute(
                "DELETE FROM students WHERE id=%s",
                (int(student_id),)
            )


            connection.commit()


            if cursor.rowcount > 0:

                messagebox.showinfo(
                    "Success",
                    "Student deleted successfully!"
                )

                update_excel()

            else:

                messagebox.showinfo(
                    "Delete",
                    "Student not found"
                )


            cursor.close()
            connection.close()


            clear_fields()

            view_students()


        except ValueError:

            messagebox.showerror(
                "Error",
                "Student ID must be a number"
            )


        except Error as e:

            messagebox.showerror(
                "Database Error",
                str(e)
            )


    # =================================================
    # BUTTON FRAME
    # =================================================

    button_frame = tk.Frame(root)

    button_frame.pack(pady=10)


    tk.Button(
        button_frame,
        text="Add Student",
        width=15,
        command=add_student
    ).grid(
        row=0,
        column=0,
        padx=5
    )


    tk.Button(
        button_frame,
        text="View Students",
        width=15,
        command=view_students
    ).grid(
        row=0,
        column=1,
        padx=5
    )


    tk.Button(
        button_frame,
        text="Search",
        width=15,
        command=search_student
    ).grid(
        row=0,
        column=2,
        padx=5
    )


    tk.Button(
        button_frame,
        text="Update",
        width=15,
        command=update_student
    ).grid(
        row=0,
        column=3,
        padx=5
    )


    tk.Button(
        button_frame,
        text="Delete",
        width=15,
        command=delete_student
    ).grid(
        row=0,
        column=4,
        padx=5
    )


    tk.Button(
        button_frame,
        text="Clear",
        width=15,
        command=clear_fields
    ).grid(
        row=0,
        column=5,
        padx=5
    )


    # =================================================
    # TABLE
    # =================================================

    table_frame = tk.Frame(root)

    table_frame.pack(
        fill=tk.BOTH,
        expand=True,
        padx=20,
        pady=10
    )


    columns = (
        "ID",
        "Name",
        "Class",
        "Email",
        "Phone"
    )


    tree = ttk.Treeview(
        table_frame,
        columns=columns,
        show="headings"
    )


    for column in columns:

        tree.heading(
            column,
            text=column
        )

        tree.column(
            column,
            width=170
        )


    tree.pack(
        fill=tk.BOTH,
        expand=True
    )


    # Load existing students
    view_students()


    root.mainloop()


# =====================================================
# PROGRAM START
# =====================================================

if __name__ == "__main__":

    create_excel_file()

    if connect_database():

        create_login()

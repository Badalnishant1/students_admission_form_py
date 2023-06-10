import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import admission_form
import openpyxl
import os

def login():
    username = username_entry.get()
    password = password_entry.get()

    if username == "User1" and password == "123":
        login_window.destroy()
        main_menu()
    else:
        messagebox.showerror("Error", "Invalid username or password")

def new_admission():
    admission_form.open_admission_form()

def load_details():
    try:
        workbook = openpyxl.load_workbook("admission_data.xlsx")
        workbook_path = os.path.abspath("admission_data.xlsx")

        # Open the Excel file using the default application
        os.startfile(workbook_path)

        workbook.close()
    except FileNotFoundError:
        messagebox.showerror("Error", "Admission data file not found.")

def main_menu():
    menu_window = tk.Tk()
    menu_window.title("Admission Form")
    menu_window.geometry("780x500")

    # Load and set the background image
    background_image = Image.open("GOVT NANAKPUR.jpg")
    background_photo = ImageTk.PhotoImage(background_image)
    background_label = tk.Label(menu_window, image=background_photo)
    background_label.place(x=0, y=0, relwidth=1, relheight=1)

    new_admission_button = tk.Button(menu_window, text="New Admission Form", command=new_admission)
    new_admission_button.pack(pady=10)

    load_details_button = tk.Button(menu_window, text="Load Details of Already Admitted Students", command=load_details)
    load_details_button.pack(pady=10)

    menu_window.mainloop()

login_window = tk.Tk()
login_window.title("Login")
login_window.geometry("300x300")

username_label = tk.Label(login_window, text="Username:")
username_label.pack()

username_entry = tk.Entry(login_window)
username_entry.pack(pady=5)

password_label = tk.Label(login_window, text="Password:")
password_label.pack()

password_entry = tk.Entry(login_window, show="*")
password_entry.pack(pady=5)

login_button = tk.Button(login_window, text="Login", command=login)
login_button.pack(pady=10)

login_window.mainloop()

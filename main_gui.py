import tkinter as tk
from tkinter import messagebox
import openpyxl
from tkinter import ttk
import re

def validate_input():
    age = age_entry.get()
    mobile_number = mobile_entry.get()
    
    if not re.match(r"^\d{2}$", age):
        messagebox.showerror("Invalid Input", "Please enter a 2-digit age.")
        return False
    
    if not re.match(r"^\d{10}$", mobile_number):
        messagebox.showerror("Invalid Input", "Please enter a 10-digit mobile number.")
        return False
    
    return True

def submit_form():
    if not validate_input():
        return
    
    name = name_entry.get()
    age = age_entry.get()
    father_name = father_entry.get()
    mother_name = mother_entry.get()
    admission_class = admission_combobox.get()
    address = address_entry.get()
    mobile_number = mobile_entry.get()
    
    # Create a new row with the form data
    new_row = [name, age, father_name, mother_name, admission_class, address, mobile_number]
    sheet.append(new_row)
    workbook.save("admitted_students.xlsx")
    
    messagebox.showinfo("Success", "Form submitted successfully!")
    clear_form()

def clear_form():
    name_entry.delete(0, tk.END)
    age_entry.delete(0, tk.END)
    father_entry.delete(0, tk.END)
    mother_entry.delete(0, tk.END)
    admission_combobox.set('')
    address_entry.delete(0, tk.END)
    mobile_entry.delete(0, tk.END)

# Create the main application window
root = tk.Tk()
root.title("Admission Form")
# root.geometry("1280x720")

# Create labels and entry fields for each detail
name_label = tk.Label(root, text="Name:")
name_label.grid(row=0, column=0, padx=10, pady=10)
name_entry = tk.Entry(root)
name_entry.grid(row=0, column=1, padx=10, pady=10)

age_label = tk.Label(root, text="Age:")
age_label.grid(row=1, column=0, padx=10, pady=10)
age_entry = tk.Entry(root)
age_entry.grid(row=1, column=1, padx=10, pady=10)

father_label = tk.Label(root, text="Father Name:")
father_label.grid(row=2, column=0, padx=10, pady=10)
father_entry = tk.Entry(root)
father_entry.grid(row=2, column=1, padx=10, pady=10)

mother_label = tk.Label(root, text="Mother Name:")
mother_label.grid(row=3, column=0, padx=10, pady=10)
mother_entry = tk.Entry(root)
mother_entry.grid(row=3, column=1, padx=10, pady=10)

admission_label = tk.Label(root, text="Select Branch:")
admission_label.grid(row=4, column=0, padx=10, pady=10)
admission_combobox = ttk.Combobox(root, values=["Computer Engineering", "Mechanical Engineering", "Electronics Engineering", "Civil Engineering", "Electrical Engineering"])
admission_combobox.grid(row=4, column=1, padx=10, pady=10)

address_label = tk.Label(root, text="Address:")
address_label.grid(row=5, column=0, padx=10, pady=10)
address_entry = tk.Entry(root)
address_entry.grid(row=5, column=1, padx=10, pady=10)

mobile_label = tk.Label(root, text="Mobile Number:")
mobile_label.grid(row=6, column=0, padx=10, pady=10)
mobile_entry = tk.Entry(root)
mobile_entry.grid(row=6, column=1, padx=10, pady=10)

# Create a submit button
submit_button = tk.Button(root, text="Submit", command=submit_form)
submit_button.grid(row=7, column=0, columnspan=2, padx=10, pady=10)

# Create an Excel workbook and sheet
workbook = openpyxl.Workbook()
sheet = workbook.active

# Add headers to the Excel sheet
headers = ["Name", "Age", "Father Name", "Mother Name", "Admission Class", "Address", "Mobile Number"]
sheet.append(headers)

# Start the GUI event loop
root.mainloop()

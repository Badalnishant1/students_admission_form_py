import tkinter as tk
from tkinter import messagebox
import openpyxl
from tkinter import ttk
import subprocess
from PIL import ImageTk, Image
def open_new_admission():
    try:
        subprocess.Popen(['python', 'main_gui.py'])
    except Exception as e:
        messagebox.showerror("Error", str(e))

def open_student_details():
    student_details_window = tk.Toplevel(root)
    student_details_window.title("Admitted Students Details")

    # Create a Treeview widget to display the details
    tree = ttk.Treeview(student_details_window)

    # Define the columns
    tree["columns"] = ("S.No", "Name", "Age", "Father Name", "Mother Name", "Admission Class", "Address", "Mobile Number")

    # Format the columns
    tree.column("#0", width=0, stretch=tk.NO)
    tree.column("S.No", width=50, anchor=tk.CENTER)
    tree.column("Name", width=100, anchor=tk.CENTER)
    tree.column("Age", width=50, anchor=tk.CENTER)
    tree.column("Father Name", width=100, anchor=tk.CENTER)
    tree.column("Mother Name", width=100, anchor=tk.CENTER)
    tree.column("Admission Class", width=80, anchor=tk.CENTER)
    tree.column("Address", width=150, anchor=tk.CENTER)
    tree.column("Mobile Number", width=100, anchor=tk.CENTER)

    # Define the column headings
    tree.heading("#0", text="", anchor=tk.CENTER)
    tree.heading("S.No", text="S.No", anchor=tk.CENTER)
    tree.heading("Name", text="Name", anchor=tk.CENTER)
    tree.heading("Age", text="Age", anchor=tk.CENTER)
    tree.heading("Father Name", text="Father Name", anchor=tk.CENTER)
    tree.heading("Mother Name", text="Mother Name", anchor=tk.CENTER)
    tree.heading("Admission Class", text="Admission Class", anchor=tk.CENTER)
    tree.heading("Address", text="Address", anchor=tk.CENTER)
    tree.heading("Mobile Number", text="Mobile Number", anchor=tk.CENTER)

    # Load and display the student details from the Excel file
    try:
        workbook = openpyxl.load_workbook("admitted_students.xlsx")
        sheet = workbook.active

        # Get the header row and data rows
        headers = [cell.value for cell in sheet[1]]
        data = [[cell.value for cell in row] for row in sheet.iter_rows(min_row=2)]

        # Insert the details into the Treeview widget
        for i, row in enumerate(data):
            tree.insert(parent="", index="end", iid=i, text="", values=[i+1] + row)

        tree.pack(fill=tk.BOTH, expand=1)
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Create the main application window
root = tk.Tk()
root.title("Welcome to Government Polytechnic Nanakpur")
root.geometry("600x400")

# Load the background image
background_image = Image.open("GOVT NANAKPUR.jpg")
background_photo = ImageTk.PhotoImage(background_image)

# Create a label for the background image
background_label = tk.Label(root, image=background_photo)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# Set the background label at the back of all other widgets
background_label.lower()
# Create a label for the school name
school_label = tk.Label(root, text="Govt. Poly. College", font=("Arial", 24))
school_label.pack(pady=20)

# Create buttons for new admission and viewing student details
new_admission_button = tk.Button(root, text="New Admission", command=open_new_admission, width=20)
new_admission_button.pack(pady=10)

view_details_button = tk.Button(root, text="View Details", command=open_student_details, width=20)
view_details_button.pack(pady=10)

# Start the GUI event loop
root.mainloop()
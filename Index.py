import tkinter as tk
from tkinter import messagebox
import openpyxl

# Create a new Excel workbook and select the active worksheet
file_path = "Book1.xlsx"  # Replace with your actual file path
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

def add_student_data(student_id, name, roll_number, branch):
    # Append the data to the sheet
    sheet.append([student_id, name, roll_number, branch])

def submit_form():
    student_id = student_id_entry.get()
    name = name_entry.get()
    roll_number = roll_number_entry.get()
    branch = branch_entry.get()

    if student_id and name and roll_number and branch:
        add_student_data(student_id, name, roll_number, branch)
        messagebox.showinfo("Success", "Student data has been added successfully!")
        student_id_entry.delete(0, tk.END)
        name_entry.delete(0, tk.END)
        roll_number_entry.delete(0, tk.END)
        branch_entry.delete(0, tk.END)
    else:
        messagebox.showerror("Error", "Please fill in all the fields!")

# Create the main window
root = tk.Tk()
root.title("Student Data Form")
root.configure(background="#ADD8E6")  # Blue background color

# Create labels and entry fields
student_id_label = tk.Label(root, text="Student ID:", bg="#ADD8E6")
student_id_label.pack()
student_id_entry = tk.Entry(root)
student_id_entry.pack()

name_label = tk.Label(root, text="Name:", bg="#ADD8E6")
name_label.pack()
name_entry = tk.Entry(root)
name_entry.pack()

roll_number_label = tk.Label(root, text="Roll Number:", bg="#ADD8E6")
roll_number_label.pack()
roll_number_entry = tk.Entry(root)
roll_number_entry.pack()

branch_label = tk.Label(root, text="Branch:", bg="#ADD8E6")
branch_label.pack()
branch_entry = tk.Entry(root)
branch_entry.pack()

# Create a submit button
submit_button = tk.Button(root, text="Submit", command=submit_form)
submit_button.pack()

# Run the application
root.mainloop()

# Save the Excel file
workbook.save(file_path)
print(f"Student data has been saved to '{file_path}'")
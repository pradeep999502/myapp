import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import styles
from reportlab.lib.units import inch
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

def resource_path(relative_path):
    # Works for both .py and .exe
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

students_path = resource_path("students.csv")
marks_path = resource_path("marks.csv")

students_df = pd.read_csv(students_path)
marks_df = pd.read_csv(marks_path)
# Main Window
root = tk.Tk()
root.title("Student Marksheet Generator")
root.geometry("500x600")

# Global variable to store current student data
current_student = None
current_marks = None

def search_student():
    global current_student, current_marks
    
    roll = roll_entry.get()
    
    student = students_df[students_df["roll_no"] == int(roll)]
    marks = marks_df[marks_df["roll_no"] == int(roll)]
    
    if student.empty or marks.empty:
        messagebox.showerror("Error", "Student not found!")
        return
    
    current_student = student.iloc[0]
    current_marks = marks.iloc[0]
    
    display_marksheet()

def display_marksheet():
    result_text.delete("1.0", tk.END)
    
    total = 0
    subjects = ["Maths", "Science", "English", "Computer", "Social"]
    
    result_text.insert(tk.END, "------- STUDENT MARKSHEET -------\n\n")
    result_text.insert(tk.END, f"Roll No: {current_student['roll_no']}\n")
    result_text.insert(tk.END, f"Name: {current_student['name']}\n")
    result_text.insert(tk.END, f"Class: {current_student['class']}\n")
    result_text.insert(tk.END, f"Section: {current_student['section']}\n")
    result_text.insert(tk.END, f"DOB: {current_student['dob']}\n\n")
    
    result_text.insert(tk.END, "Subjects and Marks:\n")
    
    for sub in subjects:
        marks = current_marks[sub]
        total += marks
        result_text.insert(tk.END, f"{sub}: {marks}\n")
    
    percentage = total / len(subjects)
    
    result_text.insert(tk.END, "\n")
    result_text.insert(tk.END, f"Total Marks: {total}\n")
    result_text.insert(tk.END, f"Percentage: {percentage:.2f}%\n")

def download_pdf():
    if current_student is None:
        messagebox.showwarning("Warning", "Search a student first!")
        return
    
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                             filetypes=[("PDF files", "*.pdf")])
    if not file_path:
        return
    
    doc = SimpleDocTemplate(file_path)
    elements = []
    
    style = styles.getSampleStyleSheet()
    elements.append(Paragraph("Student Marksheet", style['Heading1']))
    elements.append(Spacer(1, 0.3 * inch))
    
    data = [
        ["Roll No", current_student["roll_no"]],
        ["Name", current_student["name"]],
        ["Class", current_student["class"]],
        ["Section", current_student["section"]],
        ["DOB", current_student["dob"]]
    ]
    
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('GRID', (0,0), (-1,-1), 1, colors.black)
    ]))
    
    elements.append(table)
    elements.append(Spacer(1, 0.5 * inch))
    
    subjects = ["Maths", "Science", "English", "Computer", "Social"]
    marks_data = [["Subject", "Marks"]]
    total = 0
    
    for sub in subjects:
        mark = current_marks[sub]
        total += mark
        marks_data.append([sub, mark])
    
    percentage = total / len(subjects)
    
    marks_data.append(["Total", total])
    marks_data.append(["Percentage", f"{percentage:.2f}%"])
    
    marks_table = Table(marks_data)
    marks_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightblue),
        ('GRID', (0,0), (-1,-1), 1, colors.black)
    ]))
    
    elements.append(marks_table)
    
    doc.build(elements)
    
    messagebox.showinfo("Success", "Marksheet Downloaded Successfully!")

# GUI Layout
tk.Label(root, text="Enter Roll Number", font=("Arial", 12)).pack(pady=10)

roll_entry = tk.Entry(root, font=("Arial", 12))
roll_entry.pack(pady=5)

tk.Button(root, text="Search", command=search_student).pack(pady=5)

result_text = tk.Text(root, height=20, width=60)
result_text.pack(pady=10)

tk.Button(root, text="Download PDF", command=download_pdf, bg="green", fg="white").pack(pady=10)

root.mainloop()
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import sqlite3


#TODO: check ob excel files das richtige ist und endung 
def read_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            df.to_sql('students', con=conn, if_exists='replace', index=False)
            status_label.config(text='Excel file read successfully')
        except Exception as e:
            status_label.config(text=f'Error: {str(e)}')

#TODO: siehe oben
def read_excel_file_2():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            df.to_sql('events', con=conn, if_exists='replace', index=False)
            status_label.config(text='Excel file read successfully')
        except Exception as e:
            status_label.config(text=f'Error: {str(e)}')

#TODO: siehe oben
def read_excel_file_3():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            df.to_sql('companies', con=conn, if_exists='replace', index=False)
            status_label.config(text='Excel file read successfully')
        except Exception as e:
            status_label.config(text=f'Error: {str(e)}')

def display_data():
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM students")
        students = cursor.fetchall()
        cursor.execute("SELECT * FROM events")
        events = cursor.fetchall()
        cursor.execute("SELECT * FROM companies")
        companies = cursor.fetchall()
        result_label.config(text=f'Students: {students}\nEvents: {events}\nCompanies: {companies}')
    except Exception as e:
        result_label.config(text=f'Error: {str(e)}')

def close_program():
    root.destroy()

root = tk.Tk()
root.geometry('400x400')


#TODO: centre text
status_label = tk.Label(root, text='Die Deutsche Presse-Agentur präsentiert \n den Schedulizer9000 GaLiGrü!', anchor='w')
status_label.pack(fill='x', padx=10, pady=10)

read_excel_button = tk.Button(root, text='Read Excel File (Students)', command=read_excel_file)
read_excel_button.pack(fill='x', padx=10, pady=10)

read_excel_button_2 = tk.Button(root, text='Read Excel File (Events)', command=read_excel_file_2)
read_excel_button_2.pack(fill='x', padx=10, pady=10)

read_excel_button_3 = tk.Button(root, text='Read Excel File (Companies)', command=read_excel_file_3)
read_excel_button_3.pack(fill='x', padx=10, pady=10)

display_data_button = tk.Button(root, text='Display Data', command=display_data)
display_data_button.pack(fill='x', padx=10, pady=10)

close_button = tk.Button(root, text='Close', command=close_program)
close_button.pack(fill='x', padx=10, pady=10)
result_label = tk.Label(root, text='', anchor='w')
result_label.pack(fill='x', padx=10, pady=10)

conn = sqlite3.connect(':memory:')

root.mainloop()
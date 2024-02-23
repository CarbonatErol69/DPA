import tkinter as tk
from tkinter import filedialog
import pandas as pd
import sqlite3

def create_database():
    try:
        conn = sqlite3.connect(f'database_{year_entry.get()}.db')
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS students (id INTEGER PRIMARY KEY, name TEXT, age INTEGER, company TEXT)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS events (id INTEGER PRIMARY KEY, name TEXT, date TEXT, company TEXT)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS companies (id INTEGER PRIMARY KEY, name TEXT, restrictions TEXT)''')
        conn.commit()
        status_label.config(text='Database erfolgreich erstellt')
    except Exception as e:
        status_label.config(text=f'Error: {str(e)}')
    finally:
        conn.close()

def read_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            conn = sqlite3.connect(f'database_{year_entry.get()}.db')
            df.to_sql('students', con=conn, if_exists='append', index=False)
            conn.commit()
            status_label.config(text='Datei erfolgreich geladen')
        except Exception as e:
            status_label.config(text=f'Error: {str(e)}')
        finally:
            conn.close()

def read_excel_file_2():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            conn = sqlite3.connect(f'database_{year_entry.get()}.db')
            df.to_sql('events', con=conn, if_exists='append', index=False)
            conn.commit()
            status_label.config(text='Datei erfolgreich geladen')
        except Exception as e:
            status_label.config(text=f'Error: {str(e)}')
        finally:
            conn.close()

def read_excel_file_3():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            conn = sqlite3.connect(f'database_{year_entry.get()}.db')
            df.to_sql('companies', con=conn, if_exists='append', index=False)
            conn.commit()
            status_label.config(text='Datei erfolgreich geladen')
        except Exception as e:
            status_label.config(text=f'Error: {str(e)}')
        finally:
            conn.close()

def distribute_students():
    try:
        conn = sqlite3.connect(f'database_{year_entry.get()}.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM students")
        students = cursor.fetchall()
        cursor.execute("SELECT * FROM events")
        events = cursor.fetchall()
        cursor.execute("SELECT * FROM companies")
        companies = cursor.fetchall()
        result = distribute_students_over_events(students, events, companies)
        result_label.config(text=result)
    except Exception as e:
        result_label.config(text=f'Error: {str(e)}')
    finally:
        conn.close()

def distribute_students_over_events(students, events, companies):
    result = ''
    distributed_students = []
    for student in students:
        student_company = student[3]
        restrictions = [company[2] for company in companies if company[1] == student_company]
        available_events = [event for event in events if event[3] == student_company and event[1] not in restrictions]
        if available_events:
            event = available_events[0]
            result += f'Student {student[1]} will attend {event[1]} on {event[2]}\n'
            distributed_students.append((student[1], event[1], event[2])) # Collect distributed student info
        else:
            result += f'No available events for student {student[1]}\n'
    create_excel_files(distributed_students) # Create Excel file with distributed students
    return result

#TODO: alle Listen erstellen lassen, nicht nur Anwesenheitsliste
def create_excel_files(distributed_students):
    df = pd.DataFrame(distributed_students, columns=['Schüler Name', 'Veranstaltung', 'Zeitrahmen'])
    df.to_excel('Anwesenheitsliste.xlsx', index=False)
    status_label.config(text='Anwesenheitsliste.xlsx erfolgreich erstellt')

def close_program():
    root.destroy()

root = tk.Tk()
root.geometry('500x550')
root.title('Die Deutsche Presse Agentur präsentiert:')

status_label = tk.Label(root, text='Bereit wenn Sie es sind!', anchor='w', bg='lightblue')
status_label.pack(fill='x', padx=10, pady=10)

year_label = tk.Label(root, text='Jahr:')
year_label.pack(fill='x', padx=10, pady=10)

year_entry = tk.Entry(root)
year_entry.pack(fill='x', padx=10, pady=5)

create_database_button = tk.Button(root, text='Database erstellen', command=create_database)
create_database_button.pack(fill='x', padx=10, pady=5)

read_excel_button = tk.Button(root, text='Wahl der Schüler laden', command=read_excel_file)
read_excel_button.pack(fill='x', padx=10, pady=5)

read_excel_button_2 = tk.Button(root, text='Raumliste laden', command=read_excel_file_2)
read_excel_button_2.pack(fill='x', padx=10, pady=5)

read_excel_button_3 = tk.Button(root, text='Veranstaltungsliste laden', command=read_excel_file_3)
read_excel_button_3.pack(fill='x', padx=10, pady=5)

distribute_button = tk.Button(root, text='Schüler verteilen', command=distribute_students)
distribute_button.pack(fill='x', padx=10, pady=5)

create_excel_button = tk.Button(root, text='Listen erstellen', command=create_excel_files)
create_excel_button.pack(fill='x', padx=10, pady=10)

result_label1 = tk.Label(root, text='', anchor='w', bg='black')
result_label1.pack(fill='x', padx=10, pady=10)

result_label2 = tk.Label(root, text='', anchor='w', bg='red')
result_label2.pack(fill='x', padx=10, pady=10)

result_label3 = tk.Label(root, text='', anchor='w', bg='yellow')
result_label3.pack(fill='x', padx=10, pady=10)

close_button = tk.Button(root, text='Schließen', command=close_program)
close_button.pack(fill='x', padx=10, pady=5)

root.mainloop()

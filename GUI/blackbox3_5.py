import tkinter as tk
from tkinter import filedialog
import pandas as pd

def read_excel_file(list):
    file_path = filedialog.askopenfilename(filetypes=[("csv files", "*.csv")])
    
    if file_path:
        try:
            if list == 'Wahl':
                df_wahl = pd.read_csv(file_path, delimiter=';')
                status_label.config(text='Datei erfolgreich geladen')
            elif list == 'Veranstaltung':
                df_veranstaltung = pd.read_csv(file_path, delimiter=';')
                status_label.config(text='Datei erfolgreich geladen')
            elif list == 'Raum':
                df_raum = pd.read_csv(file_path, delimiter=';')
                status_label.config(text='Datei erfolgreich geladen')
        except Exception as e:
            status_label.config(text=f'Error: {str(e)}')

def distribute_students():
    pass

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

read_excel_button = tk.Button(root, text='Wahl der Schüler laden', command=read_excel_file('Wahl'))
read_excel_button.pack(fill='x', padx=10, pady=5)

read_excel_button_2 = tk.Button(root, text='Raumliste laden', command=read_excel_file('Raum'))
read_excel_button_2.pack(fill='x', padx=10, pady=5)

read_excel_button_3 = tk.Button(root, text='Veranstaltungsliste laden', command=read_excel_file('Veranstaltung'))
read_excel_button_3.pack(fill='x', padx=10, pady=5)

distribute_button = tk.Button(root, text='Schüler verteilen', command=distribute_students)
distribute_button.pack(fill='x', padx=10, pady=5)

create_excel_button = tk.Button(root, text='Listen erstellen', command=create_excel_files)
create_excel_button.pack(fill='x', padx=10, pady=10)

result_label1 = tk.Label(root, text='Deutsche', anchor='w', bg='black', fg='white', justify='center')
result_label1.pack(fill='x', padx=10, pady=10) 

result_label2 = tk.Label(root, text='Presse', anchor='w', bg='red', justify='center')
result_label2.pack(fill='x', padx=10, pady=10)

result_label3 = tk.Label(root, text='Agentur', anchor='w', bg='yellow', justify='center')
result_label3.pack(fill='x', padx=10, pady=10)

close_button = tk.Button(root, text='Schließen', command=close_program)
close_button.pack(fill='x', padx=10, pady=5)

root.mainloop()

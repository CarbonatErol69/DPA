import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import sqlite3

class EventSchedulerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Schedulize")
        self.geometry("500x500")
        self.resizable(1, 1)

        # Initialize SQLite database connection
        self.connection = sqlite3.connect("schedulize.db")
        self.cursor = self.connection.cursor()

        # Create student table if not exists
        self.create_student_table()

        # Create and place widgets
        self.create_widgets()

    def create_student_table(self):
        # Create student table if not exists
        create_table_query = '''
            CREATE TABLE IF NOT EXISTS Student (
                Name TEXT,
                Vorname TEXT,
                Klasse TEXT,
                Wahl1 TEXT,
                Wahl2 TEXT,
                Wahl3 TEXT,
                Wahl4 TEXT,
                Wahl5 TEXT,
                Wahl6 TEXT
            )
        '''
        self.cursor.execute(create_table_query)
        self.connection.commit()

    def create_widgets(self):
        # Your widget creation code remains unchanged
        pass

    def add_student(self):
        # Retrieve student information from entry fields
        student_name = self.entry_student_name.get()
        student_firstname = self.entry_first_name.get()
        student_class = self.entry_class.get()
        event_preferences = [entry.get() for entry in self.entry_event_preferences]

        # Insert student data into the SQLite database
        insert_query = '''
            INSERT INTO Student (Name, Vorname, Klasse, Wahl1, Wahl2, Wahl3, Wahl4, Wahl5, Wahl6)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        '''
        self.cursor.execute(insert_query, (student_name, student_firstname, student_class, *event_preferences))
        self.connection.commit()

        # Clear entry fields
        self.entry_student_name.delete(0, tk.END)
        self.entry_first_name.delete(0, tk.END)
        self.entry_class.delete(0, tk.END)
        for entry in self.entry_event_preferences:
            entry.delete(0, tk.END)

if __name__ == "__main__":
    app = EventSchedulerApp()
    app.mainloop()

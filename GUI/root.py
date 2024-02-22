import tkinter as tk
import sqlite3

class MainPage(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent

        # Create SQLite database connection
        self.db_conn = sqlite3.connect('schedulize.db')

        # Create label
        self.label = tk.Label(self, text="Welcome to this beautiful application that does stuff :) !")

        # Create buttons
        self.add_student_button = tk.Button(self, text="Add student", command=self.go_add_student)
        self.add_company_button = tk.Button(self, text="Add company", command=self.go_add_company)
        self.make_list_button = tk.Button(self, text="Make list", command=self.make_list)
        self.close_button = tk.Button(self, text="Close", command=self.close_program)
        self.delete_student_button = tk.Button(self, text="Delete student", command=self.delete_student)
        self.delete_company_button = tk.Button(self, text="Delete company", command=self.delete_company)

        # Grid for buttons to be placed on
        self.label.grid(row=0, column=0, pady=10)
        self.add_student_button.grid(row=1, column=0, padx=10, pady=10)
        self.add_company_button.grid(row=2, column=0, padx=10, pady=10)
        self.make_list_button.grid(row=3, column=0, padx=10, pady=10)
        self.close_button.grid(row=4, column=0, padx=10, pady=10)
        self.delete_student_button.grid(row=5, column=0, padx=10, pady=10)
        self.delete_company_button.grid(row=6, column=0, padx=10, pady=10)

    def go_add_student(self):
        # Implement logic to add a student to the database
        print("Adding student...")  # Placeholder message

    def go_add_company(self):
        # Implement logic to add a company to the database
        print("Adding company...")  # Placeholder message

    def make_list(self):
        # Implement logic to retrieve data from the database and make a list
        print("Making list...")  # Placeholder message

    def close_program(self):
        # Close the application
        print("Closing program...")
        self.db_conn.close()
        print("Database connection closed.")
        self.parent.destroy()
    
    def delete_student(self):
        # Implement logic to delete a student from the database
        print("Deleting student...")  # Placeholder message

    def delete_company(self):
        # Implement logic to delete a company from the database
        print("Deleting company...") # Placeholder message

if __name__ == "__main__":
    # Creating root window
    root = tk.Tk()
    root.title("Career Orientation")
    root.geometry("500x500")
    root.resizable(0, 0)

    # Create MainPage object and pack it
    main_page = MainPage(root)
    main_page.pack()

    # Update and start mainloop
    root.mainloop()

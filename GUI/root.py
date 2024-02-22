import tkinter as tk

class MainPage(tk.Frame):
    def __init__(self, parent):
        
        tk.Frame.__init__(self, parent)

        # Create label
        self.label = tk.Label(self, text="Welcome to this beautiful application that does shit all :) !")

        # Create buttons
        self.add_student_button = tk.Button(self, text="Add student", command=self.go_add_student)
        self.add_company_button = tk.Button(self, text="Add company", command=self.go_add_company)
        self.make_list_button = tk.Button(self, text="Make list", command=self.make_list)
        self.close_button = tk.Button(self, text="Close", command=self.close_program)
        self.delete_student_button = tk.Button(self, text="Delete student", command=self.delete_student)
        self.delete_company_button = tk.Button(self, text="Delete company", command=self.delete_company)

        # Grid fo buttons to be placed on
        self.label.grid(row=0, column=0, pady=10)
        self.add_student_button.grid(row=1, column=0, padx=10, pady=10)
        self.add_company_button.grid(row=2, column=0, padx=10, pady=10)
        self.make_list_button.grid(row=3, column=0, padx=10, pady=10)
        self.close_button.grid(row=4, column=0, padx=10, pady=10)
        self.delete_student_button.grid(row=5, column=0, padx=10, pady=10)
        self.delete_company_button.grid(row=6, column=0, padx=10, pady=10)

    def go_add_student(self):
        # Implement logic
        print("Adding student...")  # Placeholder message

    def go_add_company(self):
        # Implement logic
        print("Adding company...")  # Placeholder message

    def make_list(self):
        # Implement logic
        print("Making list...")  # Placeholder message

    def close_program(self):
        # Close the application
        root.destroy()
    
    def delete_student(self):
        #logic here
        print("Deleting student...")  # Placeholder message

    def delete_company(self):
        #logic here
        print("Deleting company...") # Placeholder message

if __name__ == "__main__":
    # Creating root window
    root = tk.Tk()
    root.title("DPAvsBWV")
    root.geometry("500x500")
    root.resizable(0, 0)

    # Create MainPage object and pack it
    main_page = MainPage(root)
    main_page.pack()

    # Update and start mainloop
    root.mainloop()

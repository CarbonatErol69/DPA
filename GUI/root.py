import tkinter as tk

#creating root window
root = tk.Tk()
root.title("DPAvsBWV")
root.geometry("500x500")
root.resizable(0,0)

class MainPage(tk.Frame):
    def __init__(self, parent):

        #call the constructor of the master class
        tk.Frame.__init__(self, parent)

        #create labl widgets
        self.label = tk.Label(self, text="Welcome to DPAvsBWV")

        #create all buttons
        self.add_student_button = tk.Button(self, text="Add student", command=self.go_add_student)
        self.add_company_button = tk.Button(self, text="Add company", command=self.go_add_company)
        self.make_list_button = tk.Button(self, text="Make list", command=self.make_list)
        self.close_button = tk.Button(self, text="Close", command=self.close_program)

        #grid to place buttons
        self.label.grid(row=0, column=0, pady=10)
        self.back_button.grid(row=1, column=0, padx=10, pady=10)


root.mainloop()
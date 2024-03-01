import tkinter as tk

def button_clicked():
    print("Button pressed!")
    def make_button_funny():
        button.config(text="fun!", bg="yellow", fg="purple")

    button.config(command=make_button_funny)

root = tk.Tk()
root.title("Button Program")

button = tk.Button(root, text="PRESS ME!", command=button_clicked, bg="red", fg="white", font=("Arial", 20))
button.pack(pady=50)

root.mainloop()
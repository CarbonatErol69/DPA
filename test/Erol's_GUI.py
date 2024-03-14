import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import os

def open_excel_file(button_num):
    initial_dir = r'H:\Schule\03 Oberstufe\SuD\00 Projekt\DPA'
    file_path = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xlsx")], initialdir=initial_dir)
    if file_path:
        try:
            # Öffne die Excel-Datei mit openpyxl
            workbook = openpyxl.load_workbook(file_path)
            # Zeige eine Meldung, dass die Datei geöffnet wurde
            messagebox.showinfo("Excel-Datei geöffnet", "Die Excel-Datei wurde erfolgreich geöffnet.")
            # Hier kannst du die Daten der Excel-Datei verarbeiten oder an eine andere Funktion übergeben
            process_excel_data(workbook, button_num, file_path)
        except FileNotFoundError:
            # Zeige eine Meldung, wenn die Datei nicht gefunden wurde
            messagebox.showerror("Fehler", "Die Excel-Datei wurde nicht gefunden.")
        except Exception as e:
            # Zeige eine Meldung, wenn ein anderer Fehler aufgetreten ist
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {str(e)}")

# Funktion zur Verarbeitung der Excel-Daten
def process_excel_data(workbook, button_num, file_path):
    # Hier kannst du die Daten der Excel-Datei verarbeiten oder an eine andere Funktion übergeben
    if button_num == 1:
        file_path_var1.set(os.path.basename(file_path))
    elif button_num == 2:
        file_path_var2.set(os.path.basename(file_path))
    elif button_num == 3:
        file_path_var3.set(os.path.basename(file_path))

# Funktion zum Löschen des Dateipfads
def clear_file_path(button_num):
    if button_num == 1:
        file_path_var1.set("")
    elif button_num == 2:
        file_path_var2.set("")
    elif button_num == 3:
        file_path_var3.set("")

# Funktion zur Übernahme der Daten
def schnittstelle(): #TODO: Hier muss noch die Funktion eingefügt werden, die die Daten übernimmt
    # Hier kannst du die ausgewählten Daten übernehmen und weiterverarbeiten
    

# Erstelle das Hauptfenster
root = tk.Tk()
root.title("Deutsche Presse Agentur")

# Funktionen für die weiteren Buttons
def open_excel_file_2():
    open_excel_file(2)

def open_excel_file_3():
    open_excel_file(3)

# Erstelle die Textvariablen für die ausgewählten Dateipfade
file_path_var1 = tk.StringVar()
file_path_var2 = tk.StringVar()
file_path_var3 = tk.StringVar()

# Erstelle den ersten Button und füge ihn dem Hauptfenster hinzu
open_excel_button = tk.Button(root, text="Wählen Sie hier die Excel-Datei der Wahl aus!", command=lambda: open_excel_file(1))
open_excel_button.pack(padx=10, pady=5, fill=tk.X)
file_path_entry1 = tk.Entry(root, textvariable=file_path_var1, state='readonly', width=50)
file_path_entry1.pack(padx=10, pady=5, fill=tk.X)
delete_button1 = tk.Button(file_path_entry1, text="X", command=lambda: clear_file_path(1), width=2)
delete_button1.pack(side=tk.RIGHT)

# Erstelle den zweiten Button und füge ihn dem Hauptfenster hinzu
open_excel_button_2 = tk.Button(root, text="Wählen Sie hier die Excel-Datei der Räume aus!", command=open_excel_file_2)
open_excel_button_2.pack(padx=10, pady=5, fill=tk.X)
file_path_entry2 = tk.Entry(root, textvariable=file_path_var2, state='readonly', width=50)
file_path_entry2.pack(padx=10, pady=5, fill=tk.X)
delete_button2 = tk.Button(file_path_entry2, text="X", command=lambda: clear_file_path(2), width=2)
delete_button2.pack(side=tk.RIGHT)

# Erstelle den dritten Button und füge ihn dem Hauptfenster hinzu
open_excel_button_3 = tk.Button(root, text="Wählen Sie hier die Excle-Datei der Events aus!", command=open_excel_file_3)
open_excel_button_3.pack(padx=10, pady=5, fill=tk.X)
file_path_entry3 = tk.Entry(root, textvariable=file_path_var3, state='readonly', width=50)
file_path_entry3.pack(padx=10, pady=5, fill=tk.X)
delete_button3 = tk.Button(file_path_entry3, text="X", command=lambda: clear_file_path(3), width=2)
delete_button3.pack(side=tk.RIGHT)

# Erstelle den Button zur Übernahme der Daten und füge ihn dem Hauptfenster hinzu
take_over_button = tk.Button(root, text="Daten übernehmen", command=schnittstelle)
take_over_button.pack(padx=10, pady=10, fill=tk.X)

# Starte die Hauptschleife des Hauptfensters
root.mainloop()
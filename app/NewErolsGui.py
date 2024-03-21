import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import os
from AlexGPT import AlexGPT

class RoundedButton(tk.Button):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.config(relief="flat", borderwidth=0)
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
    
    def on_enter(self, event):
        self.config(bg="#A0A0A0")
    
    def on_leave(self, event):
        self.config(bg="#808080")

class RoundedEntry(tk.Entry):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.config(relief="flat", borderwidth=0)

def open_excel_file(button_num):
    initial_dir = r'C:\Users\Tim.Schmidt\OneDrive - Zentis\02 Schule\03 Oberstufe\SuD\00 Abschlussprojekt\DPA\DPA\Data' # 
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
        file_path_var1.set(file_path) # Nur Dateiname os.path.basename()
    elif button_num == 2:
        file_path_var2.set(file_path)
    elif button_num == 3:
        file_path_var3.set(file_path)

# Funktion zum Löschen des Dateipfads
def clear_file_path(button_num):
    if button_num == 1:
        file_path_var1.set("")
    elif button_num == 2:
        file_path_var2.set("")
    elif button_num == 3:
        file_path_var3.set("")

# Funktion zur Übernahme der Daten
def schnittstelle():
    # Abrufen der ausgewählten Dateipfade
    file_path_1 = file_path_var1.get()
    file_path_2 = file_path_var2.get()
    file_path_3 = file_path_var3.get()
    
    # Überprüfen, ob alle Dateipfade ausgewählt wurden
    if file_path_1 and file_path_2 and file_path_3:
        # Übergeben der Dateipfade an die AlexGPT-Klasse
        alexGPT = AlexGPT(file_path_1, file_path_2, file_path_3)
        # Aufruf der Hauptmethode der AlexGPT-Klasse
        alexGPT.main()
    else:
        # Zeige eine Fehlermeldung, wenn nicht alle Dateipfade ausgewählt wurden
        messagebox.showerror("Fehler", "Bitte wählen Sie alle Excel-Dateien aus.")


# BRAUCHEN WIR DAS???
"""
def take_over_data():
    # Hier kannst du die ausgewählten Daten übernehmen und weiterverarbeiten
    pass
"""

# Erstelle das Hauptfenster
root = tk.Tk()
root.title("Auswahl der Excel-Listen")
root.state('zoomed')  # Fenster maximieren
root.configure(bg="#404040")  # Hintergrundfarbe des Hauptfensters

# Funktionen für die weiteren Buttons
def open_excel_file_2():
    open_excel_file(2)

def open_excel_file_3():
    open_excel_file(3)

# Erstelle das Hauptframe
main_frame = tk.Frame(root, bg="#404040")
main_frame.pack(expand=True, fill="both")

# Erstelle ein unsichtbares Gitter und platziere es in der Mitte des Fensters
center_frame = tk.Frame(main_frame, bg="#404040")
center_frame.pack(expand=True, padx=20, pady=20, anchor="center")

# Erstelle die Textvariablen für die ausgewählten Dateipfade
file_path_var1 = tk.StringVar()
file_path_var2 = tk.StringVar()
file_path_var3 = tk.StringVar()

# Erstelle den ersten Button und füge ihn dem Hauptfenster hinzu
open_excel_button = RoundedButton(center_frame, text="Wählen Sie hier die Excel-Datei der Wahl aus!", command=lambda: open_excel_file(1), bg="#808080", fg="#FFFFFF")
open_excel_button.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
file_path_entry1 = RoundedEntry(center_frame, textvariable=file_path_var1, state='readonly', width=50, bg="#FFFFFF")
file_path_entry1.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
delete_button1 = tk.Button(file_path_entry1, text="X", command=lambda: clear_file_path(1), width=2, bg="#808080", fg="#FFFFFF")
delete_button1.pack(side=tk.RIGHT)

# Erstelle den zweiten Button und füge ihn dem Hauptfenster hinzu
open_excel_button_2 = RoundedButton(center_frame, text="Wählen Sie hier die Excel-Datei der Räume aus!", command=open_excel_file_2, bg="#808080", fg="#FFFFFF")
open_excel_button_2.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
file_path_entry2 = RoundedEntry(center_frame, textvariable=file_path_var2, state='readonly', width=50, bg="#FFFFFF")
file_path_entry2.grid(row=4, column=0, padx=20, pady=5, sticky="ew")
delete_button2 = tk.Button(file_path_entry2, text="X", command=lambda: clear_file_path(2), width=2, bg="#808080", fg="#FFFFFF")
delete_button2.pack(side=tk.RIGHT)

# Erstelle den dritten Button und füge ihn dem Hauptfenster hinzu
open_excel_button_3 = RoundedButton(center_frame, text="Wählen Sie hier die Excle-Datei der Events aus!", command=open_excel_file_3, bg="#808080", fg="#FFFFFF")
open_excel_button_3.grid(row=5, column=0, padx=20, pady=10, sticky="ew")
file_path_entry3 = RoundedEntry(center_frame, textvariable=file_path_var3, state='readonly', width=50, bg="#FFFFFF")
file_path_entry3.grid(row=6, column=0, padx=20, pady=5, sticky="ew")
delete_button3 = tk.Button(file_path_entry3, text="X", command=lambda: clear_file_path(3), width=2, bg="#808080", fg="#FFFFFF")
delete_button3.pack(side=tk.RIGHT)

# Erstelle den Button zur Übernahme der Daten und füge ihn dem Hauptfenster hinzu
take_over_button = RoundedButton(center_frame, text="Daten übernehmen", command=schnittstelle, bg="#808080", fg="#FFFFFF")
take_over_button.grid(row=7, column=0, padx=20, pady=20, sticky="ew")
take_over_button.configure(height=3)  # Ändere die vertikale Größe des Buttons

# Starte die Hauptschleife des Hauptfensters
root.mainloop()

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import openpyxl
import webbrowser

"""
def open_youtube_video():
    webbrowser.open("https://www.youtube.com/watch?v=-9HAB1vypcw")   # Carbonat's Video: https://www.youtube.com/watch?v=gBDQ06An-iU
"""
# Erstelle das Hauptfenster
root = tk.Tk()
root.title("YouTube Video Link")

"""
# Funktion für den Link
def open_youtube():
    open_youtube_video()

# Erstelle das Label als Link
youtube_link_label = tk.Label(root, text="Klick hier, um das Video anzusehen", fg="blue", cursor="hand2")
youtube_link_label.pack(pady=10)
youtube_link_label.bind("<Button-1>", lambda e: open_youtube())
"""


# Starte die Hauptschleife des Hauptfensters
root.mainloop()







def open_excel_file(button_num):
    initial_dir = r'H:\Schule\03 Oberstufe\SuD\00 Projekt\DPA'    # eigenen Pfad auswählen, in dem die Excel-Dateien liegen C:\Users\Tim.Schmidt\OneDrive - Zentis\02 Schule\03 Oberstufe\SuD\00 Abschlussprojekt\DPA\DPA
    file_path = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xlsx")], initialdir=initial_dir)
    if file_path:
        try:
            # Öffne die Excel-Datei mit openpyxl
            workbook = openpyxl.load_workbook(file_path)
            # Zeige eine Meldung, dass die Datei geöffnet wurde
            messagebox.showinfo("Excel-Datei geöffnet", "Die Excel-Datei wurde erfolgreich geöffnet.")
        except FileNotFoundError:
            # Zeige eine Meldung, wenn die Datei nicht gefunden wurde
            messagebox.showerror("Fehler", "Die Excel-Datei wurde nicht gefunden.")
        except Exception as e:
            # Zeige eine Meldung, wenn ein anderer Fehler aufgetreten ist
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {str(e)}")

# Erstelle das Hauptfenster
root = tk.Tk()
root.title("Deutsche Presse Agentur")

# Funktionen für die weiteren Buttons
def open_excel_file_2():
    open_excel_file(2)

def open_excel_file_3():
    open_excel_file(3)


"""
# Zusätzliche Zeile Text 
text_label = tk.Label(root, text="Ich trag nh 100.000 € Uhr am Handgelek!")
text_label.pack(pady=10)
"""

# Erstelle den ersten Button und füge ihn dem Hauptfenster hinzu
open_excel_button = tk.Button(root, text="Wählen Sie hier die Excel-Datei der Wahl aus!", command=lambda: open_excel_file(1))
open_excel_button.pack(padx=10, pady=10)

# Erstelle den zweiten Button und füge ihn dem Hauptfenster hinzu
open_excel_button_2 = tk.Button(root, text="Wählen Sie hier die Excel-Datei der Räume aus!", command=open_excel_file_2)
open_excel_button_2.pack(padx=10, pady=10)

# Erstelle den dritten Button und füge ihn dem Hauptfenster hinzu
open_excel_button_3 = tk.Button(root, text="Wählen Sie hier die Excle-Datei der Events aus!", command=open_excel_file_3)
open_excel_button_3.pack(padx=10, pady=10)

# Starte die Hauptschleife des Hauptfensters
root.mainloop()
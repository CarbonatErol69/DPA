from collections import defaultdict
import numpy as np
import pandas as pd

# Dateipfade
schueler_wahlen_path = 'C://Users//SE//Schulprojekt//DPA//data//IMPORT BOT2_Schülerwahlen.xlsx'
raumliste_path = 'C://Users//SE//Schulprojekt//DPA//data//IMPORT BOT0_Raumliste.xlsx'
veranstaltungsliste_path = 'C://Users//SE//Schulprojekt//DPA//data//IMPORT BOT1_Veranstaltungsliste.xlsx'

# Einlesen der Dateien
schuelerwahlen_df = pd.read_excel(schueler_wahlen_path)
raumliste_df = pd.read_excel(raumliste_path)
veranstaltungsliste_df = pd.read_excel(veranstaltungsliste_path)

# Anzeigen der ersten Zeilen jeder Tabelle, um die Struktur zu verstehen
print(schuelerwahlen_df.head(), raumliste_df.head(), veranstaltungsliste_df.head())

# Konvertiere 'Frühester Zeitpunkt' in numerische Werte, falls noch nicht geschehen
# Beispiel: {"A": 1, "B": 2, ...} - Angepasst an das tatsächliche Format in Ihrer Datei

# Zuweisung von Veranstaltungen zu Räumen und Zeitslots
# Initialisiere eine Struktur für Raum- und Zeitslot-Zuweisungen
event_assignments = []

# Verfügbare Zeitslots pro Raum initialisieren
available_slots_per_room = {room: [1, 2, 3, 4, 5] for room in raumliste_df['Raum']}

# Veranstaltungen basierend auf dem frühesten Zeitpunkt und verfügbaren Räumen zuordnen
for _, event in veranstaltungsliste_df.iterrows():
    for room, available_slots in available_slots_per_room.items():
        if event['Frühester Zeitpunkt'] in available_slots:
            # Weise Raum und Zeitslot zu und entferne den Zeitslot aus der Verfügbarkeit für diesen Raum
            event_assignments.append((event['Nr. '], room, event['Frühester Zeitpunkt']))
            available_slots.remove(event['Frühester Zeitpunkt'])
            break

# Transformiere Zuweisungen in ein DataFrame für eine bessere Handhabung
event_assignments_df = pd.DataFrame(event_assignments, columns=['VeranstaltungsNr', 'Raum', 'Zeitslot'])

print(event_assignments_df.head())





# Schritt 1: Auszählung wieviele Events benötigt werden
# Consolidate all choices into a single Series
all_choices = pd.concat([schuelerwahlen_df[col] for col in schuelerwahlen_df.columns if 'Wahl' in col])

# Count the frequency of each choice
choice_counts = all_choices.value_counts().sort_index()

# Merge choice counts with event information to determine the number of required events
choice_counts_df = pd.DataFrame(choice_counts).reset_index()
choice_counts_df.columns = ['VeranstaltungsNr', 'Anzahl Wünsche']

# Merge with event information to get the max participants per event
event_requirements_df = pd.merge(choice_counts_df, veranstaltungsliste_df, left_on="VeranstaltungsNr", right_on="Nr. ", how="left")

# Calculate the minimum required number of events for each choice based on max participants
event_requirements_df['Mindestanzahl Events'] = (event_requirements_df['Anzahl Wünsche'] / event_requirements_df['Max. Teilnehmer']).apply(np.ceil)

event_requirements_df[['VeranstaltungsNr', 'Unternehmen', 'Fachrichtung', 'Anzahl Wünsche', 'Max. Teilnehmer', 'Mindestanzahl Events']]




# Schritt 2: Vereinfachte Zuweisung von Veranstaltungen zu Räumen

# Nehmen wir an, dass jede Veranstaltung mindestens einmal stattfindet und wir die Veranstaltungen auf verfügbare Räume verteilen.
# Diese Vereinfachung hilft, den Prozess zu starten. Anpassungen für spezifische Zeitfenster und Mehrfachveranstaltungen können später hinzugefügt werden.

# Da wir bereits die Veranstaltungen und Räume geladen haben, weisen wir jedem Event basierend auf der Reihenfolge einen Raum zu.

# Vereinfachte Annahme: Jedes Event bekommt einen Raum, unabhängig von der Anzahl der benötigten Durchführungen.
# Diese Annahme dient nur der Demonstration. Die tatsächliche Implementierung sollte die Mindestanzahl der Events berücksichtigen.

event_raum_zuweisung = {}

# Gehe durch jede Veranstaltung
for index, event in veranstaltungsliste_df.iterrows():
    # Suche nach einem verfügbaren Raum (vereinfacht: der erste verfügbare Raum)
    # In der realen Implementierung sollten Sie die Raumkapazität und andere Anforderungen berücksichtigen
    verfuegbarer_raum = raumliste_df.iloc[index % len(raumliste_df)]  # Wir nutzen Modulo, um bei Überschreitung der Raumliste von vorne zu beginnen
    event_raum_zuweisung[event["Nr. "]] = verfuegbarer_raum["Raum"]

# Konvertiere die Zuordnung in einen DataFrame für eine bessere Darstellung
event_raum_zuweisung_df = pd.DataFrame(list(event_raum_zuweisung.items()), columns=['VeranstaltungsNr', 'Raum'])

event_raum_zuweisung_df.head()



# Schritt 3: Schüler den Veranstaltungen zuweisen

# Initialisiere die Variablen erneut, um sicherzustellen, dass sie im aktuellen Kontext definiert sind
aktuelle_teilnehmerzahl = defaultdict(int)  # Verfolgt die aktuelle Teilnehmerzahl pro Event
schueler_zuweisung = defaultdict(list)  # Endgültige Zuweisung der Schüler zu Veranstaltungen

# Durchlaufe die Wahlen jedes Schülers
for index, schueler in schuelerwahlen_df.iterrows():
    zugewiesen = False
    for wahl in ['Wahl 1', 'Wahl 2', 'Wahl 3']:  # Berücksichtigt nur die ersten drei Wünsche
        if pd.isna(schueler[wahl]):
            continue  # Überspringe, wenn keine Wahl getroffen wurde
        veranstaltungs_nr = int(schueler[wahl])
        # Überprüfe, ob die Kapazität der gewählten Veranstaltung noch nicht erreicht ist
        if aktuelle_teilnehmerzahl[veranstaltungs_nr] < veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == veranstaltungs_nr, 'Max. Teilnehmer'].values[0]:
            # Weise den Schüler dieser Veranstaltung zu und aktualisiere die Teilnehmerzahl
            schueler_zuweisung[schueler['Name']].append(veranstaltungs_nr)
            aktuelle_teilnehmerzahl[veranstaltungs_nr] += 1
            zugewiesen = True
            break  # Breche ab, da der Schüler einer Veranstaltung zugewiesen wurde

# Hier werden alle Schhüler ohne Wahlen verteilt

# Konvertiere die Zuweisungen in einen DataFrame für eine bessere Übersicht
schueler_zuweisung_df = pd.DataFrame([(k, v) for k, v in schueler_zuweisung.items()], columns=['Schüler', 'Zugewiesene VeranstaltungsNrn'])

schueler_zuweisung_df.head()





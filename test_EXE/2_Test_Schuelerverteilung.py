from collections import defaultdict
import numpy as np
import pandas as pd
import os
import random

# Daten einlesen
# Dateipfade
basisimportpfad = 'C://Users//SE//Schulprojekt//DPA//data//'
schueler_wahlen_path = basisimportpfad + 'IMPORT_BOT2_Wahl.xlsx'
raumliste_path = basisimportpfad +'IMPORT_BOT0_Raumliste.xlsx'
veranstaltungsliste_path =  basisimportpfad + 'IMPORT_BOT1_Veranstaltungsliste.xlsx'

# Basispfad für den Export festlegen
export_basispfad = 'C://Users//SE//Desktop//Laufzettel//'
anwesenheitslisten_path = export_basispfad +'EXPORT_BOT5_Anwesenheitslisten.xlsx'
raum_zeitplan_path = export_basispfad + 'EXPORT_BOT3_Raum_und_Zeitplan.xlsx'

# Einlesen der Dateien in Dataframes
schuelerwahlen_df = pd.read_excel(schueler_wahlen_path)
raumliste_df = pd.read_excel(raumliste_path)
veranstaltungsliste_df = pd.read_excel(veranstaltungsliste_path)


# Preprocessing
# Define the convert_dict_to_df function
def convert_dict_to_df(schueler_zuweisung):
    schueler_zuweisung_list = [{'Schüler': k, 'Zugewiesene VeranstaltungsNrn': v} for k, v in schueler_zuweisung.items()]
    return pd.DataFrame(schueler_zuweisung_list)

# Extrahiere die ersten fünf Wahlen
first_five_choices = pd.concat([
    schuelerwahlen_df['Wahl 1'], schuelerwahlen_df['Wahl 2'], 
    schuelerwahlen_df['Wahl 3'], schuelerwahlen_df['Wahl 4'], 
    schuelerwahlen_df['Wahl 5']
])
# Zähle, wie oft jede Veranstaltung gewählt wurde
event_counts = first_five_choices.value_counts()

# Überprüfe die Spaltennamen
print(schuelerwahlen_df.columns)

# Falls der Spaltenname 'Schüler' nicht vorhanden ist, korrigiere ihn
if 'Schüler' not in schuelerwahlen_df.columns:
    # Annahme: Vorname und Nachname sind in separaten Spalten 'Vorname' und 'Nachname'
    schuelerwahlen_df['Schüler'] = schuelerwahlen_df['Vorname'] + ' ' + schuelerwahlen_df['Name']

# Überprüfe die aktualisierten Spaltennamen
print(schuelerwahlen_df.columns)

# Füge eine neue Spalte für die Zählung der Events in 'veranstaltungsliste_df' hinzu
veranstaltungsliste_df['Event Counts'] = veranstaltungsliste_df['Nr. '].map(event_counts).fillna(0)

# Berechne 'Mindestanzahl Events' basierend auf 'Max. Teilnehmer' und 'Event Counts'
veranstaltungsliste_df['Mindestanzahl Events'] = np.ceil(veranstaltungsliste_df['Event Counts'] / veranstaltungsliste_df['Max. Teilnehmer'])
print(veranstaltungsliste_df.head)

# Unmittelbar vor dem Zugriff auf 'Wahl 6'
if 'Wahl 6' not in schuelerwahlen_df.columns:
    schuelerwahlen_df['Wahl 6'] = np.nan  # Füge 'Wahl 6' mit NaN-Werten hinzu

# Convert 'Frühester Zeitpunkt' to numeric values
zeitpunkt_mapping = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5}
veranstaltungsliste_df['Frühester Zeitpunkt Num'] = veranstaltungsliste_df['Frühester Zeitpunkt'].map(zeitpunkt_mapping)

# Calculating 'Mindestanzahl Events'

# Berechnungen für Veranstaltungsliste
event_counts = pd.concat([schuelerwahlen_df['Wahl 1'], schuelerwahlen_df['Wahl 2'], schuelerwahlen_df['Wahl 3']]).value_counts()
veranstaltungsliste_df['Event Counts'] = veranstaltungsliste_df['Nr. '].map(event_counts).fillna(0)
veranstaltungsliste_df['Mindestanzahl Events'] = np.ceil(veranstaltungsliste_df['Event Counts'] / veranstaltungsliste_df['Max. Teilnehmer'])

# Display the adjusted 'veranstaltungsliste_df' for verification
veranstaltungsliste_df


# Konvertiere 'Frühester Zeitpunkt' in numerische Werte
zeitpunkt_mapping = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5}
veranstaltungsliste_df['Frühester Zeitpunkt Num'] = veranstaltungsliste_df['Frühester Zeitpunkt'].map(zeitpunkt_mapping)

# Schritt 1: Auszählung wieviele Events benötigt werden
def count_courses_by_field(veranstaltungsliste_df):
    course_counts = veranstaltungsliste_df.groupby('Fachrichtung').size().reset_index(name='Anzahl Kurse')
    return course_counts

course_counts_df = count_courses_by_field(veranstaltungsliste_df)
print(course_counts_df)

# Schritt 2: Veranstaltungen den Räumen und Zeitslots zuordnen
def assign_courses_to_rooms(veranstaltungsliste_df, raumliste_df):
    veranstaltungsliste_df.sort_values(by=['Unternehmen', 'Fachrichtung', 'Frühester Zeitpunkt Num'], inplace=True)
    raum_zeitplan = []
    
    for _, veranstaltung in veranstaltungsliste_df.iterrows():
        veranstaltungs_id = veranstaltung['Nr. ']
        unternehmen = veranstaltung['Unternehmen']
        fachrichtung = veranstaltung['Fachrichtung']
        mindest_anzahl_events = int(veranstaltung['Mindestanzahl Events']) 

        for _, raum in raumliste_df.iterrows():
            raum_name = raum['Raum']
            for zeitslot in range(1, 6):
                raum_zeitplan.append((veranstaltungs_id, raum_name, zeitslot))
                mindest_anzahl_events -= 1
                if mindest_anzahl_events == 0:
                    break
            if mindest_anzahl_events == 0:
                break

    return pd.DataFrame(raum_zeitplan, columns=['VeranstaltungsNr', 'Raum', 'Zeitslot'])

def zuordnung_veranstaltungen_zu_raeumen(veranstaltungsliste_df, raumliste_df):
    return assign_courses_to_rooms(veranstaltungsliste_df, raumliste_df)

veranstaltung_zeitslot_raum_df = zuordnung_veranstaltungen_zu_raeumen(veranstaltungsliste_df, raumliste_df)


# Zuordnungsfunktion
def assign_courses_to_students(schuelerwahlen_df, veranstaltungsliste_df, veranstaltung_zeitslot_raum_df):
    schueler_zuweisung = defaultdict(list)
    # Copy the DataFrame to preserve the original
    remaining_veranstaltungsliste_df = veranstaltungsliste_df.copy()
    for _, row in schuelerwahlen_df.iterrows():
        schueler_name = row['Name']
        wuensche = [row['Wahl 1'], row['Wahl 2'], row['Wahl 3'], row['Wahl 4'], row['Wahl 5'], row['Wahl 6']]
        # Shuffle the wishes to randomize selection
        random.shuffle(wuensche)
        for wahl_nr in wuensche:
            if pd.notna(wahl_nr) and wahl_nr in remaining_veranstaltungsliste_df['Nr. '].values:
                # Check if there are available slots for the event
                slots_available = veranstaltung_zeitslot_raum_df[veranstaltung_zeitslot_raum_df['VeranstaltungsNr'] == wahl_nr]
                if not slots_available.empty:
                    # Assign the student to the event
                    schueler_zuweisung[schueler_name].append(wahl_nr)
                    # Remove the assigned event from remaining events
                    remaining_veranstaltungsliste_df = remaining_veranstaltungsliste_df[remaining_veranstaltungsliste_df['Nr. '] != wahl_nr]
    # If there are remaining events, assign students randomly
    if not remaining_veranstaltungsliste_df.empty:
        for _, row in schuelerwahlen_df.iterrows():
            schueler_name = row['Name']
            for _ in range(6 - len(schueler_zuweisung[schueler_name])):
                # Randomly select an available event
                random_event = random.choice(remaining_veranstaltungsliste_df['Nr. '].values)
                # Assign the student to the event
                schueler_zuweisung[schueler_name].append(random_event)
                # Remove the assigned event from remaining events
                remaining_veranstaltungsliste_df = remaining_veranstaltungsliste_df[remaining_veranstaltungsliste_df['Nr. '] != random_event]
                if remaining_veranstaltungsliste_df.empty:
                    break
            if remaining_veranstaltungsliste_df.empty:
                break
    return schueler_zuweisung

# Usage example
schueler_zuweisung = assign_courses_to_students(schuelerwahlen_df, veranstaltungsliste_df, veranstaltung_zeitslot_raum_df)
schueler_zuweisung_df = convert_dict_to_df(schueler_zuweisung)
print(schueler_zuweisung_df)


# Konvertieren Sie das defaultdict zu einem DataFrame
def convert_dict_to_df(schueler_zuweisung):
    schueler_zuweisung_list = [{'Schüler': k, 'Zugewiesene VeranstaltungsNrn': v} for k, v in schueler_zuweisung.items()]
    return pd.DataFrame(schueler_zuweisung_list)

schueler_zuweisung_df = convert_dict_to_df(schueler_zuweisung)


def export_klassenweise_laufzettel(schueler_zuweisung_df, veranstaltungsliste_df, schuelerwahlen_df, export_basispfad):
  for klasse in schuelerwahlen_df['Klasse'].unique():
    schueler_der_klasse = schuelerwahlen_df[schuelerwahlen_df['Klasse'] == klasse]
    zugewiesene_veranstaltungen = schueler_zuweisung_df.merge(schueler_der_klasse, on='Schüler')
    laufzettel_df = zugewiesene_veranstaltungen.merge(veranstaltungsliste_df, left_on='VeranstaltungsNr', right_on='Nr. ', how='left').sort_values(by=['Zeitslot'])
    laufzettel_pfad = os.path.join(export_basispfad, klasse, "Laufzettel_alle.xlsx")
    laufzettel_df.to_excel(laufzettel_pfad, index=False)



# Überarbeitete Funktion export_data
def export_data(schueler_zuweisung_df, veranstaltung_zeitslot_raum_df, schuelerwahlen_df, veranstaltungsliste_df, raum_zeitplan_path, anwesenheitslisten_path, export_basispfad):

    export_klassenweise_laufzettel(schueler_zuweisung_df, veranstaltungsliste_df, schuelerwahlen_df, export_basispfad)


    # Sicher stellen, dass schueler_zuweisung_df ein DataFrame ist
    print('Debugging: schueler_zuweisung_df:')
    print(schueler_zuweisung_df.head())
    if isinstance(schueler_zuweisung_df, pd.DataFrame):
        # 1. Anwesenheitslisten je Veranstaltung
        anwesenheitslisten = defaultdict(list)
        for index, row in schueler_zuweisung_df.iterrows():
            schueler = row['Schüler']
            veranstaltungsnummern = row['Zugewiesene VeranstaltungsNrn']
            for veranstaltungsnummer in veranstaltungsnummern:
                anwesenheitslisten[veranstaltungsnummer].append(schueler)
        pass
    else:
        raise ValueError("schueler_zuweisung_df ist kein pandas DataFrame. Bitte überprüfe die Konvertierung.")

    anwesenheitslisten_df = pd.DataFrame([(veranstaltungsnummer, ", ".join(schueler)) for veranstaltungsnummer, schueler in anwesenheitslisten.items()], columns=['VeranstaltungsNr', 'Teilnehmende Schüler'])
    anwesenheitslisten_df.to_excel(anwesenheitslisten_path, index=False)

    # 2. Raum- und Zeitplan
    raum_zeitplan_df = veranstaltung_zeitslot_raum_df.merge(veranstaltungsliste_df, left_on='VeranstaltungsNr', right_on='Nr. ', how='left')
    raum_zeitplan_df = raum_zeitplan_df[['VeranstaltungsNr', 'Raum', 'Zeitslot', 'Unternehmen', 'Fachrichtung']]
    raum_zeitplan_df.to_excel(raum_zeitplan_path, index=False)

    # 3. Laufzettel für Schüler
    for index, row in schueler_zuweisung_df.iterrows():
        schueler = row['Schüler']
        veranstaltungs_nrn = row['Zugewiesene VeranstaltungsNrn']
        klasse = schuelerwahlen_df.loc[schuelerwahlen_df['Name'] == schueler, 'Klasse'].iloc[0]
        klassenpfad = os.path.join(export_basispfad, klasse)
        if not os.path.exists(klassenpfad):
            os.makedirs(klassenpfad)
        veranstaltungen_details = veranstaltung_zeitslot_raum_df[veranstaltung_zeitslot_raum_df['VeranstaltungsNr'].isin(veranstaltungs_nrn)]
        veranstaltungen_details = veranstaltungen_details.merge(veranstaltungsliste_df, left_on='VeranstaltungsNr', right_on='Nr. ', how='left').sort_values(by=['Zeitslot'])
        laufzettel = []
        for _, detail_row in veranstaltungen_details.iterrows():
            laufzettel.append([schueler, detail_row['VeranstaltungsNr'], detail_row['Raum'], detail_row['Zeitslot'], detail_row['Unternehmen'], detail_row['Fachrichtung']])
        laufzettel_df = pd.DataFrame(laufzettel, columns=['Schüler', 'VeranstaltungsNr', 'Raum', 'Zeitslot', 'Unternehmen', 'Fachrichtung'])
        laufzettel_pfad = os.path.join(klassenpfad, f"Laufzettel_{schueler.replace(' ', '_')}.xlsx")
        laufzettel_df.to_excel(laufzettel_pfad, index=False)

# Anwendung der Funktionen
# Generieren Sie veranstaltung_zeitslot_raum_df
veranstaltung_zeitslot_raum_df = zuordnung_veranstaltungen_zu_raeumen(veranstaltungsliste_df, raumliste_df)

# Weisen Sie Kurse den Schülern zu
schueler_zuweisung_df = assign_courses_to_students(schuelerwahlen_df, veranstaltungsliste_df, veranstaltung_zeitslot_raum_df)

schueler_zuweisung_df = convert_dict_to_df(schueler_zuweisung)

# Exportiere die Daten
export_data(schueler_zuweisung_df, veranstaltung_zeitslot_raum_df, schuelerwahlen_df, veranstaltungsliste_df, raum_zeitplan_path, anwesenheitslisten_path, export_basispfad)

# Ausgabe der Pfade zu den erstellten Excel-Dateien
print(f"Anwesenheitslisten gespeichert unter: {anwesenheitslisten_path}")
print(f"Raum- und Zeitplan gespeichert unter: {raum_zeitplan_path}")

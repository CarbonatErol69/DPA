from collections import defaultdict
import numpy as np
import pandas as pd
import os


# Dateipfade
basisimportpfad = 'C://Users//SE//Schulprojekt//DPA//data//'
schueler_wahlen_path = basisimportpfad + 'IMPORT_BOT2_Wahl.xlsx'
raumliste_path = basisimportpfad +'IMPORT_BOT0_Raumliste.xlsx'
veranstaltungsliste_path =  basisimportpfad + 'IMPORT_BOT1_Veranstaltungsliste.xlsx'


# Einlesen der Dateien in Dataframes
schuelerwahlen_df = pd.read_excel(schueler_wahlen_path)
raumliste_df = pd.read_excel(raumliste_path)
veranstaltungsliste_df = pd.read_excel(veranstaltungsliste_path)

klasse_filter = schuelerwahlen_df['Name'] == 'Schüler1'
if klasse_filter.any():
    klasse = schuelerwahlen_df.loc[klasse_filter, 'Klasse'].iloc[0]
else:
    print('No student found with the name "Schüler1"')

print('Debugging: schuelerwahlen_df')
print(schuelerwahlen_df.columns)
print(schuelerwahlen_df.head())


if 'Wahl 6' not in schuelerwahlen_df.columns:
    schuelerwahlen_df['Wahl 6'] = None  # Oder einen anderen angemessenen Standardwert


# Extrahiere die ersten drei Wahlen
first_three_choices = pd.concat([schuelerwahlen_df['Wahl 1'], schuelerwahlen_df['Wahl 2'], schuelerwahlen_df['Wahl 3']])
# Zähle, wie oft jede Veranstaltung gewählt wurde
event_counts = first_three_choices.value_counts()

# Füge eine neue Spalte für die Zählung der Events in 'veranstaltungsliste_df' hinzu
veranstaltungsliste_df['Event Counts'] = veranstaltungsliste_df['Nr. '].map(event_counts).fillna(0)

# Berechne 'Mindestanzahl Events' basierend auf 'Max. Teilnehmer' und 'Event Counts'
veranstaltungsliste_df['Mindestanzahl Events'] = np.ceil(veranstaltungsliste_df['Event Counts'] / veranstaltungsliste_df['Max. Teilnehmer'])

# Unmittelbar vor dem Zugriff auf 'Wahl 6'
if 'Wahl 6' not in schuelerwahlen_df.columns:
    schuelerwahlen_df['Wahl 6'] = np.nan  # Füge 'Wahl 6' mit NaN-Werten hinzu


schuelerwahlen_df = pd.DataFrame({
    'Klasse': ['Klasse1', 'Klasse2'],
    'Name': ['Schüler1', 'Schüler2'],
    'Wahl 1': [1, 2],
    'Wahl 2': [2, 3],
    'Wahl 3': [1, 4],
    'Wahl 4': [3, 1],
    'Wahl 5': [4, 2],
    'Wahl 6': [None, None]  # Stelle sicher, dass 'Wahl 6' existiert
})




veranstaltungsliste_df = pd.DataFrame({
    'Nr. ': [1, 2, 3, 4],
    'Unternehmen': ['Unternehmen1', 'Unternehmen2', 'Unternehmen1', 'Unternehmen2'],
    'Fachrichtung': ['Fach1', 'Fach2', 'Fach3', 'Fach4'],
    'Max. Teilnehmer': [20, 25, 20, 25],
    'Frühester Zeitpunkt': ['A', 'B', 'C', 'D']
})

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
def assign_courses_to_students(schuelerwahlen_df, veranstaltungsliste_df):
    schueler_zuweisung = defaultdict(list)
    for _, row in schuelerwahlen_df.iterrows():
        schueler_name = row['Name']
        for wahl in ['Wahl 1', 'Wahl 2', 'Wahl 3', 'Wahl 4', 'Wahl 5', 'Wahl 6']:
            wahl_nr = row[wahl]
            if pd.notna(wahl_nr) and wahl_nr in veranstaltungsliste_df['Nr. '].values:
                schueler_zuweisung[schueler_name].append(wahl_nr)
    
    return schueler_zuweisung

# Verwende die Funktion, um Schülern Veranstaltungen zuzuweisen
schueler_zuweisung = assign_courses_to_students(schuelerwahlen_df, veranstaltungsliste_df)
schueler_zuweisung_df = pd.DataFrame(list(schueler_zuweisung.items()), columns=['Schüler', 'Zugewiesene VeranstaltungsNrn'])
print(schueler_zuweisung_df)


# def assign_courses_to_students(schuelerwahlen_df, veranstaltungsliste_df):
#     schueler_zuweisung = []
#     for index, schueler in schuelerwahlen_df['Name'].items():
#         klasse = schuelerwahlen_df.loc[schuelerwahlen_df['Name'] == schueler, 'Klasse'].iloc[0]
#         wahl1 = schuelerwahlen_df.loc[schuelerwahlen_df['Name'] == schueler, 'Wahl 1'].iloc[0]
#         wahl2 = schuelerwahlen_df.loc[schuelerwahlen_df['Name'] == schueler, 'Wahl 2'].iloc[0]
#         wahl3 = schuelerwahlen_df.loc[schuelerwahlen_df['Name'] == schueler, 'Wahl 3'].iloc[0]
#         wahl4 = schuelerwahlen_df.loc[schuelerwahlen_df['Name'] == schueler, 'Wahl 4'].iloc[0]
#         wahl5 = schuelerwahlen_df.loc[schuelerwahlen_df['Name'] == schueler, 'Wahl 5'].iloc[0]
#         wahl6 = schuelerwahlen_df.loc[schuelerwahlen_df['Name'] == schueler, 'Wahl 6'].iloc[0]
#         wahl_liste = [wahl1, wahl2, wahl3, wahl4, wahl5, wahl6]
#         wahl_liste = [wahl for wahl in wahl_liste if wahl != 0]
#         veranstaltungen = veranstaltungsliste_df[veranstaltungsliste_df['Fachrichtung'].isin(wahl_liste)]
#         veranstaltungen = veranstaltungen.reset_index(drop=True)
#         veranstaltungen_zuweisung = assign_courses_to_classes(veranstaltungen, klasse)
#         schueler_zuweisung.append(veranstaltungen_zuweisung)
#     return schueler_zuweisung

# schueler_zuweisung = assign_courses_to_students(schuelerwahlen_df, veranstaltungsliste_df)
# schueler_zuweisung_df = pd.DataFrame([(k, v) for k, v in schueler_zuweisung.items()], columns=['Schüler', 'Zugewiesene VeranstaltungsNrn'])
# print(schueler_zuweisung_df.head())

# Konvertieren Sie das defaultdict zu einem DataFrame
def convert_dict_to_df(schueler_zuweisung):
    schueler_zuweisung_list = [{'Schüler': k, 'Zugewiesene VeranstaltungsNrn': v} for k, v in schueler_zuweisung.items()]
    return pd.DataFrame(schueler_zuweisung_list)

schueler_zuweisung_df = convert_dict_to_df(schueler_zuweisung)

# Überarbeitete Funktion export_data
def export_data(schueler_zuweisung_df, veranstaltung_zeitslot_raum_df, schuelerwahlen_df, veranstaltungsliste_df, raum_zeitplan_path, anwesenheitslisten_path, export_basispfad):
    # Stellen Sie sicher, dass schueler_zuweisung_df tatsächlich ein DataFrame ist
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
        raise ValueError("schueler_zuweisung_df ist kein pandas DataFrame. Bitte überprüfen Sie die Konvertierung.")

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
veranstaltung_zeitslot_raum_df = zuordnung_veranstaltungen_zu_raeumen(veranstaltungsliste_df, raumliste_df)
schueler_zuweisung_df = assign_courses_to_students(schuelerwahlen_df, veranstaltungsliste_df)

# Basispfad für den Export festlegen
export_basispfad = 'C://Users//SE//Desktop//Laufzettel//'
anwesenheitslisten_path = export_basispfad +'EXPORT_BOT5_Anwesenheitslisten.xlsx'
raum_zeitplan_path = export_basispfad + 'EXPORT_BOT3_Raum_und_Zeitplan.xlsx'

schueler_zuweisung_df = convert_dict_to_df(schueler_zuweisung)

# Exportiere die Daten
export_data(schueler_zuweisung_df, veranstaltung_zeitslot_raum_df, schuelerwahlen_df, veranstaltungsliste_df, raum_zeitplan_path, anwesenheitslisten_path, export_basispfad)

# Ausgabe der Pfade zu den erstellten Excel-Dateien
print(f"Anwesenheitslisten gespeichert unter: {anwesenheitslisten_path}")
print(f"Raum- und Zeitplan gespeichert unter: {raum_zeitplan_path}")

from collections import defaultdict
import numpy as np
import pandas as pd

# Dateipfade
schueler_wahlen_path = 'C://Users//SE//Desktop//IMPORT_BOT2_Wahl.xlsx'
raumliste_path = 'C://Users//SE//Desktop//IMPORT_BOT0_Raumliste.xlsx'
veranstaltungsliste_path = 'C://Users//SE//Desktop//IMPORT_BOT1_Veranstaltungsliste.xlsx'

# Einlesen der Dateien in Dataframes
schuelerwahlen_df = pd.read_excel(schueler_wahlen_path)
raumliste_df = pd.read_excel(raumliste_path)
veranstaltungsliste_df = pd.read_excel(veranstaltungsliste_path)

# Anzeigen der ersten Zeilen jeder Tabelle, um die Struktur zu verstehen
print(schuelerwahlen_df.head(), raumliste_df.head(), veranstaltungsliste_df.head())

# Schritt 1: Auszählung wieviele Events benötigt werden
all_choices = pd.concat([schuelerwahlen_df[col] for col in schuelerwahlen_df.columns if 'Wahl' in col])
choice_counts = all_choices.value_counts().sort_index()
choice_counts_df = pd.DataFrame(choice_counts).reset_index()
choice_counts_df.columns = ['VeranstaltungsNr', 'Anzahl Wünsche']
event_requirements_df = pd.merge(choice_counts_df, veranstaltungsliste_df, left_on="VeranstaltungsNr", right_on="Nr. ", how="left")
event_requirements_df['Mindestanzahl Events'] = (event_requirements_df['Anzahl Wünsche'] / event_requirements_df['Max. Teilnehmer']).apply(np.ceil)
print(event_requirements_df[['VeranstaltungsNr', 'Unternehmen', 'Fachrichtung', 'Anzahl Wünsche', 'Max. Teilnehmer', 'Mindestanzahl Events']])

# Schritt 2: Zuweisung von Veranstaltungen zu Räumen und Zeitslots
# Korrektur: Initialisieren Sie den DataFrame korrekt, um den AttributeError zu vermeiden
veranstaltung_zeitslot_raum = pd.DataFrame(columns=['VeranstaltungsNr', 'Raum', 'Zeitslot'])
available_slots_per_room = {room: [1, 2, 3, 4, 5] for room in raumliste_df['Raum']}
raum_zeitslot_belegt = defaultdict(set)

for _, event in veranstaltungsliste_df.iterrows():
    veranstaltungs_nr = event['Nr. ']
    unternehmen = event['Unternehmen']
    frühester_zeitpunkt = event['Frühester Zeitpunkt Num']
    
    for raum_id in raumliste_df['Raum']:
        for zeitslot in range(frühester_zeitpunkt, 6):  # Angenommen, es gibt 5 Zeitslots pro Tag
            if zeitslot not in raum_zeitslot_belegt[raum_id]:
                raum_zeitslot_belegt[raum_id].add(zeitslot)
                neuer_eintrag = pd.DataFrame([[veranstaltungs_nr, raum_id, zeitslot]], columns=['VeranstaltungsNr', 'Raum', 'Zeitslot'])
                veranstaltung_zeitslot_raum = pd.concat([veranstaltung_zeitslot_raum, neuer_eintrag], ignore_index=True)
                break  # Beende die Schleife, wenn Zuweisung erfolgt
        if any(veranstaltung_zeitslot_raum['VeranstaltungsNr'] == veranstaltungs_nr):
            break  # Beende die Schleife, wenn Raum zugewiesen wurde

print(veranstaltung_zeitslot_raum.head())

# Schritt 3: Schüler den Veranstaltungen zuweisen (Unverändert)
aktuelle_teilnehmerzahl = defaultdict(int)
schueler_zuweisung = defaultdict(list)
for index, schueler in schuelerwahlen_df.iterrows():
    zugewiesen = False
    for wahl in ['Wahl 1', 'Wahl 2', 'Wahl 3']:
        if pd.isna(schueler[wahl]):
            continue
        veranstaltungs_nr = int(schueler[wahl])
        if aktuelle_teilnehmerzahl[veranstaltungs_nr] < veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == veranstaltungs_nr, 'Max. Teilnehmer'].values[0]:
            schueler_zuweisung[schueler['Name']].append(veranstaltungs_nr)
            aktuelle_teilnehmerzahl[veranstaltungs_nr] += 1
            zugewiesen = True
            break
schueler_zuweisung_df = pd.DataFrame([(k, v) for k, v in schueler_zuweisung.items()], columns=['Schüler', 'Zugewiesene VeranstaltungsNrn'])
print(schueler_zuweisung_df.head())

# Schritt 4: Exportieren der Daten in Excel-Dateien

# 1. Anwesenheitslisten je Veranstaltung
anwesenheitslisten_path = 'C://Users//SE//Schulprojekt//DPA//Data//Export//EXPORT_BOT5_Anwesenheitslisten.xlsx'
anwesenheitslisten = defaultdict(list)
for schueler, veranstaltungsnummern in schueler_zuweisung.items():
    for veranstaltungsnummer in veranstaltungsnummern:
        anwesenheitslisten[veranstaltungsnummer].append(schueler)
anwesenheitslisten_df = pd.DataFrame([(veranstaltungsnummer, ", ".join(schueler)) for veranstaltungsnummer, schueler in anwesenheitslisten.items()], columns=['VeranstaltungsNr', 'Teilnehmende Schüler'])
anwesenheitslisten_df.to_excel(anwesenheitslisten_path, index=False)

# 2. Raum- und Zeitplan
raum_zeitplan_path = 'C://Users//SE//Schulprojekt//DPA//Data//Export//EXPORT_BOT3_Raum_und_Zeitplan.xlsx'
raum_zeitplan_df = veranstaltung_zeitslot_raum.merge(veranstaltungsliste_df, left_on='VeranstaltungsNr', right_on='Nr. ', how='left')[['VeranstaltungsNr', 'Raum', 'Zeitslot', 'Unternehmen', 'Fachrichtung']]
raum_zeitplan_df.to_excel(raum_zeitplan_path, index=False)

# 3. Laufzettel für Schüler
laufzettel_path = 'C://Users//SE//Schulprojekt//DPA//Data//Export//EXPORT_BOT6_Laufzettel.xlsx'
laufzettel_list = []
for schueler, veranstaltungsnummern in schueler_zuweisung.items():
    for veranstaltungsnummer in veranstaltungsnummern:
        raum_und_zeit = raum_zeitplan_df[raum_zeitplan_df['VeranstaltungsNr'] == veranstaltungsnummer][['Raum', 'Zeitslot']].values[0]
        laufzettel_list.append([schueler, veranstaltungsnummer, raum_und_zeit[0], raum_und_zeit[1]])
laufzettel_df = pd.DataFrame(laufzettel_list, columns=['Schüler', 'VeranstaltungsNr', 'Raum', 'Zeitslot'])
laufzettel_df.to_excel(laufzettel_path, index=False)


# Ausgabe der Pfade zu den erstellten Excel-Dateien
print(f"Anwesenheitslisten gespeichert unter: {anwesenheitslisten_path}")
print(f"Raum- und Zeitplan gespeichert unter: {raum_zeitplan_path}")
print(f"Laufzettel für Schüler gespeichert unter: {laufzettel_path}")

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

# Schritt 1: Auszählung wieviele Events benötigt werden
all_choices = pd.concat([schuelerwahlen_df[col] for col in schuelerwahlen_df.columns if 'Wahl' in col])
choice_counts = all_choices.value_counts().sort_index()
choice_counts_df = pd.DataFrame(choice_counts).reset_index()
choice_counts_df.columns = ['VeranstaltungsNr', 'Anzahl Wünsche']
event_requirements_df = pd.merge(choice_counts_df, veranstaltungsliste_df, left_on="VeranstaltungsNr", right_on="Nr. ", how="left")
event_requirements_df['Mindestanzahl Events'] = (event_requirements_df['Anzahl Wünsche'] / event_requirements_df['Max. Teilnehmer']).apply(np.ceil)
print(event_requirements_df[['VeranstaltungsNr', 'Unternehmen', 'Fachrichtung', 'Anzahl Wünsche', 'Max. Teilnehmer', 'Mindestanzahl Events']])

# Schritt 2: Zuweisung von Veranstaltungen zu Räumen und Zeitslots
event_assignments = []
available_slots_per_room = {room: [1, 2, 3, 4, 5] for room in raumliste_df['Raum']}
for _, event in veranstaltungsliste_df.iterrows():
    for room, available_slots in available_slots_per_room.items():
        if event['Frühester Zeitpunkt'] in available_slots:
            event_assignments.append((event['Nr. '], room, event['Frühester Zeitpunkt']))
            available_slots.remove(event['Frühester Zeitpunkt'])
            break
event_assignments_df = pd.DataFrame(event_assignments, columns=['VeranstaltungsNr', 'Raum', 'Zeitslot'])
print(event_assignments_df.head())


# Konvertiere 'Frühester Zeitpunkt' in numerische Werte, falls noch nicht geschehen
zeitpunkt_mapping = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5}
veranstaltungsliste_df['Frühester Zeitpunkt Num'] = veranstaltungsliste_df['Frühester Zeitpunkt'].map(zeitpunkt_mapping)

# Sortiere Veranstaltungen nach Unternehmen und Frühestem Zeitpunkt Num
veranstaltungsliste_df_sorted = veranstaltungsliste_df.sort_values(by=['Unternehmen', 'Frühester Zeitpunkt Num'])

# Erstelle eine Mapping-Tabelle für die Zuweisung von Veranstaltungen zu Zeitslots und Räumen
veranstaltung_zeitslot_raum = pd.DataFrame(columns=['VeranstaltungsNr', 'Raum', 'Zeitslot'])

# Halte fest, welche Zeitslots in welchen Räumen bereits belegt sind
raum_zeitslot_belegt = defaultdict(set)

# Iteriere durch sortierte Veranstaltungsliste und weise Zeitslots und Räume zu
for _, event in veranstaltungsliste_df_sorted.iterrows():
    veranstaltungs_nr = event['Nr. ']
    unternehmen = event['Unternehmen']
    frühester_zeitpunkt = event['Frühester Zeitpunkt Num']
    
    for raum_id in raumliste_df['Raum']:
        for zeitslot in range(frühester_zeitpunkt, 6):  # Angenommen, es gibt 5 Zeitslots pro Tag
            # Überprüfe, ob der Zeitslot für diesen Raum bereits belegt ist
            if zeitslot not in raum_zeitslot_belegt[raum_id]:
                # Zuweisung des Zeitslots und Raums zur Veranstaltung
                raum_zeitslot_belegt[raum_id].add(zeitslot)
                veranstaltung_zeitslot_raum = veranstaltung_zeitslot_raum.append({
                    'VeranstaltungsNr': veranstaltungs_nr,
                    'Raum': raum_id,
                    'Zeitslot': zeitslot
                }, ignore_index=True)
                
                # Überprüfe die nächste Veranstaltung des gleichen Unternehmens
                # und versuche, sie im gleichen Raum und im direkt folgenden Zeitslot zu platzieren
                break  # Breche die innere Schleife ab, da ein Zeitslot zugewiesen wurde
        if veranstaltung_zeitslot_raum[veranstaltung_zeitslot_raum['VeranstaltungsNr'] == veranstaltungs_nr].shape[0] > 0:
            break  # Breche die äußere Schleife ab, da der Raum zugewiesen wurde

# Anzeigen der aktualisierten Zuweisungen
print(veranstaltung_zeitslot_raum.head())



# Schritt 3: Schüler den Veranstaltungen zuweisen
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

from collections import defaultdict
import numpy as np

import pandas as pd

# Dateipfade
schueler_wahlen_path = 'C://Users//SE//Schulprojekt//DPA//data//IMPORT BOT2_Schülerwahlen.xlsx'
raumliste_path = 'C://Users//SE//Schulprojekt//DPA//data//IMPORT BOT0_Raumliste.xlsx'
veranstaltungsliste_path = 'C://Users//SE//Schulprojekt//DPA//data//IMPORT BOT1_Veranstaltungsliste.xlsx'

# Einlesen der Dateien
schueler_wahlen_df = pd.read_excel(schueler_wahlen_path)
raumliste_df = pd.read_excel(raumliste_path)
veranstaltungsliste_df = pd.read_excel(veranstaltungsliste_path)

# Anzeigen der ersten Zeilen jeder Tabelle, um die Struktur zu verstehen
(schueler_wahlen_df.head(), raumliste_df.head(), veranstaltungsliste_df.head())



# Veranstaltungen auf Räume aufteilen
# Annahme: Jeder Raum kann jede Veranstaltung an jedem Tag aufnehmen, solange die Raumkapazität ausreicht
veranstaltung_zu_raum = {}
raum_verfuegbarkeit = defaultdict(list)

for index, veranstaltung in veranstaltungsliste_df.iterrows():
    veranstaltungs_nr = veranstaltung['Nr.']
    max_teilnehmer = veranstaltung['Max. Teilnehmer']
    
    # Suche nach einem Raum, der die Kapazität erfüllt
    for _, raum in raumliste_df.iterrows():
        if raum['Kapazität'] >= max_teilnehmer:
            veranstaltung_zu_raum[veranstaltungs_nr] = raum['Raum']
            raum_verfuegbarkeit[raum['Raum']].append(veranstaltungs_nr)
            break

# Beginn der Schülerzuweisung
# Initialisiere die Teilnehmerliste für jede Veranstaltung
teilnehmer_zu_veranstaltung = defaultdict(list)

# Funktion, um Schüler den Veranstaltungen zuzuweisen
def schueler_zu_veranstaltung_zuweisen(schueler_wahlen_df, teilnehmer_zu_veranstaltung):
    # Zufällige Zuweisung für Schüler ohne Angaben
    verfuegbare_veranstaltungen = list(veranstaltung_zu_raum.keys())
    np.random.shuffle(verfuegbare_veranstaltungen)
    
    for _, schueler in schueler_wahlen_df.iterrows():
        zugewiesen = False
        wahlen = [schueler[f'Wahl {i}'] for i in range(1, 7) if not pd.isnull(schueler[f'Wahl {i}'])]
        
        # Versuche, den ersten drei Wünschen zu entsprechen
        for wahl in wahlen[:3]:
            if len(teilnehmer_zu_veranstaltung[wahl]) < veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr.'] == wahl, 'Max. Teilnehmer'].iloc[0]:
                teilnehmer_zu_veranstaltung[wahl].append(schueler['Name'])
                zugewiesen = True
                break
        
        # Wenn kein Wunsch erfüllt werden konnte, zufällige Zuweisung
        if not zugewiesen:
            for veranstaltung in verfuegbare_veranstaltungen:
                if len(teilnehmer_zu_veranstaltung[veranstaltung]) < veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr.'] == veranstaltung, 'Max. Teilnehmer'].iloc[0]:
                    teilnehmer_zu_veranstaltung[veranstaltung].append(schueler['Name'])
                    break
    
    return teilnehmer_zu_veranstaltung

# Schüler den Veranstaltungen zuweisen
teilnehmer_zu_veranstaltung = schueler_zu_veranstaltung_zuweisen(schueler_wahlen_df, teilnehmer_zu_veranstaltung)

# Einige Ergebnisse anzeigen
list(veranstaltung_zu_raum.items())[:5], list(teilnehmer_zu_veranstaltung.items())[:5]

from collections import defaultdict
import numpy as np
import pandas as pd
import os
import random

# Dateipfade
## Importpfade
basisimportpfad = 'C://Users//SE//Schulprojekt//DPA//data//'
schueler_wahlen_path = basisimportpfad + 'IMPORT_BOT2_Wahl.xlsx'
raumliste_path = basisimportpfad + 'IMPORT_BOT0_Raumliste.xlsx'
veranstaltungsliste_path = basisimportpfad + 'IMPORT_BOT1_Veranstaltungsliste.xlsx'

## Exportpfade
export_basispfad = 'C://Users//SE//Desktop//Laufzettel//'
anwesenheitslisten_path = export_basispfad + 'EXPORT_BOT5_Anwesenheitslisten.xlsx'
raum_zeitplan_path = export_basispfad + 'EXPORT_BOT3_Raum_und_Zeitplan.xlsx'

# Daten in Dataframes laden
def load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path):
    schuelerwahlen_df = pd.read_excel(schueler_wahlen_path)
    raumliste_df = pd.read_excel(raumliste_path)
    veranstaltungsliste_df = pd.read_excel(veranstaltungsliste_path)

    # Hinzufügen einer eindeutigen Schüler-ID
    schuelerwahlen_df = schuelerwahlen_df.reset_index().rename(columns={'index': 'SchuelerID'})

    return schuelerwahlen_df, raumliste_df, veranstaltungsliste_df

def prepare_veranstaltungsliste(veranstaltungsliste_df):
    zeitpunkt_mapping = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5}
    veranstaltungsliste_df['Frühester Zeitpunkt Num'] = veranstaltungsliste_df['Frühester Zeitpunkt'].map(zeitpunkt_mapping)
    print(veranstaltungsliste_df.columns)
    return veranstaltungsliste_df

def calculate_min_events(schuelerwahlen_df, veranstaltungsliste_df):
    gewichtung = {1: 6, 2: 5, 3: 4, 4: 3, 5: 2, 6: 1}  # Gewichtung für die Wahlen 1-6
    event_demand = defaultdict(int)

    # Durchlaufe jede Wahl und summiere die gewichtete Nachfrage für jede Veranstaltung
    for wahl_num in range(1, 7):
        for _, row in schuelerwahlen_df.iterrows():
            event_id = row[f'Wahl {wahl_num}']
            if pd.notna(event_id):
                event_demand[event_id] += gewichtung[wahl_num]

    # Berechne die Gesamtnachfrage für jedes Unternehmen
    unternehmen_demand = defaultdict(int)
    for event_id, demand in event_demand.items():
        unternehmen = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == event_id, 'Unternehmen'].values[0]
        unternehmen_demand[unternehmen] += demand

    # Berechne die Mindestanzahl von Veranstaltungen pro Unternehmen
    min_events_per_unternehmen = {}
    for unternehmen, demand in unternehmen_demand.items():
        max_teilnehmer = veranstaltungsliste_df.loc[veranstaltungsliste_df['Unternehmen'] == unternehmen, 'Max. Teilnehmer'].max()
        min_events = np.ceil(demand / max_teilnehmer / 6)  # Teile durch 6, um die Gewichtung zu berücksichtigen
        min_events_per_unternehmen[unternehmen] = max(1, min_events)  # Stelle sicher, dass mindestens eine Veranstaltung angeboten wird

    # Zuweisung der Mindestanzahl von Veranstaltungen pro Veranstaltung basierend auf dem Unternehmen
    veranstaltungsliste_df['Mindestanzahl Events'] = veranstaltungsliste_df.apply(lambda row: min_events_per_unternehmen[row['Unternehmen']], axis=1)

    return veranstaltungsliste_df


def assign_rooms_for_entire_day(veranstaltungsliste_df, raumliste_df):
    raum_zuweisung = []
    # Sortiere die Veranstaltungen nach Unternehmen und Mindestanzahl von Events
    sorted_veranstaltungen = veranstaltungsliste_df.sort_values(['Unternehmen', 'Fachrichtung', 'Mindestanzahl Events'], ascending=[True, True, False])
    
    # Halte die Zuweisung von Räumen zu Unternehmen fest
    raum_zu_unternehmen = {}
    unternehmen_zu_raum = {}
    unternehmen_veranstaltungs_counter = defaultdict(int)

    for _, veranstaltung in sorted_veranstaltungen.iterrows():
        unternehmen = veranstaltung['Unternehmen']
        fachrichtung = veranstaltung['Fachrichtung']
        veranstaltungsnummer = veranstaltung['Nr. ']
        mindestanzahl_events = int(veranstaltung['Mindestanzahl Events'])

        if unternehmen in unternehmen_zu_raum:
            # Unternehmen wurde bereits einem Raum zugewiesen
            raum = unternehmen_zu_raum[unternehmen]
            for zeitslot in range(1, mindestanzahl_events + 1):  # Bis zur Mindestanzahl von Events
                if unternehmen_veranstaltungs_counter[unternehmen] < 5:  # Maximal 5 Veranstaltungen pro Tag
                    raum_zuweisung.append({
                        'Unternehmen': unternehmen,
                        'Fachrichtung': fachrichtung,
                        'Veranstaltung': veranstaltungsnummer,
                        'Raum': raum,
                        'Zeitslot': zeitslot
                    })
                    unternehmen_veranstaltungs_counter[unternehmen] += 1
        else:
            # Finde einen verfügbaren Raum, der noch keinem Unternehmen zugewiesen wurde
            for _, raum_info in raumliste_df.iterrows():
                raum = raum_info['Raum']
                if raum not in raum_zu_unternehmen:
                    # Weise den Raum dem Unternehmen für den ganzen Tag zu
                    raum_zu_unternehmen[raum] = unternehmen
                    unternehmen_zu_raum[unternehmen] = raum
                    for zeitslot in range(1, mindestanzahl_events + 1):
                        if unternehmen_veranstaltungs_counter[unternehmen] < 5:  # Maximal 5 Veranstaltungen pro Tag
                            raum_zuweisung.append({
                                'Unternehmen': unternehmen,
                                'Fachrichtung': fachrichtung,
                                'Veranstaltung': veranstaltungsnummer,
                                'Raum': raum,
                                'Zeitslot': zeitslot
                            })
                            unternehmen_veranstaltungs_counter[unternehmen] += 1
                    break

    raum_zeitslot_df = pd.DataFrame(raum_zuweisung)
    return raum_zeitslot_df

def assign_students_to_events(schuelerwahlen_df, veranstaltungsliste_df, raum_zeitslot_df):
    schueler_zuweisungen = defaultdict(list)  # SchülerID -> Liste von Veranstaltungs-IDs und Zeitslots
    veranstaltungs_besetzungen = defaultdict(lambda: defaultdict(int))  # Veranstaltungs-ID -> Zeitslot -> Anzahl Schüler
    verfuegbare_zeitslots_pro_veranstaltung = raum_zeitslot_df.groupby('Veranstaltung')['Zeitslot'].apply(set).to_dict()

    def kann_zuweisen(schueler_id, veranstaltung, zeitslot):
        if zeitslot in [slot for _, slot in schueler_zuweisungen[schueler_id]]:
            return False  # Der Schüler hat bereits eine Veranstaltung in diesem Zeitslot
        if any(veranstaltungs_besetzungen[assigned_event][(assigned_slot // 5) * 5 + zeitslot] > 0 for assigned_event, assigned_slot in schueler_zuweisungen[schueler_id]):
            return False  # Der Schüler hat bereits eine Veranstaltung desselben Typs am selben Tag
        max_kapazitaet = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == veranstaltung, 'Max. Teilnehmer'].iloc[0]
        if veranstaltungs_besetzungen[veranstaltung][zeitslot] >= max_kapazitaet:
            return False  # Die maximale Kapazität der Veranstaltung ist erreicht
        return True

    def sort_events_by_demand(schueler_id, wahlen):
        # Filtere ungültige Wahlnummern heraus
        gueltige_wahlen = [wahl_num for wahl_num in wahlen if pd.notna(schueler[f'Wahl {wahl_num}']) and 1 <= wahl_num <= 6]

        # Konvertiere die Werte in Ganzzahlen, um sicherzustellen, dass sie als Schlüssel im gewichtung-Dictionary verwendet werden können
        gueltige_wahlen_int = [int(w) for w in gueltige_wahlen]

        # Sortiere die gültigen Wahlen basierend auf der Nachfrage (gewichtete Anzahl der Schüler, die diese Wahl gewählt haben)
        gewichtung = {1: 6, 2: 5, 3: 4, 4: 3, 5: 2, 6: 1}
        return sorted(gueltige_wahlen_int, key=lambda w: gewichtung[w], reverse=True)

    for _, schueler in schuelerwahlen_df.iterrows():
        schueler_id = schueler['SchuelerID']  # Geändert von 'Name' zu 'SchuelerID'
        wahlen = [wahl_num for wahl_num in range(1, 7) if pd.notna(schueler[f'Wahl {wahl_num}'])]

        if not wahlen:
            continue  # Der Schüler hat keine gültigen Wahlen getroffen

        # Sortiere die Wahlen basierend auf der Nachfrage (gewichtete Anzahl der Schüler, die diese Wahl gewählt haben)
        sorted_wahlen = sort_events_by_demand(schueler_id, wahlen)

        for wahl in sorted_wahlen:
            wahl_id = schueler[f'Wahl {wahl}']
            if wahl_id not in verfuegbare_zeitslots_pro_veranstaltung:
                continue  # Diese Veranstaltung hat keine verfügbaren Zeitslots

            verfuegbare_zeitslots = verfuegbare_zeitslots_pro_veranstaltung[wahl_id]
            for zeitslot in verfuegbare_zeitslots:
                if kann_zuweisen(schueler_id, wahl_id, zeitslot):
                    schueler_zuweisungen[schueler_id].append((wahl_id, zeitslot))
                    veranstaltungs_besetzungen[wahl_id][zeitslot] += 1
                    break

    return schueler_zuweisungen



def complete_student_assignments(schueler_zuweisungen, kurs_besetzungen, veranstaltungsliste_df, raum_zeitslot_df, verfuegbare_slots_pro_veranstaltung):
    schueler_fehlende_slots = {schueler_id: 5 - len(zuweisungen) for schueler_id, zuweisungen in schueler_zuweisungen.items() if len(zuweisungen) < 5}

    for schueler_id, fehlende_slots in schueler_fehlende_slots.items():
        for _ in range(fehlende_slots):
            for _, veranstaltung in raum_zeitslot_df.iterrows():
                veranstaltung_id = veranstaltung['Veranstaltung']
                zeitslot = veranstaltung['Zeitslot']
                if any(zeitslot == slot for _, slot in schueler_zuweisungen[schueler_id]):
                    continue
                max_kapazitaet = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == veranstaltung_id, 'Max. Teilnehmer'].values[0]
                if kurs_besetzungen[veranstaltung_id][zeitslot] < max_kapazitaet:
                    schueler_zuweisungen[schueler_id].append((veranstaltung_id, zeitslot))
                    kurs_besetzungen[veranstaltung_id][zeitslot] += 1
                    break

    return schueler_zuweisungen

def export_results(veranstaltungsliste_df, raumliste_df, schuelerwahlen_df, export_basispfad):
    # Exportieren der Daten in PDF-Dateien.
    pass

# Exportiere Raum- und Zeitplan
def export_anwesenheitslisten(schueler_zuweisungen, schuelerwahlen_df, export_basispfad):
    # Erstelle eine Liste aus den Zuweisungen für den Export
    zuweisungs_liste = []
    for schueler_id, veranstaltungen in schueler_zuweisungen.items():
        for veranstaltung, zeitslot in veranstaltungen:
            zuweisungs_liste.append({'SchuelerID': schueler_id, 'Veranstaltung': veranstaltung, 'Zeitslot': zeitslot})
    zuweisungs_df = pd.DataFrame(zuweisungs_liste)

    # Erstelle eine Pivot-Tabelle für die Anwesenheitslisten
    anwesenheitslisten_df = zuweisungs_df.pivot_table(index='SchuelerID', columns='Zeitslot', values='Veranstaltung', aggfunc='first', fill_value='').infer_objects()

    # Füge die Klasse und den Namen basierend auf der SchuelerID hinzu
    schueler_info = schuelerwahlen_df[['SchuelerID', 'Name', 'Klasse']].drop_duplicates().set_index('SchuelerID')
    anwesenheitslisten_df = anwesenheitslisten_df.merge(schueler_info, on='SchuelerID', how='left').set_index(['Klasse', 'Name'])

    # Exportiere die Anwesenheitslisten pro Klasse
    for klasse in anwesenheitslisten_df.index.get_level_values(0).unique():
        klasse_df = anwesenheitslisten_df.xs(klasse, level='Klasse')
        klasse_df.to_excel(f'{export_basispfad}Anwesenheitsliste_{klasse}.xlsx')

def erstelle_raumplan(veranstaltungsliste_df, raum_zeitslot_df):
    # Erstelle eine Zusammenführung von Veranstaltungs- und Raum-Zeitslot-Daten
    veranstaltungs_raum_df = pd.merge(veranstaltungsliste_df, raum_zeitslot_df, left_on='Nr. ', right_on='Veranstaltung', how='left')

    # Überprüfe, ob die Spalte "Fachrichtung" vorhanden ist
    fachrichtung_spalten = ['Fachrichtung_x', 'Fachrichtung_y']
    unternehmen_spalten = ['Unternehmen_x', 'Unternehmen_y']
    
    for fachrichtung_spalte, unternehmen_spalte in zip(fachrichtung_spalten, unternehmen_spalten):
        if fachrichtung_spalte in veranstaltungs_raum_df.columns:
            # Anpassen der Spaltennamen für den Pivot-Prozess
            # Hier nehmen wir an, dass 'Unternehmen_x' die gewünschte Spalte für Unternehmen ist
            raumplan_df = veranstaltungs_raum_df.pivot_table(index=[unternehmen_spalte, fachrichtung_spalte], columns='Zeitslot', values='Raum', aggfunc='first', fill_value='')
            return raumplan_df

    print("Die Spalte 'Fachrichtung' konnte nicht gefunden werden.")
    return None




schuelerwahlen_df, raumliste_df, veranstaltungsliste_df = load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path)
veranstaltungsliste_df = prepare_veranstaltungsliste(veranstaltungsliste_df)
veranstaltungsliste_df = calculate_min_events(schuelerwahlen_df, veranstaltungsliste_df)
raum_zeitslot_df = assign_rooms_for_entire_day(veranstaltungsliste_df, raumliste_df)
schueler_zuweisungen = assign_students_to_events(schuelerwahlen_df, veranstaltungsliste_df, raum_zeitslot_df)
export_results(veranstaltungsliste_df, raumliste_df, schuelerwahlen_df, raum_zeitplan_path)
export_anwesenheitslisten(schueler_zuweisungen, schuelerwahlen_df, anwesenheitslisten_path)

# Exportiere Raumplan
# Aufruf der Funktion erstelle_raumplan
raumplan_df = erstelle_raumplan(veranstaltungsliste_df, raum_zeitslot_df)
if raumplan_df is not None:
    exportpfad_raumplan = export_basispfad + 'EXPORT_BOT6_Raumplan.xlsx'
    raumplan_df.to_excel(exportpfad_raumplan)

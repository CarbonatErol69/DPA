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
    veranstaltungsliste_df['Mindestanzahl Events'] = veranstaltungsliste_df['Unternehmen'].map(min_events_per_unternehmen)

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

def assign_students_to_empty_slots(schuelerwahlen_df, veranstaltungsliste_df, raum_zeitslot_df, schueler_zuweisungen=None):
    if schueler_zuweisungen is None:
        schueler_zuweisungen = defaultdict(list)
    
    verfuegbare_zeitslots_pro_fachrichtung = raum_zeitslot_df.groupby('Fachrichtung')['Zeitslot'].apply(set).to_dict()

    def kann_zuweisen(schueler_id, fachrichtung, zeitslot):
        if zeitslot in [assignment['Zeitslot'] for assignment in schueler_zuweisungen[schueler_id]]:
            return False  # Der Schüler hat bereits eine Veranstaltung in diesem Zeitslot
        if any(fachrichtung in assignment['Fachrichtung'] for assignment in schueler_zuweisungen[schueler_id]):
            return False  # Der Schüler hat bereits eine Veranstaltung derselben Fachrichtung am selben Tag
        return True

    for _, schueler in schuelerwahlen_df.iterrows():
        schueler_id = schueler['SchuelerID']
        leere_wahlen = [wahl_num for wahl_num in range(1, 7) if pd.isna(schueler[f'Wahl {wahl_num}'])]

        if not leere_wahlen:
            continue  # Der Schüler hat keine leeren Wahlen

        random.shuffle(leere_wahlen)

        for leere_wahl in leere_wahlen:
            leere_fachrichtung = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == leere_wahl, 'Fachrichtung'].iloc[0]
            if leere_fachrichtung not in verfuegbare_zeitslots_pro_fachrichtung:
                continue  # Diese Fachrichtung hat keine verfügbaren Zeitslots

            verfuegbare_zeitslots = verfuegbare_zeitslots_pro_fachrichtung[leere_fachrichtung]
            for zeitslot in verfuegbare_zeitslots:
                if kann_zuweisen(schueler_id, leere_fachrichtung, zeitslot):
                    schueler_zuweisungen[schueler_id].append({'Wahl': leere_wahl, 'Fachrichtung': leere_fachrichtung, 'Zeitslot': zeitslot})
                    break

    return schueler_zuweisungen


def complete_student_assignments(schueler_zuweisungen, veranstaltungsliste_df, raum_zeitslot_df):
    verfuegbare_zeitslots_pro_fachrichtung = raum_zeitslot_df.groupby('Fachrichtung')['Zeitslot'].apply(set).to_dict()
    schueler_fehlende_slots = {schueler_id: 5 - len(zuweisungen) for schueler_id, zuweisungen in schueler_zuweisungen.items() if len(zuweisungen) < 5}

    def kann_zuweisen(schueler_id, fachrichtung, zeitslot):
        if zeitslot in [assignment['Zeitslot'] for assignment in schueler_zuweisungen[schueler_id]]:
            return False  # Der Schüler hat bereits eine Veranstaltung in diesem Zeitslot
        if any(fachrichtung in assignment['Fachrichtung'] for assignment in schueler_zuweisungen[schueler_id]):
            return False  # Der Schüler hat bereits eine Veranstaltung derselben Fachrichtung am selben Tag
        return True

    for schueler_id, fehlende_slots in schueler_fehlende_slots.items():
        for _ in range(fehlende_slots):
            for fachrichtung, verfuegbare_zeitslots in verfuegbare_zeitslots_pro_fachrichtung.items():
                for zeitslot in verfuegbare_zeitslots:
                    if kann_zuweisen(schueler_id, fachrichtung, zeitslot):
                        schueler_zuweisungen[schueler_id].append({'Wahl': None, 'Fachrichtung': fachrichtung, 'Zeitslot': zeitslot})
                        break

    return schueler_zuweisungen



def export_results(veranstaltungsliste_df, raumliste_df, schuelerwahlen_df, export_basispfad):
    # Exportieren der Daten in PDF-Dateien.
    pass

# Exportiere Raum- und Zeitplan
def export_anwesenheitslisten(schueler_zuweisungen, schuelerwahlen_df, export_basispfad):
    print(schueler_zuweisungen)
    print('Exportiere Anwesenheitslisten...')
    print(schuelerwahlen_df)
    zuweisungs_liste = []
    for schueler_id, veranstaltungen in schueler_zuweisungen.items():
        for zuweisung in veranstaltungen:
            if zuweisung['Wahl'] is not None:
                zuweisungs_liste.append({'SchuelerID': schueler_id, 'Veranstaltung': zuweisung['Wahl'], 'Zeitslot': zuweisung['Zeitslot']})
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

schueler_zuweisungen = defaultdict(list)
schueler_zuweisungen = assign_students_to_empty_slots(schuelerwahlen_df, veranstaltungsliste_df, raum_zeitslot_df)


# Iterativ vervollständige die Zuweisungen für Schüler mit leeren Wahlen und lege zusätzliche leere Slots an, bis möglichst alle Slots belegt sind
while True:
    # Kopiere die aktuelle Zuweisung für spätere Überprüfung
    previous_assignment = schueler_zuweisungen.copy()

    # Weise Schüler mit leeren Wahlen leere Slots zu
    schueler_zuweisungen = assign_students_to_empty_slots(schuelerwahlen_df, veranstaltungsliste_df, raum_zeitslot_df, schueler_zuweisungen)

    # Vervollständige die Zuweisungen für Schüler mit unzureichender Anzahl von Veranstaltungen
    schueler_zuweisungen = complete_student_assignments(schueler_zuweisungen, veranstaltungsliste_df, raum_zeitslot_df)

    # Überprüfe, ob sich die Zuweisungen geändert haben
    if previous_assignment == schueler_zuweisungen:
        break  # Wenn keine Änderungen mehr vorgenommen wurden, beenden Sie die Schleife

export_results(veranstaltungsliste_df, raumliste_df, schuelerwahlen_df, raum_zeitplan_path)
export_anwesenheitslisten(schueler_zuweisungen, schuelerwahlen_df, anwesenheitslisten_path)

# Exportiere Raumplan
# Aufruf der Funktion erstelle_raumplan
raumplan_df = erstelle_raumplan(veranstaltungsliste_df, raum_zeitslot_df)
if raumplan_df is not None:
    exportpfad_raumplan = export_basispfad + 'EXPORT_BOT6_Raumplan.xlsx'
    raumplan_df.to_excel(exportpfad_raumplan)
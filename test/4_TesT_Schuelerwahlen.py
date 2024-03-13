from collections import defaultdict
import numpy as np
import pandas as pd
import os
import random

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
    return veranstaltungsliste_df

def calculate_min_events(schuelerwahlen_df, veranstaltungsliste_df):
    gewichtung = {1: 6, 2: 5, 3: 4, 4: 3, 5: 2, 6: 1}  # Gewichtung für die Wahlen 1-6
    event_demand = defaultdict(int)

    for wahl_num in range(1, 7):
        for _, row in schuelerwahlen_df.iterrows():
            event_id = row[f'Wahl {wahl_num}']
            if pd.notna(event_id):
                event_demand[event_id] += gewichtung[wahl_num]

    unternehmen_demand = defaultdict(int)
    for event_id, demand in event_demand.items():
        unternehmen = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == event_id, 'Unternehmen'].values[0]
        unternehmen_demand[unternehmen] += demand

    min_events_per_unternehmen = {}
    for unternehmen, demand in unternehmen_demand.items():
        max_teilnehmer = veranstaltungsliste_df.loc[veranstaltungsliste_df['Unternehmen'] == unternehmen, 'Max. Teilnehmer'].max()
        min_events = np.ceil(demand / max_teilnehmer / 6)  
        min_events_per_unternehmen[unternehmen] = max(1, min_events)  

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
                    # Unternehmen wurde bereits einem Raum zugewiesen, füge den Zeitslot hinzu
                    raum_zuweisung.append({
                        'Unternehmen': unternehmen,
                        'Fachrichtung': fachrichtung,
                        'Veranstaltung': veranstaltungsnummer,
                        'Raum': raum,
                        'Zeitslot': zeitslot + 1
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
                            # Unternehmen wurde bereits einem Raum zugewiesen, füge den Zeitslot hinzu
                            raum_zuweisung.append({
                                'Unternehmen': unternehmen,
                                'Fachrichtung': fachrichtung,
                                'Veranstaltung': veranstaltungsnummer,
                                'Raum': raum,
                                'Zeitslot': zeitslot + 1
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
            return False  
        if any(fachrichtung in assignment['Fachrichtung'] for assignment in schueler_zuweisungen[schueler_id]):
            return False  
        return True

    for _, schueler in schuelerwahlen_df.iterrows():
        schueler_id = schueler['SchuelerID']
        leere_wahlen = [wahl_num for wahl_num in range(1, 7) if pd.isna(schueler[f'Wahl {wahl_num}'])]

        if not leere_wahlen:
            continue  

        random.shuffle(leere_wahlen)

        for leere_wahl in leere_wahlen:
            if leere_wahl not in veranstaltungsliste_df['Nr. '].values:
                continue  # Wenn leere_wahl nicht in veranstaltungsliste_df vorhanden ist, überspringe sie
            
            leere_fachrichtung = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == leere_wahl, 'Fachrichtung'].iloc[0]
            if leere_fachrichtung not in verfuegbare_zeitslots_pro_fachrichtung:
                continue  

            verfuegbare_zeitslots = verfuegbare_zeitslots_pro_fachrichtung[leere_fachrichtung]
            for zeitslot in verfuegbare_zeitslots:
                if kann_zuweisen(schueler_id, leere_fachrichtung, zeitslot):
                    schueler_zuweisungen[schueler_id].append({'Wahl': leere_wahl, 'Fachrichtung': leere_fachrichtung, 'Zeitslot': zeitslot})
                    break

    print(schueler_zuweisungen)
    return schueler_zuweisungen



def complete_student_assignments(schueler_zuweisungen, veranstaltungsliste_df, raum_zeitslot_df):
    verfuegbare_zeitslots_pro_fachrichtung = raum_zeitslot_df.groupby('Fachrichtung')['Zeitslot'].apply(set).to_dict()
    schueler_fehlende_slots = {schueler_id: 5 - len(zuweisungen) for schueler_id, zuweisungen in schueler_zuweisungen.items() if len(zuweisungen) < 5}

    def kann_zuweisen(schueler_id, fachrichtung, zeitslot):
        if zeitslot in [assignment['Zeitslot'] for assignment in schueler_zuweisungen[schueler_id]]:
            return False  
        if any(fachrichtung in assignment['Fachrichtung'] for assignment in schueler_zuweisungen[schueler_id]):
            return False  
        return True

    for schueler_id, fehlende_slots in schueler_fehlende_slots.items():
        for _ in range(fehlende_slots):
            for fachrichtung, verfuegbare_zeitslots in verfuegbare_zeitslots_pro_fachrichtung.items():
                for zeitslot in verfuegbare_zeitslots:
                    if kann_zuweisen(schueler_id, fachrichtung, zeitslot):
                        schueler_zuweisungen[schueler_id].append({'Wahl': None, 'Fachrichtung': fachrichtung, 'Zeitslot': zeitslot})
                        break
    
    print(schueler_zuweisungen)
    return schueler_zuweisungen

def export_results(veranstaltungsliste_df, raumliste_df, schuelerwahlen_df, export_basispfad, raum_zeitslot_df):
    # Exportieren der Daten in PDF-Dateien.
    pass

def export_anwesenheitslisten(schueler_zuweisungen, schuelerwahlen_df, export_basispfad):
    print('Exportiere Anwesenheitslisten...')
    
    zuweisungs_liste = []
    
    for schueler_id, veranstaltungen in schueler_zuweisungen.items():
        for zuweisung in veranstaltungen:
            zuweisungs_liste.append({
                'SchuelerID': schueler_id,
                'Veranstaltung': zuweisung['Wahl'] if zuweisung['Wahl'] is not None else '',  
                'Zeitslot': zuweisung['Zeitslot']
            })
    
    zuweisungs_df = pd.DataFrame(zuweisungs_liste)

    schueler_info = schuelerwahlen_df[['SchuelerID', 'Name', 'Klasse']].drop_duplicates().set_index('SchuelerID')
    anwesenheitslisten_df = zuweisungs_df.pivot_table(index=['SchuelerID', 'Name', 'Klasse'], columns='Zeitslot', values='Veranstaltung', aggfunc='first', fill_value='').infer_objects()
    
    # Sortiere die Spalten entsprechend der Zeitslots
    anwesenheitslisten_df = anwesenheitslisten_df.reindex(sorted(anwesenheitslisten_df.columns), axis=1)
    
    # Exportiere Anwesenheitslisten pro Klasse
    for klasse in anwesenheitslisten_df.index.get_level_values(2).unique():
        klasse_df = anwesenheitslisten_df.xs(klasse, level='Klasse')
        klasse_df.to_excel(f'{export_basispfad}Anwesenheitsliste_{klasse}.xlsx')

    print("Anwesenheitslisten wurden erfolgreich exportiert.")


def formatiere_und_exportiere_raumplan(raum_zeitslot_df, export_pfad):
    # Erstellen der Pivot-Tabelle
    raumplan_pivot = raum_zeitslot_df.pivot_table(
        index=['Unternehmen', 'Fachrichtung'], 
        columns='Zeitslot', 
        values='Raum', 
        aggfunc='first',
        fill_value='')

    # Exportieren des formatierten Raumplans in eine Excel-Datei
    raumplan_pivot.to_excel(export_pfad, index=True)
    print(f"Formatierter Raumplan wurde erfolgreich nach {export_pfad} exportiert.")


def main():

    # Dateipfade
    basisimportpfad = 'C://Users//SE//Schulprojekt//DPA//data//'
    schueler_wahlen_path = basisimportpfad + 'IMPORT_BOT2_Wahl.xlsx'
    raumliste_path = basisimportpfad + 'IMPORT_BOT0_Raumliste.xlsx'
    veranstaltungsliste_path = basisimportpfad + 'IMPORT_BOT1_Veranstaltungsliste.xlsx'
    export_basispfad = 'C://Users//SE//Desktop//Laufzettel//'
    anwesenheitslisten_path = export_basispfad + 'EXPORT_BOT5_Anwesenheitslisten.xlsx'
    raum_zeitplan_path = export_basispfad + 'EXPORT_BOT3_Raum_und_Zeitplan.xlsx'

    schuelerwahlen_df, raumliste_df, veranstaltungsliste_df = load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path)
    veranstaltungsliste_df = prepare_veranstaltungsliste(veranstaltungsliste_df)
    veranstaltungsliste_df = calculate_min_events(schuelerwahlen_df, veranstaltungsliste_df)
    
    schueler_zuweisungen = defaultdict(list)  # Initialisierung hier
    
    max_iterations = 1000  # Maximale Anzahl von Iterationen
    iteration = 0
    
    while True:
        previous_assignment = schueler_zuweisungen.copy()
        
        for unternehmen in veranstaltungsliste_df['Unternehmen'].unique():
            unternehmen_veranstaltungsliste_df = veranstaltungsliste_df[veranstaltungsliste_df['Unternehmen'] == unternehmen]
            
            raum_zeitslot_df = assign_rooms_for_entire_day(unternehmen_veranstaltungsliste_df, raumliste_df)
            
            schueler_zuweisungen = assign_students_to_empty_slots(schuelerwahlen_df, unternehmen_veranstaltungsliste_df, raum_zeitslot_df, schueler_zuweisungen)
            schueler_zuweisungen = complete_student_assignments(schueler_zuweisungen, unternehmen_veranstaltungsliste_df, raum_zeitslot_df)
        
        iteration += 1
        if previous_assignment == schueler_zuweisungen or iteration >= max_iterations:
            break  
    
    # Exportiere die Ergebnisse
    export_results(veranstaltungsliste_df, raumliste_df, schuelerwahlen_df, export_basispfad, raum_zeitslot_df)
    export_anwesenheitslisten(schueler_zuweisungen, schuelerwahlen_df, export_basispfad)
    formatiere_und_exportiere_raumplan(raum_zeitslot_df, export_basispfad + 'Raumplan.xlsx')

    # Ergebnisse ausgeben
    print("Finaler Raum- und Zeitplan:")
    print(raum_zeitslot_df)  # Zeigt den Raum- und Zeitplan an
    print("\nAnwesenheitslisten:")
    print(schueler_zuweisungen)  # Zeigt die Zuweisungen der Schüler zu Veranstaltungen an

if __name__ == "__main__":
    main()

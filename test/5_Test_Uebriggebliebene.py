from collections import defaultdict
import numpy as np
import pandas as pd
import random

def load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path):
    """Lädt die Daten aus den Excel-Dateien."""
    schuelerwahlen_df = pd.read_excel(schueler_wahlen_path)
    raumliste_df = pd.read_excel(raumliste_path)
    veranstaltungsliste_df = pd.read_excel(veranstaltungsliste_path)
    schuelerwahlen_df = schuelerwahlen_df.reset_index().rename(columns={'index': 'SchuelerID'})
    print('Raumliste:', raumliste_df.columns)
    print('Veranstaltungsliste:', veranstaltungsliste_df.columns)
    print('Schülerwahlen:', schuelerwahlen_df.columns)
    return schuelerwahlen_df, raumliste_df, veranstaltungsliste_df

def prepare_veranstaltungsliste(veranstaltungsliste_df):
    """Bereitet die Veranstaltungsliste vor."""
    zeitpunkt_mapping = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5}
    veranstaltungsliste_df['Frühester Zeitpunkt Num'] = veranstaltungsliste_df['Frühester Zeitpunkt'].map(zeitpunkt_mapping)
    return veranstaltungsliste_df

def calculate_event_demand(schuelerwahlen_df, veranstaltungsliste_df):
    """Berechnet die Nachfrage nach Veranstaltungen."""
    event_demand = defaultdict(int)
    for wahl_num in range(1, 7):
        for _, row in schuelerwahlen_df.iterrows():
            event_id = row[f'Wahl {wahl_num}']
            if pd.notna(event_id):
                event_demand[event_id] += 1
    return event_demand

def assign_rooms_to_companies(veranstaltungsliste_df, raumliste_df):
    """Weist Räume Unternehmen und Fachrichtungen zu."""
    raum_zuweisung = {}
    unternehmen_fachrichtungen = defaultdict(set)
    for _, veranstaltung in veranstaltungsliste_df.iterrows():
        unternehmen = veranstaltung['Unternehmen'].strip()
        fachrichtung = veranstaltung['Fachrichtung']
        raum = raumliste_df.loc[raumliste_df['Unternehmen'] == unternehmen, 'Raum'].iloc[0]
        raum_zuweisung[(unternehmen, fachrichtung)] = raum
        unternehmen_fachrichtungen[unternehmen].add(fachrichtung)
    return raum_zuweisung, unternehmen_fachrichtungen

def assign_rooms_to_events_based_on_capacity(raumliste_df, veranstaltungsliste_df):
    """Weist Räume zu Veranstaltungen basierend auf der Kapazität zu."""
    sorted_raumliste_df = raumliste_df.sort_values(by='Kapazität', ascending=False).reset_index(drop=True)
    sorted_veranstaltungsliste_df = veranstaltungsliste_df.sort_values(by='Max. Teilnehmer', ascending=False).reset_index(drop=True)
    raum_zuweisung = {}
    for _, veranstaltung in sorted_veranstaltungsliste_df.iterrows():
        for _, raum in sorted_raumliste_df.iterrows():
            if raum['Kapazität'] >= veranstaltung['Max. Teilnehmer'] and raum['Raum'] not in raum_zuweisung.values():
                raum_zuweisung[veranstaltung['Nr. ']] = raum['Raum']
                break
    return raum_zuweisung

def estimate_event_count(event_demand, unternehmen_fachrichtungen, veranstaltungsliste_df):
    """Schätzt die Anzahl der Veranstaltungen pro Unternehmen und Fachrichtung."""
    estimated_event_count = {}
    for unternehmen, fachrichtungen in unternehmen_fachrichtungen.items():
        total_demand = sum(event_demand[event_id] for event_id in veranstaltungsliste_df[veranstaltungsliste_df['Unternehmen'] == unternehmen]['Nr. '])
        estimated_event_count[unternehmen] = max(1, np.ceil(total_demand / len(fachrichtungen)))
    return estimated_event_count

def assign_events_to_rooms(veranstaltungsliste_df, raum_zuweisung, estimated_event_count):
    """Weist Veranstaltungen zu Räumen zu."""
    raum_event_zuweisung = defaultdict(list)
    for _, veranstaltung in veranstaltungsliste_df.iterrows():
        unternehmen = veranstaltung['Unternehmen'].strip()
        fachrichtung = veranstaltung['Fachrichtung']
        event_id = veranstaltung['Nr. ']
        raum = raum_zuweisung[(unternehmen, fachrichtung)]
        estimated_count = estimated_event_count[unternehmen]
        for _ in range(int(estimated_count)):
            raum_event_zuweisung[(unternehmen, raum)].append(event_id)
    return raum_event_zuweisung

def assign_students_to_events(schuelerwahlen_df, veranstaltungsliste_df, raum_event_zuweisung):
    """Weist Schüler zu Veranstaltungen zu."""
    schueler_zuweisungen = defaultdict(list)
    for _, schueler in schuelerwahlen_df.iterrows():
        schueler_id = schueler['SchuelerID']
        for wahl_num in range(1, 7):
            event_id = schueler[f'Wahl {wahl_num}']
            if pd.notna(event_id):
                unternehmen = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == event_id, 'Unternehmen'].iloc[0].strip()
                fachrichtung = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == event_id, 'Fachrichtung'].iloc[0]
                raum = raum_event_zuweisung.get((unternehmen, fachrichtung), None)
                if raum:
                    schueler_zuweisungen[schueler_id].append((event_id, raum))
    return schueler_zuweisungen

def complete_student_assignments(schueler_zuweisungen, veranstaltungsliste_df, raum_event_zuweisung):
    """Vervollständigt die Zuweisung von Schülern zu Veranstaltungen."""
    verbleibende_wuensche = defaultdict(list)
    for schueler_id, zuweisungen in schueler_zuweisungen.items():
        for event_id, raum in zuweisungen:
            if event_id in raum_event_zuweisung[(unternehmen, fachrichtung)]:
                raum_event_zuweisung[(unternehmen, fachrichtung)].remove(event_id)
            else:
                verbleibende_wuensche[(unternehmen, fachrichtung)].append(event_id)
    for (unternehmen, fachrichtung), events in verbleibende_wuensche.items():
        raum = raum_event_zuweisung[(unternehmen, fachrichtung)][0]
        for event_id in events:
            schueler_zuweisungen[schueler_id].append((event_id, raum))
    return schueler_zuweisungen

def assign_remaining_students(schueler_zuweisungen, raum_event_zuweisung):
    """Weist verbleibende Schüler zufällig zu Veranstaltungen zu."""
    for schueler_id, zuweisungen in schueler_zuweisungen.items():
        for event_id, raum in zuweisungen:
            if event_id in raum_event_zuweisung[(unternehmen, fachrichtung)]:
                raum_event_zuweisung[(unternehmen, fachrichtung)].remove(event_id)
            else:
                raum = random.choice(list(raum_event_zuweisung[(unternehmen, fachrichtung)]))
                zuweisungen.append((event_id, raum))
    return schueler_zuweisungen

def export_results(veranstaltungsliste_df, schuelerwahlen_df, schueler_zuweisungen, raum_zuweisung, export_basispfad):
    # Bereite einen DataFrame vor, der die finale Zuweisung von Schülern zu Veranstaltungen und Räumen enthält
    export_liste = []
    for schueler_id, zuweisungen in schueler_zuweisungen.items():
        for event_id, raum in zuweisungen:
            schueler_info = schuelerwahlen_df.loc[schuelerwahlen_df['SchuelerID'] == schueler_id, ['Klasse', 'Name', 'Vorname']].iloc[0]
            veranstaltung_info = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == event_id, ['Unternehmen', 'Fachrichtung']].iloc[0]
            export_liste.append({
                'SchuelerID': schueler_id,
                'Klasse': schueler_info['Klasse'],
                'Name': f"{schueler_info['Vorname']} {schueler_info['Name']}",
                'VeranstaltungID': event_id,
                'Unternehmen': veranstaltung_info['Unternehmen'],
                'Fachrichtung': veranstaltung_info['Fachrichtung'],
                'Raum': raum
            })
    export_df = pd.DataFrame(export_liste)
    
    # Überprüfe, ob Spalten 'Klasse' und 'Name' existieren, bevor sortiert wird
    if 'Klasse' in export_df.columns and 'Name' in export_df.columns:
        export_df.sort_values(by=['Klasse', 'Name'], inplace=True)
    else:
        print("Warnung: Spalten 'Klasse' oder 'Name' nicht im DataFrame gefunden. Sortierung übersprungen.")
    
    print(export_df)


def main():
    """Hauptfunktion, die alle Prozesse ausführt."""
    basisimportpfad = 'C://Users//Alex-//Desktop//Schulprojekt//DPA//data//'
    schueler_wahlen_path = basisimportpfad + 'IMPORT_BOT2_Wahl.xlsx'
    raumliste_path = basisimportpfad + 'IMPORT_BOT0_Raumliste.xlsx'
    veranstaltungsliste_path = basisimportpfad + 'IMPORT_BOT1_Veranstaltungsliste.xlsx'
    export_basispfad = 'C://Users//Alex-//Desktop//Laufzettel//'

    schuelerwahlen_df, raumliste_df, veranstaltungsliste_df = load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path)
    veranstaltungsliste_df = prepare_veranstaltungsliste(veranstaltungsliste_df)
    event_demand = calculate_event_demand(schuelerwahlen_df, veranstaltungsliste_df)
    raum_zuweisung = assign_rooms_to_events_based_on_capacity(raumliste_df, veranstaltungsliste_df)
    schueler_zuweisungen = assign_students_to_events(schuelerwahlen_df, veranstaltungsliste_df, raum_zuweisung)
    schueler_zuweisungen = assign_remaining_students(schueler_zuweisungen, raum_zuweisung)
    export_results(veranstaltungsliste_df, schuelerwahlen_df, schueler_zuweisungen, raum_zuweisung, export_basispfad)

if __name__ == "__main__":
    main()

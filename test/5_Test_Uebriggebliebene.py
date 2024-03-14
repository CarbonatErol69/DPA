from collections import defaultdict
import numpy as np
import pandas as pd
import random

def load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path):
    """Lädt die Daten aus den Excel-Dateien."""
    print("Lade Daten aus den Excel-Dateien...")
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
    print("Vorbereite Veranstaltungsliste...")
    company_time_slots = defaultdict(list)
    zeitpunkt_mapping = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5}
    veranstaltungsliste_df['Frühester Zeitpunkt Num'] = veranstaltungsliste_df['Frühester Zeitpunkt'].map(zeitpunkt_mapping)
    for company in veranstaltungsliste_df['Unternehmen'].unique():
        earliest_time = veranstaltungsliste_df[veranstaltungsliste_df['Unternehmen'] == company]['Frühester Zeitpunkt Num'].min()
        # Assign 5 available time slots starting from the earliest time
        available_time_slots = list(range(earliest_time, earliest_time + 5))
        company_time_slots[company] = available_time_slots[:5]  # Ensure only 5 slots are assigned
    veranstaltungsliste_df['Verfügbare Zeitfenster'] = veranstaltungsliste_df.apply(lambda row: company_time_slots[row['Unternehmen']], axis=1)
    print(veranstaltungsliste_df)
    return veranstaltungsliste_df


def calculate_event_demand(schuelerwahlen_df, veranstaltungsliste_df):
    """Berechnet die Nachfrage nach Veranstaltungen."""
    print("Berechne Nachfrage nach Veranstaltungen...")
    event_demand = defaultdict(int)
    for wahl_num in range(1, 7):
        for _, row in schuelerwahlen_df.iterrows():
            event_id = row[f'Wahl {wahl_num}']
            if pd.notna(event_id):
                event_demand[event_id] += 1
    print(event_demand)
    return event_demand

def assign_rooms_to_companies(veranstaltungsliste_df, raumliste_df):
    """Weist Räume Unternehmen und Fachrichtungen zu."""
    print("Weise Räume Unternehmen und Fachrichtungen zu...")
    raum_zuweisung = {}
    unternehmen_fachrichtungen = defaultdict(set)
    
    # Sortiere die Veranstaltungen nach Max. Teilnehmer absteigend
    sorted_veranstaltungen = veranstaltungsliste_df.sort_values(by='Max. Teilnehmer', ascending=False)
    
    # Ein Dictionary, um zu überprüfen, ob ein Unternehmen bereits einem Raum zugewiesen wurde
    unternehmen_raum_mapping = defaultdict(list)
    
    for _, veranstaltung in sorted_veranstaltungen.iterrows():
        unternehmen = veranstaltung['Unternehmen'].strip()
        fachrichtung = veranstaltung['Fachrichtung']
        
        # Überprüfe, ob es bereits ein Unternehmen mit demselben Namen gibt
        if unternehmen in unternehmen_raum_mapping:
            # Wenn ja, weise diesem Unternehmen den Raum des ersten Vorkommens zu
            raum = unternehmen_raum_mapping[unternehmen][0]
            raum_zuweisung[(unternehmen, fachrichtung)] = raum
            print(f"Veranstaltung {veranstaltung['Nr. ']} von Unternehmen {unternehmen} zugewiesen zu Raum {raum}")
        else:
            # Wenn nicht, suche nach einem geeigneten Raum
            for _, raum in raumliste_df.iterrows():
                if raum['Kapazität'] >= veranstaltung['Max. Teilnehmer'] and raum['Raum'] not in raum_zuweisung.values():
                    raum_zuweisung[(unternehmen, fachrichtung)] = raum['Raum']
                    unternehmen_raum_mapping[unternehmen].append(raum['Raum'])
                    print(f"Veranstaltung {veranstaltung['Nr. ']} von Unternehmen {unternehmen} zugewiesen zu Raum {raum['Raum']}")
                    break
            else:
                print(f"Warnung: Kein passender Raum für Veranstaltung {veranstaltung['Nr. ']} von Unternehmen {unternehmen} gefunden!")
    
    return raum_zuweisung, unternehmen_fachrichtungen



def assign_rooms_to_slots(raumliste_df, veranstaltungsliste_df, tag):
    """Weist Räume den Zeitfenstern (Slots) basierend auf der Kapazität zu."""
    print(f"Weise Räume den Zeitfenstern (Slots) für Tag {tag} zu...")
    raum_zuweisung = {}
    sorted_raumliste_df = raumliste_df.sort_values(by='Kapazität', ascending=False).reset_index(drop=True)
    sorted_veranstaltungsliste_df = veranstaltungsliste_df.sort_values(by='Max. Teilnehmer', ascending=False).reset_index(drop=True)
    for _, veranstaltung in sorted_veranstaltungsliste_df.iterrows():
        for _, raum in sorted_raumliste_df.iterrows():
            if raum['Kapazität'] >= veranstaltung['Max. Teilnehmer'] and raum['Raum'] not in raum_zuweisung.values():
                raum_zuweisung[veranstaltung['Nr. ']] = raum['Raum']
                print(f"Veranstaltung {veranstaltung['Nr. ']} zugewiesen zu Raum {raum['Raum']}")
                break
    return raum_zuweisung

def assign_events_to_slots(zeitslots_df, veranstaltungsliste_df, raum_zuweisung):
    """Weist Veranstaltungen zu Slots zu."""
    print("Weise Veranstaltungen zu Slots zu...")
    for _, veranstaltung in veranstaltungsliste_df.iterrows():
        slot = veranstaltung['Frühester Zeitpunkt']
        veranstaltungs_nr = veranstaltung['Nr. ']
        raum = raum_zuweisung.get(veranstaltungs_nr, None)
        if raum:
            # Suche nach dem ersten leeren Slot für diesen Zeitpunkt
            for idx, row in zeitslots_df.iterrows():
                if pd.isnull(row[slot]):
                    zeitslots_df.at[idx, slot] = f"{veranstaltungs_nr}: {raum}"
                    print(f"Veranstaltung {veranstaltungs_nr} zugewiesen zu Zeitslot {idx+1}, Raum: {raum}, Zeitpunkt: {slot}")
                    break
    return zeitslots_df


def assign_students_to_events(schuelerwahlen_df, veranstaltungsliste_df, raum_zuweisung):
    """Weist Schüler zu Veranstaltungen zu."""
    print("Weise Schüler zu Veranstaltungen zu...")
    schueler_zuweisungen = defaultdict(list)
    for _, schueler in schuelerwahlen_df.iterrows():
        schueler_id = schueler['SchuelerID']
        for wahl_num in range(1, 7):
            event_id = schueler[f'Wahl {wahl_num}']
            if pd.notna(event_id):
                raum = raum_zuweisung.get(event_id, None)
                if raum:
                    schueler_zuweisungen[schueler_id].append((event_id, raum))
                    print(f"Schüler {schueler_id} zugewiesen zu Veranstaltung {event_id}, Raum: {raum}")
                    break
    return schueler_zuweisungen

def complete_student_assignments(schueler_zuweisungen, raum_zuweisung, schuelerwahlen_df):
    """Vervollständigt die Zuweisung von Schülern zu Veranstaltungen."""
    print("Vervollständige Zuweisung von Schülern zu Veranstaltungen...")
    verbleibende_wuensche = defaultdict(list)
    
    # Iteriere über alle Schüler und deren Zuweisungen
    for schueler_id, zuweisungen in schueler_zuweisungen.items():
        # Entferne alle Wahlen, wenn weniger als 10 Zuweisungen existieren
        if len(zuweisungen) < 10:
            print(f"Warnung: Weniger als 10 Schüler ({len(zuweisungen)}) für Veranstaltung zugewiesen. Entferne Wahlen...")
            for wahl_num in range(1, 7):
                schuelerwahlen_df.loc[schuelerwahlen_df['SchuelerID'] == schueler_id, f'Wahl {wahl_num}'] = np.nan
                # Hier könntest du verbleibende Wünsche oder eine andere Logik hinzufügen
        else:
            print(f"Schüler {schueler_id} hat alle Wünsche zugewiesen.")

    return schueler_zuweisungen, verbleibende_wuensche


def redistribute_students(schueler_zuweisungen, verbleibende_wuensche, raum_zuweisung, veranstaltungsliste_df):
    """Verteilt Schüler auf noch freie Räume."""
    print("Verteile Schüler auf noch freie Räume...")
    for schueler_id, verbleibende_wahl in verbleibende_wuensche.items():
        for event_id in verbleibende_wahl:
            raum = raum_zuweisung.get(event_id, None)
            if raum:
                schueler_zuweisungen[schueler_id].append((event_id, raum))
                print(f"Schüler {schueler_id} erneut zugewiesen zu Veranstaltung {event_id}, Raum: {raum}")
    print(schueler_zuweisungen)
    return schueler_zuweisungen
        

def generate_time_slots(veranstaltungsliste_df, num_slots):
    """Generiert Zeitslots entsprechend den Zeitfenstern in der Veranstaltungsliste."""
    print("Generiere Zeitslots...")
    unique_slots = veranstaltungsliste_df['Frühester Zeitpunkt'].unique()
    # Erstelle eine leere Liste für die Zeilen
    rows = []
    for _ in range(num_slots):
        # Für jede Zeile, erstelle ein Dictionary, das jeden Slot auf None setzt
        row = {slot: None for slot in unique_slots}
        rows.append(row)
    # Erstelle den DataFrame aus der Liste der Dictionaries
    slots_df = pd.DataFrame(rows)
    return slots_df

def remove_insufficient_slots(zeitslots_df):
    """Entfernt Zeitslots mit weniger als 10 Belegungen."""
    print("Überprüfe Zeitslots auf ausreichende Belegung...")
    print(zeitslots_df)
    for slot in zeitslots_df.columns:
        # Zähle, wie viele Veranstaltungen in jedem Slot zugewiesen sind
        count = zeitslots_df[slot].notnull().sum()
        if count < 10:
            print(f"Warnung: Zeitslot {slot} hat weniger als 10 Belegungen ({count}). Überlege, diesen Slot zu entfernen oder anders zu handhaben.")
            # Füge hier die Logik zum Entfernen oder Umgang mit diesem Slot hinzu.



def main():
    """Hauptfunktion des Programms."""

    # Dateipfade
    ## Importpfade
    basisimportpfad = 'C://Users//Alex-//Desktop//Schulprojekt//DPA//data//'
    schueler_wahlen_path = basisimportpfad + 'IMPORT_BOT2_Wahl.xlsx'
    raumliste_path = basisimportpfad + 'IMPORT_BOT0_Raumliste.xlsx'
    veranstaltungsliste_path = basisimportpfad + 'IMPORT_BOT1_Veranstaltungsliste.xlsx'

    tag = 1
    num_slots = 5
    
    schuelerwahlen_df, raumliste_df, veranstaltungsliste_df = load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path)
    veranstaltungsliste_df = prepare_veranstaltungsliste(veranstaltungsliste_df)
    event_demand = calculate_event_demand(schuelerwahlen_df, veranstaltungsliste_df)
    raum_zuweisung, unternehmen_fachrichtungen = assign_rooms_to_companies(veranstaltungsliste_df, raumliste_df)

    raum_zuweisung = assign_rooms_to_slots(raumliste_df, veranstaltungsliste_df, tag)

    
    zeitslots_df = generate_time_slots(veranstaltungsliste_df, num_slots)
    for i in range(num_slots):
        print(f"\nBearbeite Zeitslot {i+1}:")
        zeitslots_df = assign_events_to_slots(zeitslots_df, veranstaltungsliste_df, raum_zuweisung)
        schueler_zuweisungen = assign_students_to_events(schuelerwahlen_df, veranstaltungsliste_df, raum_zuweisung)
        schueler_zuweisungen, verbleibende_wuensche = complete_student_assignments(schueler_zuweisungen, raum_zuweisung, schuelerwahlen_df)
        if verbleibende_wuensche:
            schueler_zuweisungen = redistribute_students(schueler_zuweisungen, verbleibende_wuensche, raum_zuweisung, veranstaltungsliste_df)

    remove_insufficient_slots(zeitslots_df)

if __name__ == "__main__":
    main()

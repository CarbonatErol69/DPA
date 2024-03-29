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
    
    # Bestimme die Anzahl der benötigten Veranstaltungen
    total_event_demand = veranstaltungsliste_df['Max. Teilnehmer'].sum()
    num_events = int(np.ceil(total_event_demand / veranstaltungsliste_df['Max. Teilnehmer'].max()))
    
    for company, group in veranstaltungsliste_df.groupby(['Unternehmen', 'Fachrichtung']):
        earliest_time = group['Frühester Zeitpunkt Num'].min()
        available_time_slots = list(range(earliest_time, earliest_time + num_events))
        company_time_slots[company] = available_time_slots[:num_events]  # Sicherstellen, dass nur so viele Slots wie nötig zugewiesen werden
    
    veranstaltungsliste_df['Verfügbare Zeitfenster'] = veranstaltungsliste_df.apply(lambda row: company_time_slots[(row['Unternehmen'], row['Fachrichtung'])], axis=1)
    print(veranstaltungsliste_df)
    return veranstaltungsliste_df

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


def calculate_event_demand(schuelerwahlen_df, veranstaltungsliste_df):
    """Berechnet die Nachfrage nach Veranstaltungen."""
    print("Berechne Nachfrage nach Veranstaltungen...")
    event_demand = defaultdict(int)
    for wahl_num in range(1, 5):  # Nur die Wahlen von 1 bis 4 berücksichtigen
        for _, row in schuelerwahlen_df.iterrows():
            event_id = row[f'Wahl {wahl_num}']
            if pd.notna(event_id):
                event_demand[event_id] += 1
    max_participants = veranstaltungsliste_df['Max. Teilnehmer'].max()
    num_events_required = sum(event_demand.values()) / max_participants
    num_events_required = int(np.ceil(num_events_required))  # Aufgerundete Anzahl benötigter Veranstaltungen
    print(f"Anzahl der benötigten Veranstaltungen: {num_events_required}")
    return num_events_required

def generate_room_assignment_matrix(raum_zuweisung, num_slots):
    """Generiert eine Matrix für die Raumzuweisung."""
    print("Generiere Raumzuweisungsmatrix...")
    unique_slots = range(1, num_slots + 1)
    unique_companies = sorted(set(company for company, _ in raum_zuweisung.keys()))
    matrix_data = []
    for company in unique_companies:
        company_row = []
        for slot in unique_slots:
            room = raum_zuweisung.get((company, slot), '')
            company_row.append(room)
        matrix_data.append(company_row)
    room_assignment_matrix = pd.DataFrame(matrix_data, index=unique_companies, columns=unique_slots)
    print(room_assignment_matrix)
    return room_assignment_matrix

def assign_students_to_events(schuelerwahlen_df, veranstaltungsliste_df, raum_zuweisung):
    """Weist Schüler zu Veranstaltungen zu."""
    print("Weise Schüler zu Veranstaltungen zu...")
    schueler_zuweisungen = defaultdict(list)
    
    # Erstelle eine Zuordnung von Veranstaltungen zu verbleibenden Plätzen pro Zeitfenster
    remaining_slots = defaultdict(int)
    for _, row in veranstaltungsliste_df.iterrows():
        for slot in row['Verfügbare Zeitfenster']:
            remaining_slots[(row['Nr. '], slot)] = row['Max. Teilnehmer']
    
    # Iteriere durch jeden Schüler
    for _, schueler in schuelerwahlen_df.iterrows():
        schueler_id = schueler['SchuelerID']
        
        # Iteriere über die Schülerwünsche von 1 bis 5
        for wahl_num in range(1, 6):
            event_id = schueler[f'Wahl {wahl_num}']
            
            # Wenn der Schüler einen Wunsch für diese Wahl hat
            if pd.notna(event_id):
                raum = raum_zuweisung.get(event_id, None)
                
                # Wenn ein Raum für die Veranstaltung gefunden wurde und noch Platz vorhanden ist
                if raum and remaining_slots.get((event_id, raum), 0) > 0:
                    # Reduziere die Anzahl der verbleibenden Plätze für diese Veranstaltung und dieses Zeitfenster
                    remaining_slots[(event_id, raum)] -= 1
                    # Füge die Zuweisung zum Schüler hinzu
                    schueler_zuweisungen[schueler_id].append((event_id, raum))
                    print(f"Schüler {schueler_id} zugewiesen zu Veranstaltung {event_id}, Raum: {raum}")
                    break  # Breche ab, da der Schüler erfolgreich zugewiesen wurde
    print(schueler_zuweisungen)
    return schueler_zuweisungen




def assign_remaining_students(schuelerwahlen_df, schueler_zuweisungen, room_assignment_matrix):
    """Weist verbleibende Schüler zufälligen freien Slots zu."""
    print("Weise verbleibende Schüler zufälligen freien Slots zu...")

    # Iteriere durch jeden Schüler
    for _, schueler in schuelerwahlen_df.iterrows():
        schueler_id = schueler['SchuelerID']
        assigned_slots = [slot for assignment in schueler_zuweisungen[schueler_id] for slot in assignment]

        # Überprüfe, ob der Schüler bereits zugewiesene Slots hat
        if len(assigned_slots) < 5:
            remaining_slots = set(range(1, 6)) - set(assigned_slots)
            
            # Überprüfe, ob es noch freie Slots gibt
            if remaining_slots:
                # Wähle zufällig einen freien Slot aus
                random_slot = random.choice(list(remaining_slots))
                # Überprüfe, ob dieser Slot zur gleichen Zeit wie die anderen zugewiesenen Slots ist
                if all(slot in room_assignment_matrix[schueler_id] for slot in assigned_slots):
                    schueler_zuweisungen[schueler_id].append(random_slot)
                    print(f"Schüler {schueler_id} zufällig zu Zeitfenster {random_slot} zugewiesen.")
                else:
                    print(f"Warnung: Schüler {schueler_id} kann nicht zu einem freien Slot zugewiesen werden, da die Zeitslots nicht übereinstimmen.")
            else:
                print(f"Warnung: Schüler {schueler_id} hat bereits alle verfügbaren Slots zugewiesen bekommen.")
        else:
            print(f"Schüler {schueler_id} hat bereits alle 5 Slots zugewiesen bekommen.")
    
    return schueler_zuweisungen


def export_student_assignments(schueler_zuweisungen, output_path):
    """Exportiert die Schülerzuweisungen in eine Excel-Datei."""
    print("Exportiere Schülerzuweisungen...")
    
    # Erstelle eine leere Liste für die Zeilen des DataFrames
    rows = []
    
    # Fülle die Liste mit den Zuweisungen für jeden Schüler
    for schueler_id, zuweisungen in schueler_zuweisungen.items():
        if zuweisungen:  # Überprüfe, ob der Schüler Zuweisungen hat
            for zuweisung in zuweisungen:
                if isinstance(zuweisung, tuple) and len(zuweisung) == 3:
                    unternehmen, fachrichtung, zeitslot = zuweisung  # Zugriff auf Unternehmen, Fachrichtung und Zeitslot
                    rows.append({
                        "SchuelerID": schueler_id,
                        "Zeitslot": zeitslot,  # Zugriff auf den Zeitslot
                        "Unternehmen": unternehmen,
                        "Fachrichtung": fachrichtung
                    })
                else:
                    print(f"Warnung: Ungültige Zuweisung für Schüler {schueler_id}: {zuweisung}")
    
    # Erstelle ein DataFrame aus der Liste von Zeilen
    export_df = pd.DataFrame(rows)
    
    # Sortiere das DataFrame nach SchülerID und Zeitslot
    export_df = export_df.sort_values(by=["SchuelerID", "Zeitslot"])
    
    # Exportiere das DataFrame in eine Excel-Datei
    export_df.to_excel(output_path, index=False)
    
    print(f"Die Schülerzuweisungen wurden erfolgreich in '{output_path}' exportiert.")







def main():
    """Hauptfunktion des Programms."""

    # Dateipfade
    ## Importpfade
    basisimportpfad = 'C://Users//T1485841//Documents//Projekte//Schulprojekt//DPA//Imports//'
    schueler_wahlen_path = basisimportpfad + 'IMPORT_BOT2_Wahl.xlsx'
    raumliste_path = basisimportpfad + 'IMPORT_BOT0_Raumliste.xlsx'
    veranstaltungsliste_path = basisimportpfad + 'IMPORT_BOT1_Veranstaltungsliste.xlsx'

    tag = 1
    num_slots = 5
    
    schuelerwahlen_df, raumliste_df, veranstaltungsliste_df = load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path)
    veranstaltungsliste_df = prepare_veranstaltungsliste(veranstaltungsliste_df)
    event_demand = calculate_event_demand(schuelerwahlen_df, veranstaltungsliste_df)
    raum_zuweisung, unternehmen_fachrichtungen = assign_rooms_to_companies(veranstaltungsliste_df, raumliste_df)

    room_assignment_matrix = generate_room_assignment_matrix(raum_zuweisung, num_slots)

    schueler_zuweisungen = assign_students_to_events(schuelerwahlen_df, veranstaltungsliste_df, raum_zuweisung)
    schueler_zuweisungen = assign_remaining_students(schuelerwahlen_df, schueler_zuweisungen, room_assignment_matrix)
    
    # Finale Schülerzuweisung anzeigen
    print("\nFinale Schülerzuweisung:")
    for schueler, zuweisungen in schueler_zuweisungen.items():
        print(f"Schüler {schueler}: {zuweisungen}")

    # Vollständige Raumliste anzeigen
    print("\nVollständige Raumliste:")
    print(room_assignment_matrix)


    output_path = 'C://Users//Alex-//Desktop//Schueler_Zuweisungen.xlsx'
    export_student_assignments(schueler_zuweisungen, output_path)

if __name__ == "__main__":
    main()

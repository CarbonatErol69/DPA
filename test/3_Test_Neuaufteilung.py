from collections import defaultdict
import numpy as np
import pandas as pd
import os
import random


# Dateipfade
## Importpfade
basisimportpfad = 'C://Users//Alex-//Desktop//Schulprojekt//DPA//data//'
schueler_wahlen_path = basisimportpfad + 'IMPORT_BOT2_Wahl.xlsx'
raumliste_path = basisimportpfad +'IMPORT_BOT0_Raumliste.xlsx'
veranstaltungsliste_path =  basisimportpfad + 'IMPORT_BOT1_Veranstaltungsliste.xlsx'

## Exportpfade
export_basispfad = 'C://Users//Alex-//Desktop//Laufzettel//'
anwesenheitslisten_path = export_basispfad +'EXPORT_BOT5_Anwesenheitslisten.xlsx'
raum_zeitplan_path = export_basispfad + 'EXPORT_BOT3_Raum_und_Zeitplan.xlsx'

# Daten in Dataframes laden
def load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path):
    schuelerwahlen_df = pd.read_excel(schueler_wahlen_path)
    raumliste_df = pd.read_excel(raumliste_path)
    veranstaltungsliste_df = pd.read_excel(veranstaltungsliste_path)
    return schuelerwahlen_df, raumliste_df, veranstaltungsliste_df

def prepare_veranstaltungsliste(veranstaltungsliste_df):
    zeitpunkt_mapping = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5}
    veranstaltungsliste_df['Frühester Zeitpunkt Num'] = veranstaltungsliste_df['Frühester Zeitpunkt'].map(zeitpunkt_mapping)
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

    #print(veranstaltungsliste_df[['Nr. ', 'Unternehmen', 'Fachrichtung', 'Max. Teilnehmer', 'Mindestanzahl Events']])

    return veranstaltungsliste_df


def assign_rooms_and_slots(veranstaltungsliste_df, raumliste_df):
    # Sortiere Veranstaltungen nach Unternehmen und Frühestem Zeitpunkt
    sorted_veranstaltungen = veranstaltungsliste_df.sort_values(by=['Unternehmen', 'Frühester Zeitpunkt Num'])

    # Zuweisung von Räumen und Zeitslots
    raum_zuweisung = []
    for unternehmen, gruppe in sorted_veranstaltungen.groupby('Unternehmen'):
        max_teilnehmer = gruppe['Max. Teilnehmer'].max()
        # Finde einen Raum, dessen Kapazität ausreichend ist
        raum = raumliste_df[raumliste_df['Kapazität'] >= max_teilnehmer].iloc[0] if not raumliste_df[raumliste_df['Kapazität'] >= max_teilnehmer].empty else None
        if raum is not None:
            zeitslot = 1  # Beginne beim ersten Zeitslot
            for _, veranstaltung in gruppe.iterrows():
                for _ in range(int(veranstaltung['Mindestanzahl Events'])):
                    if zeitslot <= 5:  # Es gibt nur 5 Zeitslots
                        raum_zuweisung.append({
                            'Unternehmen': unternehmen,
                            'Veranstaltung': veranstaltung['Nr. '],
                            'Raum': raum['Raum'],
                            'Zeitslot': zeitslot
                        })
                        zeitslot += 1  # Nächster Zeitslot
                    else:
                        break  # Keine weiteren Zeitslots verfügbar

    # Erstelle einen DataFrame aus den Zuweisungen
    raum_zeitslot_df = pd.DataFrame(raum_zuweisung)

    #print(raum_zeitslot_df)
    return raum_zeitslot_df



def assign_students_to_events(schuelerwahlen_df, veranstaltungsliste_df, raum_zeitslot_df):
    schueler_zuweisungen = defaultdict(list)  # Schüler -> Liste von Veranstaltungs-IDs
    veranstaltungs_besetzungen = defaultdict(lambda: defaultdict(int))  # Veranstaltungs-ID -> Zeitslot -> Anzahl Schüler
    verfuegbare_zeitslots_pro_veranstaltung = raum_zeitslot_df.groupby('Veranstaltung')['Zeitslot'].apply(set).to_dict()

    # Hilfsfunktion, um zu prüfen, ob ein Schüler einem bestimmten Zeitslot zugewiesen werden kann
    def kann_zuweisen(schueler, veranstaltung, zeitslot):
        if zeitslot in [slot for _, slot in schueler_zuweisungen[schueler]]:
            return False  # Der Schüler hat bereits eine Veranstaltung in diesem Zeitslot
        max_kapazitaet = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == veranstaltung, 'Max. Teilnehmer'].iloc[0]
        if veranstaltungs_besetzungen[veranstaltung][zeitslot] >= max_kapazitaet:
            return False  # Die maximale Kapazität der Veranstaltung ist erreicht
        return True

    # Priorisiere Schüler basierend auf ihren Wünschen und weise sie den Kursen zu
    for _, schueler in schuelerwahlen_df.iterrows():
        schueler_id = schueler['Name']
        for wahl_num in range(1, 7):  # Durchlaufe alle Wahlen
            wahl = schueler[f'Wahl {wahl_num}']
            if pd.isna(wahl) or wahl not in verfuegbare_zeitslots_pro_veranstaltung:
                continue  # Überspringe ungültige oder nicht vorhandene Wahlen

            verfuegbare_zeitslots = verfuegbare_zeitslots_pro_veranstaltung[wahl]
            for zeitslot in verfuegbare_zeitslots:
                if kann_zuweisen(schueler_id, wahl, zeitslot):
                    schueler_zuweisungen[schueler_id].append((wahl, zeitslot))
                    veranstaltungs_besetzungen[wahl][zeitslot] += 1
                    break  # Beende die Zuweisung für diese Wahl, wenn erfolgreich

            if len(schueler_zuweisungen[schueler_id]) == 5:
                break  # Der Schüler hat alle 5 Zeitslots besetzt

    # Ergänze Schülerzuweisungen für jene ohne vollständige Wahlen
    schueler_zuweisungen = complete_student_assignments(schueler_zuweisungen, veranstaltungs_besetzungen, veranstaltungsliste_df, raum_zeitslot_df, verfuegbare_zeitslots_pro_veranstaltung)

    return schueler_zuweisungen


def complete_student_assignments(schueler_zuweisungen, kurs_besetzungen, veranstaltungsliste_df, raum_zeitslot_df, verfuegbare_slots_pro_veranstaltung):
    # Identifiziere Schüler, die weniger als 5 Veranstaltungen zugewiesen bekommen haben
    schueler_fehlende_slots = {schueler: 5 - len(zuweisungen) for schueler, zuweisungen in schueler_zuweisungen.items() if len(zuweisungen) < 5}

    # Durchlaufe alle Schüler, die zusätzliche Veranstaltungen benötigen
    for schueler, fehlende_slots in schueler_fehlende_slots.items():
        for _ in range(fehlende_slots):
            # Suche nach verfügbaren Veranstaltungen, die der Schüler besuchen kann
            for _, veranstaltung in raum_zeitslot_df.iterrows():
                veranstaltung_id = veranstaltung['Veranstaltung']
                zeitslot = veranstaltung['Zeitslot']
                # Überprüfe, ob der Schüler bereits in diesem Zeitslot eine Veranstaltung hat
                if any(zeitslot == verf_slot for verf_slot in verfuegbare_slots_pro_veranstaltung.get(veranstaltung_id, [])):
                    continue
                # Überprüfe, ob die Veranstaltung ihre maximale Kapazität erreicht hat
                max_kapazitaet = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr.'] == veranstaltung_id, 'Max. Teilnehmer'].values[0]
                if kurs_besetzungen[veranstaltung_id][zeitslot] < max_kapazitaet:
                    # Weise den Schüler der Veranstaltung zu
                    schueler_zuweisungen[schueler].append(veranstaltung_id)
                    kurs_besetzungen[veranstaltung_id][zeitslot] += 1
                    break  # Beende die Suche, sobald eine Zuweisung erfolgt ist

    print(schueler_zuweisungen)

    return schueler_zuweisungen



def export_results(veranstaltungsliste_df, raumliste_df, schuelerwahlen_df, export_path):
    # Exportieren der Daten in PDF-Dateien.
    pass




  
schuelerwahlen_df, raumliste_df, veranstaltungsliste_df = load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path)
veranstaltungsliste_df = prepare_veranstaltungsliste(veranstaltungsliste_df)
veranstaltungsliste_df = calculate_min_events(schuelerwahlen_df, veranstaltungsliste_df)
raum_zeitslot_df = assign_rooms_and_slots(veranstaltungsliste_df, raumliste_df)
schueler_zuweisungen = assign_students_to_events(schuelerwahlen_df, veranstaltungsliste_df, raum_zeitslot_df)

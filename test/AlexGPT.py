from collections import defaultdict
import numpy as np
import pandas as pd
import random

class AlexGPT():
    def __init__(self, schueler_wahlen_path, raumliste_path, veranstaltungsliste_path):
        #self.main()
        self.schueler_wahlen_path = schueler_wahlen_path
        self.raumliste_path = raumliste_path
        self.veranstaltungsliste_path = veranstaltungsliste_path    
    

    def main(self):
        print(self.schueler_wahlen_path)
        self.create_dataframes()
        self.prepare_veranstaltungsliste()
        self.calculate_event_demand()
        self.assign_rooms_to_companies()
        self.assign_rooms_to_events_based_on_capacity()
        self.estimate_event_count()
        self.assign_events_to_rooms()
        self.assign_students_to_events()
        self.complete_student_assignments()
        self.assign_remaining_students()
        self.export_results()

    
    def create_dataframes(self):
        """Lädt die Daten aus den Excel-Dateien."""
        self.schuelerwahlen_df = pd.read_excel(self.schueler_wahlen_path).reset_index().rename(columns={'index': 'SchuelerID'})
        self.raumliste_df = pd.read_excel(self.raumliste_path)
        self.veranstaltungsliste_df =  pd.read_excel(self.veranstaltungsliste_path)
    
        print('Raumliste:', self.raumliste_df.columns)
        print('Veranstaltungsliste:', self.veranstaltungsliste_df.columns)
        print('Schülerwahlen:', self.schuelerwahlen_df.columns)
        

    def prepare_veranstaltungsliste(self):
        """Bereitet die Veranstaltungsliste vor."""
        zeitpunkt_mapping = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5}
        self.veranstaltungsliste_df['Frühester Zeitpunkt Num'] = self.veranstaltungsliste_df['Frühester Zeitpunkt'].map(zeitpunkt_mapping)
        

    def calculate_event_demand(self):
        """Berechnet die Nachfrage nach Veranstaltungen."""
        event_demand = defaultdict(int)
        for wahl_num in range(1, 7):
             for _, row in self.schuelerwahlen_df.iterrows():
                event_id = row[f'Wahl {wahl_num}']
                if pd.notna(event_id):
                    event_demand[event_id] += 1
                
    def assign_rooms_to_companies(self):
        """Weist Räume Unternehmen und Fachrichtungen zu."""
        raum_zuweisung = {}
        unternehmen_fachrichtungen = defaultdict(set)
        for _, veranstaltung in self.veranstaltungsliste_df.iterrows():
            unternehmen = veranstaltung['Unternehmen'].strip()
            fachrichtung = veranstaltung['Fachrichtung']
            raum = self.raumliste_df.loc[self.raumliste_df['Unternehmen'] == unternehmen, 'Raum'].iloc[0]
            raum_zuweisung[(unternehmen, fachrichtung)] = raum
            unternehmen_fachrichtungen[unternehmen].add(fachrichtung)

    def assign_rooms_to_events_based_on_capacity(self):
        """Weist Räume zu Veranstaltungen basierend auf der Kapazität zu."""
        sorted_raumliste_df = self.raumliste_df.sort_values(by='Kapazität', ascending=False).reset_index(drop=True)
        sorted_veranstaltungsliste_df = self.veranstaltungsliste_df.sort_values(by='Max. Teilnehmer', ascending=False).reset_index(drop=True)
        raum_zuweisung = {}
        for _, veranstaltung in sorted_veranstaltungsliste_df.iterrows():
            for _, raum in sorted_raumliste_df.iterrows():
                if raum['Kapazität'] >= veranstaltung['Max. Teilnehmer'] and raum['Raum'] not in raum_zuweisung.values():
                    raum_zuweisung[veranstaltung['Nr. ']] = raum['Raum']
                    break 
    
    def estimate_event_count(self):
        """Schätzt die Anzahl der Veranstaltungen pro Unternehmen und Fachrichtung."""
        estimated_event_count = {}
        for unternehmen, fachrichtungen in self.unternehmen_fachrichtungen.items():
            total_demand = sum(self.event_demand[event_id] for event_id in self.veranstaltungsliste_df[self.veranstaltungsliste_df['Unternehmen'] == unternehmen]['Nr. '])
            estimated_event_count[unternehmen] = max(1, np.ceil(total_demand / len(fachrichtungen)))

    def assign_events_to_rooms(self):
        """Weist Veranstaltungen zu Räumen zu."""
        raum_event_zuweisung = defaultdict(list)
        for _, veranstaltung in self.veranstaltungsliste_df.iterrows():
            unternehmen = veranstaltung['Unternehmen'].strip()
            fachrichtung = veranstaltung['Fachrichtung']
            event_id = veranstaltung['Nr. ']
            raum = self.raum_zuweisung[(unternehmen, fachrichtung)]
            estimated_count = self.estimated_event_count[unternehmen]
            for _ in range(int(estimated_count)):
                raum_event_zuweisung[(unternehmen, raum)].append(event_id)

    def assign_students_to_events(self):
        """Weist Schüler zu Veranstaltungen zu."""
        schueler_zuweisungen = defaultdict(list)
        for _, schueler in self.schuelerwahlen_df.iterrows():
            schueler_id = schueler['SchuelerID']
            for wahl_num in range(1, 7):
                event_id = schueler[f'Wahl {wahl_num}']
                if pd.notna(event_id):
                    unternehmen = self.veranstaltungsliste_df.loc[self.veranstaltungsliste_df['Nr. '] == event_id, 'Unternehmen'].iloc[0].strip()
                    fachrichtung = self.veranstaltungsliste_df.loc[self.veranstaltungsliste_df['Nr. '] == event_id, 'Fachrichtung'].iloc[0]
                    raum = self.raum_event_zuweisung.get((unternehmen, fachrichtung), None)
                    if raum:
                        schueler_zuweisungen[schueler_id].append((event_id, raum))

    def complete_student_assignments(self):
        """Vervollständigt die Zuweisung von Schülern zu Veranstaltungen."""
        verbleibende_wuensche = defaultdict(list)
        for schueler_id, zuweisungen in self.schueler_zuweisungen.items():
            for event_id, raum in zuweisungen:
                if event_id in self.raum_event_zuweisung[(unternehmen, fachrichtung)]:
                    self.raum_event_zuweisung[(unternehmen, fachrichtung)].remove(event_id)
                else:
                    verbleibende_wuensche[(unternehmen, fachrichtung)].append(event_id)
        for (unternehmen, fachrichtung), events in verbleibende_wuensche.items():
            raum = self.raum_event_zuweisung[(unternehmen, fachrichtung)][0]
            for event_id in events:
                self.schueler_zuweisungen[schueler_id].append((event_id, raum))
    
    def assign_remaining_students(self):
        """Weist verbleibende Schüler zufällig zu Veranstaltungen zu."""
        for schueler_id, zuweisungen in self.schueler_zuweisungen.items():
            for event_id, raum in zuweisungen:
                if event_id in self.raum_event_zuweisung[(self.unternehmen, self.fachrichtung)]:
                    self.raum_event_zuweisung[(self.unternehmen, self.fachrichtung)].remove(event_id)
                else:
                    raum = random.choice(list(self.raum_event_zuweisung[(self.unternehmen, self.fachrichtung)]))
                    zuweisungen.append((event_id, raum))

    
    def export_results(self):
        # Bereite einen DataFrame vor, der die finale Zuweisung von Schülern zu Veranstaltungen und Räumen enthält
        export_liste = []
        for schueler_id, zuweisungen in self.schueler_zuweisungen.items():
            for event_id, raum in zuweisungen:
                schueler_info = self.schuelerwahlen_df.loc[self.schuelerwahlen_df['SchuelerID'] == schueler_id, ['Klasse', 'Name', 'Vorname']].iloc[0]
                veranstaltung_info = self.veranstaltungsliste_df.loc[self.veranstaltungsliste_df['Nr. '] == event_id, ['Unternehmen', 'Fachrichtung']].iloc[0]
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


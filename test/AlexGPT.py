from collections import defaultdict
import numpy as np
import pandas as pd
import random
import math

class AlexGPT():
    def __init__(self, schueler_wahlen_path, raumliste_path, veranstaltungsliste_path):
        #self.main()
        self.schueler_wahlen_path = schueler_wahlen_path
        self.raumliste_path = raumliste_path
        self.veranstaltungsliste_path = veranstaltungsliste_path    # Hola
    

    def main(self):
        self.create_dataframes()
        self.count_votes()
        self.create_event_df()
        self.company_to_events()
        self.create_timetable()
        
        '''
        print(self.schueler_wahlen_path)
        self.create_dataframes()
        self.prepare_veranstaltungsliste()
        self.calculate_event_demand()
        #self.assign_rooms_to_companies()
        #self.assign_rooms_to_events_based_on_capacity()
        #self.estimate_event_count()
        self.assign_events_to_rooms()
        #self.assign_students_to_events()
        #self.complete_student_assignments()
        #self.assign_remaining_students()
        #self.export_results()
        '''

    
    def create_dataframes(self):
        """Lädt die Daten aus den Excel-Dateien."""
        self.df_wahl = pd.read_excel(self.schueler_wahlen_path).reset_index().rename(columns={'index': 'SchuelerID'})
        self.df_raum = pd.read_excel(self.raumliste_path)
        self.df_event =  pd.read_excel(self.veranstaltungsliste_path)
    
        print('Raumliste:', self.df_raum.columns)
        print('Veranstaltungsliste:', self.df_event.columns)
        print('Schülerwahlen:', self.df_wahl.columns)

    def count_votes(self):
        self.value_counts = pd.concat([self.df_wahl['Wahl 1'].value_counts(),
                         self.df_wahl['Wahl 2'].value_counts(),
                         self.df_wahl['Wahl 3'].value_counts(),
                         self.df_wahl['Wahl 4'].value_counts(),
                         self.df_wahl['Wahl 5'].value_counts(),
                         self.df_wahl['Wahl 6'].value_counts()], axis=1)

        column_names = ['Wahl1', 'Wahl2', 'Wahl3', 'Wahl4', 'Wahl5', 'Wahl6']
        self.value_counts.columns = column_names

        self.value_counts.insert(0, 'event_number', self.value_counts.index)

    def create_event_df(self):
        self.df_event = self.df_event.rename(columns={'Nr. ': 'event_number'})

        self.merged_df = pd.merge(self.df_event, self.value_counts, on='event_number')

        self.merged_df['Total 1-3'] = self.merged_df[['Wahl1', 'Wahl2', 'Wahl3']].sum(axis=1).astype(int)

        self.merged_df['Total 4-6'] = self.merged_df[['Wahl4', 'Wahl5', 'Wahl6']].sum(axis=1).astype(int)

        self.merged_df['Total'] = self.merged_df['Total 1-3'] + self.merged_df['Total 4-6']

        self.merged_df['Nötige Slots'] = self.merged_df['Total 1-3'] / self.merged_df['Max. Teilnehmer']

        self.merged_df['Optionale Slots'] = self.merged_df['Total 4-6'] / self.merged_df['Max. Teilnehmer']

        threshhold = 5 / self.merged_df['Max. Teilnehmer']

        # Round the 'Nötige Slots' column up or down
        self.merged_df['Rundet Nötige Slots'] = self.merged_df['Nötige Slots'].apply(lambda x: math.ceil(x) if x - math.floor(x) >= 0.25 else math.floor(x))

        # Round the 'Optionale Slots' column up or down
        self.merged_df['Rundet Optionale Slots'] = self.merged_df['Optionale Slots'].apply(lambda x: math.ceil(x) if x - math.floor(x) >= 0.25 else math.floor(x))

    def create_timetable(self):
        self.df_slots = pd.DataFrame(columns=['A', 'B', 'C', 'D', 'E'])
        self.df_timetable = pd.concat([self.df_raum, self.df_slots[['A', 'B', 'C', 'D', 'E']]], axis=1)

        # Sort df_merged by 'Total 1-3' in descending order
        self.df_merged_sorted = self.merged_df.sort_values('Total 1-3', ascending=False)

        # Iterate over the sorted events
        for _, event in self.df_merged_sorted.iterrows():
            event_number = event['event_number']
            rundet_slots = event['Rundet Nötige Slots']
            
            # Find the room with the least number of events already assigned
            min_events_room = self.df_timetable.iloc[:, 1:].apply(lambda row: row.notnull().sum(), axis=1).idxmin()
            
            # Assign the event to the room with the least number of events
            room_row = self.df_timetable.loc[min_events_room]
            available_slots = rundet_slots - room_row[1:].eq(event_number).sum()
            available_columns = room_row[1:][room_row[1:].isnull()].index.tolist()
            for i in range(min(available_slots, len(available_columns))):
                self.df_timetable.loc[min_events_room, available_columns[i]] = event_number
                

        #TODO: check for frühester startzeitpunkt
                
        #TODO: check for same company in ONE room 
                
        #TODO: 

        print(self.df_timetable)

    def create_timetable2(self):
        self.df_slots = pd.DataFrame(columns=['A', 'B', 'C', 'D', 'E'])
        self.df_timetable = pd.concat([self.df_raum, self.df_slots[['A', 'B', 'C', 'D', 'E']]], axis=1)

        # Sort df_merged by 'Total 1-3' in descending order
        self.df_merged_sorted = self.merged_df.sort_values('Total 1-3', ascending=False)

        # Iterate over the sorted events
        for i in self.company_event_list:
            
            
            # Find the room with the least number of events already assigned
            min_events_room = self.df_timetable.iloc[:, 1:].apply(lambda row: row.notnull().sum(), axis=1).idxmin()
            
            # Assign the event to the room with the least number of events
            room_row = self.df_timetable.loc[min_events_room]
            available_slots = rundet_slots - room_row[1:].eq(event_number).sum()
            available_columns = room_row[1:][room_row[1:].isnull()].index.tolist()
            for i in range(min(available_slots, len(available_columns))):
                self.df_timetable.loc[min_events_room, available_columns[i]] = event_number




                
#pseudo: if n_assigned_events <= nötige_slots:

    def company_to_events(self):
        self.single_event_list = []
        self.multiple_event_list = []

        for e in self.df_merged_sorted.index:
            company = self.df_merged_sorted['Unternehmen'][e]
            event_number = self.df_merged_sorted['event_number'][e]
            
            # Überprüfen, ob das Unternehmen bereits in der Liste ist
            for i, item in enumerate(self.single_event_list):
                if item[0] == company:
                    # Wenn das Unternehmen bereits vorhanden ist, füge die Anzahl der Ereignisse hinzu
                    multiple_events_tuple = self.single_event_list[i] + (event_number,)
                    self.multiple_event_list.append(multiple_events_tuple)
                    break
            else:
                # Wenn das Unternehmen nicht vorhanden ist, füge ein neues Tupel hinzu
                self.single_event_list.append((company, event_number))
        
        self.company_event_list = self.multiple_event_list

        for i in self.single_event_list:
            if not any(i[0] == x[0] for x in self.multiple_event_list):
                self.company_event_list.append(i)

    

        


































        '''
        

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
                
    def assign_rooms_to_events(self):
        pass
    
    def assign_rooms_to_companies(self):
        """Weist Räume Unternehmen und Fachrichtungen zu."""
        raum_zuweisung = {} #was dict before
        unternehmen_fachrichtungen = defaultdict(set)
        for _, veranstaltung in self.veranstaltungsliste_df.iterrows():
            print("verstanstaltung unternehmen: ", veranstaltung['Unternehmen'])
            unternehmen = veranstaltung['Unternehmen'].strip()
            fachrichtung = veranstaltung['Fachrichtung']
            print("raumliste unternehmen:", self.raumliste_df.columns)
            raum = self.raumliste_df.loc[self.veranstaltungsliste_df['Unternehmen'] == unternehmen, 'Raum'].iloc[0]
            print("raum:", raum)
            raum_zuweisung[(unternehmen, fachrichtung)].append(raum)
            print("raum_zuweisung", raum_zuweisung)
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
        print("raum-event zuweisung", raum_event_zuweisung)

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

'''
from collections import defaultdict
import numpy as np
import pandas as pd
import os
from collections import defaultdict
import random

class Raumzuweisung():
    def __init__(self, schueler_path, raum_path, event_path):
        self.directory, filename = os.path.split(schueler_path)
        self.schueler_path = schueler_path
        self.raum_path = raum_path
        self.event_path = event_path
        self.main()

    def load_data(self):
        """Lädt die Daten aus den Excel-Dateien."""
        print("Lade Daten aus den Excel-Dateien...")
        self.schuelerwahlen_df = pd.read_excel(self.schueler_path)
        self.raumliste_df = pd.read_excel(self.raum_path)
        self.veranstaltungsliste_df = pd.read_excel(self.event_path)
        self.schuelerwahlen_df = self.schuelerwahlen_df.reset_index().rename(columns={'index': 'SchuelerID'})
        print('Raumliste:', self.raumliste_df.columns)
        print('Veranstaltungsliste:', self.veranstaltungsliste_df.columns)
        print('Schülerwahlen:', self.schuelerwahlen_df.columns)

        return self.schuelerwahlen_df, self.raumliste_df, self.veranstaltungsliste_df

    def verarbeite_schueler_wahlen(self):
        """Verarbeitet die Schülerwahlen und ermittelt die Gesamthäufigkeiten jeder Wahl."""
        wahl_columns = [col for col in self.schuelerwahlen_df.columns if "Wahl" in col]
        wahlen_count = self.schuelerwahlen_df[wahl_columns].apply(pd.Series.value_counts).sum(axis=1).astype(int)
        wahlen_count_df = wahlen_count.reset_index()
        wahlen_count_df.columns = ['Wahl_Nummer', 'Anzahl_Wahlen']
        print(wahlen_count_df)
        return wahlen_count_df.sort_values(by='Anzahl_Wahlen', ascending=False)

    def berechne_mindest_events_anpassung(self):
        self.veranstaltungsliste_df = self.veranstaltungsliste_df.merge(self.wahlen_haeufigkeiten_df, left_on='Nr.', right_on='Wahl_Nummer', how='left')
        self.veranstaltungsliste_df['Benötigte Veranstaltungen'] = self.veranstaltungsliste_df.apply(
            lambda row: np.ceil(row['Anzahl_Wahlen'] / row['Max. Teilnehmer']) if row['Anzahl_Wahlen'] >= 10 else 0, axis=1).astype(int)
        self.veranstaltungsliste_df.drop(['Wahl_Nummer', 'Anzahl_Wahlen'], axis=1, inplace=True)
        print(self.veranstaltungsliste_df)
        return self.veranstaltungsliste_df

    def identifiziere_mehrfache_veranstaltungen(self):
        """Identifiziert Unternehmen mit mehreren Veranstaltungen."""
        unternehmen_zu_veranstaltungen = defaultdict(list)
        for _, row in self.veranstaltungsliste_df.iterrows():
            unternehmen_zu_veranstaltungen[row['Unternehmen']].append(row)
        mehrfache_veranstaltungen = {k: v for k, v in unternehmen_zu_veranstaltungen.items() if len(v) > 1}
        return mehrfache_veranstaltungen

    def verteile_veranstaltungen(self):
        self.raum_zu_veranstaltung = {raum: [] for raum in self.raumliste_df['Raum']}
        unternehmen_mehrfach = self.identifiziere_mehrfache_veranstaltungen()

        for unternehmen, veranstaltungen in unternehmen_mehrfach.items():
            # Berechne die Gesamtnachfrage für die Veranstaltungen des Unternehmens
            gesamte_nachfrage = sum(veranstaltung['Benötigte Veranstaltungen'] for veranstaltung in veranstaltungen)
            gesamtveranstaltungen_pro_unternehmen = min(5, gesamte_nachfrage)
            
            # Sortiere Veranstaltungen nach Nachfrage, damit Veranstaltungen mit höherer Nachfrage bevorzugt werden
            veranstaltungen_sorted = sorted(veranstaltungen, key=lambda x: -x['Benötigte Veranstaltungen'])
            slots_verbleibend = gesamtveranstaltungen_pro_unternehmen
            
            # Anteilige Zuweisung der Slots
            for veranstaltung in veranstaltungen_sorted:
                if slots_verbleibend > 0:
                    # Zuweisung basierend auf der relativen Nachfrage
                    anteilige_slots = np.ceil(veranstaltung['Benötigte Veranstaltungen'] / gesamte_nachfrage * gesamtveranstaltungen_pro_unternehmen)
                    slots_für_veranstaltung = min(slots_verbleibend, anteilige_slots)
                    slots_verbleibend -= slots_für_veranstaltung
                    
                    # Finde einen Raum und weise die Slots zu
                    for raum, slots in self.raum_zu_veranstaltung.items():
                        if len(slots) + slots_für_veranstaltung <= 5:
                            self.raum_zu_veranstaltung[raum].extend([veranstaltung['Nr.']] * int(slots_für_veranstaltung))
                            break
                    else:
                        print(f"Warnung: Nicht genügend Platz für Veranstaltung {veranstaltung['Nr.']} von {unternehmen}.")

        # Bearbeite Veranstaltungen von Unternehmen mit nur einer Veranstaltung
        verarbeitete_veranstaltungen = set(v['Nr.'] for unternehmen in unternehmen_mehrfach for v in unternehmen_mehrfach[unternehmen])
        einzelveranstaltungen = self.veranstaltungsliste_df[~self.veranstaltungsliste_df['Nr.'].isin(verarbeitete_veranstaltungen)]

        for _, veranstaltung in einzelveranstaltungen.iterrows():
            benötigte_veranstaltungen = veranstaltung['Benötigte Veranstaltungen']
            for raum in self.raum_zu_veranstaltung:
                if len(self.raum_zu_veranstaltung[raum]) + benötigte_veranstaltungen <= 5:
                    self.raum_zu_veranstaltung[raum].extend([veranstaltung['Nr.']] * benötigte_veranstaltungen)
                    break
            else:
                print(f"Warnung: Kein geeigneter Raum für Veranstaltung {veranstaltung['Nr.']} gefunden.")

        return self.raum_zu_veranstaltung

    def exportiere_raumzuweisungen(self):
        # Initialisiere eine Liste für die Exportdaten
        export_data = []

        for index, row in self.veranstaltungsliste_df.iterrows():
            veranstaltungsnummer = row['Nr.']
            unternehmen = row['Unternehmen']
            
            # Finde den Raum und die Slots für jede Veranstaltungsnummer
            raum_slots = [None] * 5  # Initialisiere mit leeren Slots
            for raum, veranstaltungen in self.raum_zu_veranstaltung.items():
                for slot_index, veranstaltung in enumerate(veranstaltungen):
                    if veranstaltung == veranstaltungsnummer:
                        raum_slots[slot_index] = raum  # Weise den Raum dem entsprechenden Slot zu
            
            # Füge die Daten zur Liste hinzu
            export_data.append({
                'Veranstaltungsnummer': veranstaltungsnummer, 
                'Unternehmen': unternehmen, 
                'Slot 1': raum_slots[0], 
                'Slot 2': raum_slots[1], 
                'Slot 3': raum_slots[2], 
                'Slot 4': raum_slots[3], 
                'Slot 5': raum_slots[4]
            })
        
        # Erstelle einen neuen DataFrame aus der Liste
        export_df = pd.DataFrame(export_data)

        # Exportiere den DataFrame in eine Excel-Datei
        with pd.ExcelWriter(self.raum_export, engine='xlsxwriter') as writer:
            export_df.to_excel(writer, sheet_name='Raumzuweisungen', index=False)

        print(f"Die Raumzuweisungen wurden erfolgreich nach {self.raum_export} exportiert.")

    def schuelerzuweisung(self):
        self.raumzuweisungen_df = pd.read_excel(self.raum_export)

        # Extrahiere die verfügbaren Zeitslots für jede Veranstaltung aus der Raumzuweisungs-DataFrame
        veranstaltungs_slots = defaultdict(lambda: defaultdict(list))
        for index, row in self.raumzuweisungen_df.iterrows():
            veranstaltungsnummer = row['Veranstaltungsnummer']
            for slot in range(1, 6):  # Für jeden der 5 Slots
                raum = row[f'Slot {slot}']
                if not pd.isna(raum):  # Wenn ein Raum zugewiesen wurde
                    veranstaltungs_slots[veranstaltungsnummer][slot].append(raum)

        # Konvertiere die defaultdicts zu normalen dicts für eine einfachere Handhabung später
        veranstaltungs_slots = {k: dict(v) for k, v in veranstaltungs_slots.items()}

        # Initialisiere Strukturen für die Zuweisungen und belegten Slots der Schüler
        schueler_zuweisungen = defaultdict(lambda: {'zuweisungen': [], 'belegte_slots': set()})
        veranstaltungen_teilnehmer = defaultdict(lambda: defaultdict(list))

        # Gehe durch jede Wahl und versuche, Schüler ihren gewünschten Veranstaltungen zuzuweisen
        for wahl_nr in range(1, 7):  # Von Wahl 1 bis Wahl 6
            for index, schueler in self.schuelerwahlen_df.iterrows():
                gewaehlte_veranstaltung = schueler[f'Wahl {wahl_nr}']
                if pd.isna(gewaehlte_veranstaltung):  # Überspringe, falls keine Wahl getroffen wurde
                    continue

                # Überprüfe die Verfügbarkeit in den Zeitslots der gewählten Veranstaltung
                for slot, raum in veranstaltungs_slots.get(gewaehlte_veranstaltung, {}).items():
                    if slot not in schueler_zuweisungen[index]['belegte_slots']:  # Slot beim Schüler ist frei
                        max_teilnehmer = self.veranstaltungsliste_df.loc[self.veranstaltungsliste_df['Nr.'] == gewaehlte_veranstaltung, 'Max. Teilnehmer'].iloc[0]
                        if len(veranstaltungen_teilnehmer[gewaehlte_veranstaltung][slot]) < max_teilnehmer:  # Kapazität nicht überschritten
                            # Führe die Zuweisung durch
                            schueler_zuweisungen[index]['zuweisungen'].append({'Veranstaltung': gewaehlte_veranstaltung, 'Slot': slot, 'Wunsch': True})
                            schueler_zuweisungen[index]['belegte_slots'].add(slot)
                            veranstaltungen_teilnehmer[gewaehlte_veranstaltung][slot].append(index)
                            break  # Beende die Schleife, sobald eine Zuweisung erfolgt ist

        # Ergänzung: Identifiziere Schüler ohne Wahlen
        schueler_ohne_wahlen = [index for index, row in self.schuelerwahlen_df.iterrows() if row[['Wahl 1', 'Wahl 2', 'Wahl 3', 'Wahl 4', 'Wahl 5', 'Wahl 6']].isna().all()]

        # Ergänzung: Zufällige Zuweisung für Schüler ohne Wahlen
        for schueler_index in schueler_ohne_wahlen:
            zugewiesen = False
            while not zugewiesen:
                zufaellige_veranstaltung = random.choice(list(veranstaltungs_slots.keys()))
                verfügbare_slots = list(veranstaltungs_slots[zufaellige_veranstaltung].keys())
                zufaelliger_slot = random.choice(verfügbare_slots)
                
                max_teilnehmer = self.veranstaltungsliste_df.loc[self.veranstaltungsliste_df['Nr.'] == zufaellige_veranstaltung, 'Max. Teilnehmer'].iloc[0]
                if len(veranstaltungen_teilnehmer[zufaellige_veranstaltung][zufaelliger_slot]) < max_teilnehmer:
                    schueler_zuweisungen[schueler_index]['zuweisungen'].append({
                        'Veranstaltung': zufaellige_veranstaltung,
                        'Slot': zufaelliger_slot,
                        'Wunsch': False
                    })
                    schueler_zuweisungen[schueler_index]['belegte_slots'].add(zufaelliger_slot)
                    veranstaltungen_teilnehmer[zufaellige_veranstaltung][zufaelliger_slot].append(schueler_index)
                    zugewiesen = True

        # Ergänzung: Zufällige Zuweisung für Schüler mit unvollständigen Belegungen
        for schueler_index, zuweisungen_info in schueler_zuweisungen.items():
            # Prüfe, ob der Schüler weniger als 5 belegte Slots hat
            if len(zuweisungen_info['belegte_slots']) < 5:
                # Versuche, dem Schüler zusätzliche zufällige Veranstaltungen zuzuweisen
                while len(zuweisungen_info['belegte_slots']) < 5:
                    zufaellige_veranstaltung = random.choice(list(veranstaltungs_slots.keys()))
                    verfügbare_slots = list(veranstaltungs_slots[zufaellige_veranstaltung].keys())
                    # Entferne Slots, die bereits belegt sind
                    verfügbare_slots = [slot for slot in verfügbare_slots if slot not in zuweisungen_info['belegte_slots']]
                    
                    # Wenn keine verfügbaren Slots übrig sind, breche die Schleife ab und versuche eine andere Veranstaltung
                    if not verfügbare_slots:
                        continue

                    zufaelliger_slot = random.choice(verfügbare_slots)
                    max_teilnehmer = self.veranstaltungsliste_df.loc[self.veranstaltungsliste_df['Nr.'] == zufaellige_veranstaltung, 'Max. Teilnehmer'].iloc[0]
                    if len(veranstaltungen_teilnehmer[zufaellige_veranstaltung][zufaelliger_slot]) < max_teilnehmer:
                        schueler_zuweisungen[schueler_index]['zuweisungen'].append({
                            'Veranstaltung': zufaellige_veranstaltung,
                            'Slot': zufaelliger_slot,
                            'Wunsch': False
                        })
                        schueler_zuweisungen[schueler_index]['belegte_slots'].add(zufaelliger_slot)
                        veranstaltungen_teilnehmer[zufaellige_veranstaltung][zufaelliger_slot].append(schueler_index)

        # Aktualisiere export_data mit allen Zuweisungen, einschließlich der vollständig zufälligen Zuweisungen
        export_data = []
        for schueler_index, zuweisungen_info in schueler_zuweisungen.items():
            schueler_row = self.schuelerwahlen_df.iloc[schueler_index]
            for zuweisung in zuweisungen_info['zuweisungen']:
                veranstaltungsnummer = zuweisung['Veranstaltung']
                slot = zuweisung['Slot']
                wunsch = zuweisung['Wunsch']
                export_data.append({
                    'Klasse': schueler_row['Klasse'],
                    'Name': schueler_row['Name'],
                    'Vorname': schueler_row['Vorname'],
                    'Veranstaltungsnummer': veranstaltungsnummer,
                    'Slot': slot,
                    'Auf Wunsch zugewiesen': 'Ja' if wunsch else 'Zufällig'
                })

        # Erstelle einen DataFrame für den Export
        export_df = pd.DataFrame(export_data)

        # Exportiere den DataFrame in eine Excel-Datei
        export_df.to_excel(self.schueler_export, index=False)

        print(f"Die vollständigen Zuweisungen wurden erfolgreich nach {self.schueler_export} exportiert.")




    def main(self):
        self.raum_export = self.directory + '//Exports//Raumzuweisungen.xlsx'
        self.schueler_export = self.directory + '//Exports//Schueler_Zuweisungen.xlsx'

        self.schuelerwahlen_df, self.raumliste_df, self.veranstaltungsliste_df = self.load_data()
        self.wahlen_haeufigkeiten_df = self.verarbeite_schueler_wahlen()
        self.veranstaltungsliste_df = self.berechne_mindest_events_anpassung()

        self.raum_zu_veranstaltung = self.verteile_veranstaltungen()
        for raum, veranstaltungen in self.raum_zu_veranstaltung.items():
            print(f"Raum {raum} beherbergt Veranstaltungen: {veranstaltungen}")


        
        self.exportiere_raumzuweisungen()
        self.schuelerzuweisung()


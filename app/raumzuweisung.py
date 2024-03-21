from collections import defaultdict
import numpy as np
import pandas as pd
import random
import xlsxwriter

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

def verarbeite_schueler_wahlen(schuelerwahlen_df):
    """Verarbeitet die Schülerwahlen und ermittelt die Gesamthäufigkeiten jeder Wahl."""
    wahl_columns = [col for col in schuelerwahlen_df.columns if "Wahl" in col]
    wahlen_count = schuelerwahlen_df[wahl_columns].apply(pd.Series.value_counts).sum(axis=1).astype(int)
    wahlen_count_df = wahlen_count.reset_index()
    wahlen_count_df.columns = ['Wahl_Nummer', 'Anzahl_Wahlen']
    print(wahlen_count_df)
    return wahlen_count_df.sort_values(by='Anzahl_Wahlen', ascending=False)

def berechne_mindest_events_anpassung(veranstaltungsliste_df, wahlen_haeufigkeiten_df):
    veranstaltungsliste_df = veranstaltungsliste_df.merge(wahlen_haeufigkeiten_df, left_on='Nr. ', right_on='Wahl_Nummer', how='left')
    veranstaltungsliste_df['Benötigte Veranstaltungen'] = veranstaltungsliste_df.apply(
        lambda row: np.ceil(row['Anzahl_Wahlen'] / row['Max. Teilnehmer']) if row['Anzahl_Wahlen'] >= 10 else 0, axis=1).astype(int)
    veranstaltungsliste_df.drop(['Wahl_Nummer', 'Anzahl_Wahlen'], axis=1, inplace=True)
    print(veranstaltungsliste_df)
    return veranstaltungsliste_df

def identifiziere_mehrfache_veranstaltungen(veranstaltungsliste_df):
    """Identifiziert Unternehmen mit mehreren Veranstaltungen."""
    unternehmen_zu_veranstaltungen = defaultdict(list)
    for _, row in veranstaltungsliste_df.iterrows():
        unternehmen_zu_veranstaltungen[row['Unternehmen']].append(row)
    mehrfache_veranstaltungen = {k: v for k, v in unternehmen_zu_veranstaltungen.items() if len(v) > 1}
    return mehrfache_veranstaltungen

def verteile_veranstaltungen(raumliste_df, veranstaltungsliste_df, wahlen_haeufigkeiten_df):
    raum_zu_veranstaltung = {raum: [] for raum in raumliste_df['Raum']}
    unternehmen_mehrfach = identifiziere_mehrfache_veranstaltungen(veranstaltungsliste_df)

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
                for raum, slots in raum_zu_veranstaltung.items():
                    if len(slots) + slots_für_veranstaltung <= 5:
                        raum_zu_veranstaltung[raum].extend([veranstaltung['Nr. ']] * int(slots_für_veranstaltung))
                        break
                else:
                    print(f"Warnung: Nicht genügend Platz für Veranstaltung {veranstaltung['Nr. ']} von {unternehmen}.")

    # Bearbeite Veranstaltungen von Unternehmen mit nur einer Veranstaltung
    verarbeitete_veranstaltungen = set(v['Nr. '] for unternehmen in unternehmen_mehrfach for v in unternehmen_mehrfach[unternehmen])
    einzelveranstaltungen = veranstaltungsliste_df[~veranstaltungsliste_df['Nr. '].isin(verarbeitete_veranstaltungen)]

    for _, veranstaltung in einzelveranstaltungen.iterrows():
        benötigte_veranstaltungen = veranstaltung['Benötigte Veranstaltungen']
        for raum in raum_zu_veranstaltung:
            if len(raum_zu_veranstaltung[raum]) + benötigte_veranstaltungen <= 5:
                raum_zu_veranstaltung[raum].extend([veranstaltung['Nr. ']] * benötigte_veranstaltungen)
                break
        else:
            print(f"Warnung: Kein geeigneter Raum für Veranstaltung {veranstaltung['Nr. ']} gefunden.")

    return raum_zu_veranstaltung

def exportiere_raumzuweisungen(raum_zu_veranstaltung, veranstaltungsliste_df, exportpfad):
    # Initialisiere eine Liste für die Exportdaten
    export_data = []

    for index, row in veranstaltungsliste_df.iterrows():
        veranstaltungsnummer = row['Nr. ']
        unternehmen = row['Unternehmen']
        
        # Finde den Raum und die Slots für jede Veranstaltungsnummer
        raum_slots = [None] * 5  # Initialisiere mit leeren Slots
        for raum, veranstaltungen in raum_zu_veranstaltung.items():
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
    with pd.ExcelWriter(exportpfad, engine='xlsxwriter') as writer:
        export_df.to_excel(writer, sheet_name='Raumzuweisungen', index=False)

    print(f"Die Raumzuweisungen wurden erfolgreich nach {exportpfad} exportiert.")




def main():
    basisimportpfad = 'C://Users//T1485841//Documents//Projekte//Schulprojekt//DPA//Imports//'
    exportpfad = 'C://Users//T1485841//Documents//Projekte//Schulprojekt//DPA//data//Raumzuweisungen.xlsx'

    schueler_wahlen_path = basisimportpfad + 'IMPORT_BOT2_Wahl.xlsx'
    raumliste_path = basisimportpfad + 'IMPORT_BOT0_Raumliste.xlsx'
    veranstaltungsliste_path = basisimportpfad + 'IMPORT_BOT1_Veranstaltungsliste.xlsx'

    schueler_wahlen_path = basisimportpfad + 'IMPORT_BOT2_Wahl.xlsx'
    raumliste_path = basisimportpfad + 'IMPORT_BOT0_Raumliste.xlsx'
    veranstaltungsliste_path = basisimportpfad + 'IMPORT_BOT1_Veranstaltungsliste.xlsx'

    schuelerwahlen_df, raumliste_df, veranstaltungsliste_df = load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path)
    wahlen_haeufigkeiten_df = verarbeite_schueler_wahlen(schuelerwahlen_df)
    veranstaltungsliste_df = berechne_mindest_events_anpassung(veranstaltungsliste_df, wahlen_haeufigkeiten_df)

    raum_zu_veranstaltung = verteile_veranstaltungen(raumliste_df, veranstaltungsliste_df, wahlen_haeufigkeiten_df)
    for raum, veranstaltungen in raum_zu_veranstaltung.items():
        print(f"Raum {raum} beherbergt Veranstaltungen: {veranstaltungen}")


    
    exportiere_raumzuweisungen(raum_zu_veranstaltung, veranstaltungsliste_df, exportpfad)

if __name__ == "__main__":
    main()
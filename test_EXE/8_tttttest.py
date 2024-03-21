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

def verteile_veranstaltungen(raumliste_df, veranstaltungsliste_df, wahlen_haeufigkeiten_df):
    # Unternehmen identifizieren, die mehr als eine Veranstaltung haben
    unternehmen_mit_mehreren_veranstaltungen = identifiziere_mehrfache_veranstaltungen(veranstaltungsliste_df)

    # Sortiere Veranstaltungen nach der Anzahl der benötigten Veranstaltungen, um zuerst die größten zu platzieren
    veranstaltungsliste_df = veranstaltungsliste_df.sort_values(by='Benötigte Veranstaltungen', ascending=False)

    raum_zu_veranstaltung = {raum['Raum']: [] for _, raum in raumliste_df.iterrows()}
    raum_zu_unternehmen = defaultdict(list)  # Zuordnung von Raum zu Unternehmen für den ganzen Tag

    for _, row in veranstaltungsliste_df.iterrows():
        nr = row['Nr. ']
        unternehmen = row['Unternehmen']
        benötigte_veranstaltungen = row['Benötigte Veranstaltungen']

        # Wenn das Unternehmen bereits für den ganzen Tag einen Raum hat, nutze diesen Raum
        if unternehmen in raum_zu_unternehmen:
            raum = raum_zu_unternehmen[unternehmen][-1]  # Letzter Raum für das Unternehmen
            # Überprüfen, ob der Raum genug Platz für die benötigten Veranstaltungen hat
            if len(raum_zu_veranstaltung[raum]) + benötigte_veranstaltungen <= 5:
                raum_zu_veranstaltung[raum].extend([nr] * benötigte_veranstaltungen)
                continue

        # Suche nach einem Raum, der genug Platz hat und Events hintereinander platziert werden können
        for raum in raum_zu_veranstaltung:
            if len(raum_zu_veranstaltung[raum]) + benötigte_veranstaltungen <= 5:
                if benötigte_veranstaltungen >= 3:
                    # Überprüfe, ob hintereinander genug Platz ist
                    if all(len(raum_zu_veranstaltung[raum][i:i+benötigte_veranstaltungen]) == 0 for i in range(5-benötigte_veranstaltungen+1)):
                        raum_zu_veranstaltung[raum].extend([nr] * benötigte_veranstaltungen)
                        raum_zu_unternehmen[unternehmen].extend([raum] * benötigte_veranstaltungen)
                        break
                else:
                    raum_zu_veranstaltung[raum].extend([nr] * benötigte_veranstaltungen)
                    raum_zu_unternehmen[unternehmen].extend([raum] * benötigte_veranstaltungen)
                    break
        else:
            # Kein Raum gefunden, der Platz bietet, füge eine Warnung hinzu
            print(f"Warnung: Nicht genügend Platz für Veranstaltung {nr} ({benötigte_veranstaltungen} benötigt) gefunden.")

    return raum_zu_veranstaltung

def identifiziere_mehrfache_veranstaltungen(veranstaltungsliste_df):
    unternehmen_zu_veranstaltungen = defaultdict(list)
    for _, row in veranstaltungsliste_df.iterrows():
        unternehmen_zu_veranstaltungen[row['Unternehmen']].append(row['Nr. '])
    mehrfache_veranstaltungen = {k: v for k, v in unternehmen_zu_veranstaltungen.items() if len(v) > 1}
    return mehrfache_veranstaltungen



def main():
    basisimportpfad = 'C://Users//T1485841//Documents//Projekte//Schulprojekt//DPA//Imports//'
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

if __name__ == "__main__":
    main()
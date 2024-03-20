import pandas as pd
from collections import defaultdict

# Pfade zu den Dateien
basisimportpfad = 'C://Users//Alex-//Desktop//Schulprojekt//DPA//data//'
exportpfad = 'C://Users//Alex-//Desktop//Schulprojekt//DPA//data//Schueler_Zuweisungen.xlsx'

schueler_wahlen_path = basisimportpfad + 'IMPORT_BOT2_Wahl.xlsx'
raumliste_path = basisimportpfad + 'Raumzuweisungen.xlsx'
veranstaltungsliste_path = basisimportpfad + 'IMPORT_BOT1_Veranstaltungsliste.xlsx'

# Lade die Daten
schuelerwahlen_df = pd.read_excel(schueler_wahlen_path)
raumzuweisungen_df = pd.read_excel(raumliste_path)
veranstaltungsliste_df = pd.read_excel(veranstaltungsliste_path)

# Extrahiere die verfügbaren Zeitslots für jede Veranstaltung aus der Raumzuweisungs-DataFrame
veranstaltungs_slots = defaultdict(lambda: defaultdict(list))
for index, row in raumzuweisungen_df.iterrows():
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
    for index, schueler in schuelerwahlen_df.iterrows():
        gewaehlte_veranstaltung = schueler[f'Wahl {wahl_nr}']
        if pd.isna(gewaehlte_veranstaltung):  # Überspringe, falls keine Wahl getroffen wurde
            continue

        # Überprüfe die Verfügbarkeit in den Zeitslots der gewählten Veranstaltung
        for slot, raum in veranstaltungs_slots.get(gewaehlte_veranstaltung, {}).items():
            if slot not in schueler_zuweisungen[index]['belegte_slots']:  # Slot beim Schüler ist frei
                max_teilnehmer = veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr.'] == gewaehlte_veranstaltung, 'Max. Teilnehmer'].iloc[0]
                if len(veranstaltungen_teilnehmer[gewaehlte_veranstaltung][slot]) < max_teilnehmer:  # Kapazität nicht überschritten
                    # Führe die Zuweisung durch
                    schueler_zuweisungen[index]['zuweisungen'].append({'Veranstaltung': gewaehlte_veranstaltung, 'Slot': slot, 'Wunsch': True})
                    schueler_zuweisungen[index]['belegte_slots'].add(slot)
                    veranstaltungen_teilnehmer[gewaehlte_veranstaltung][slot].append(index)
                    break  # Beende die Schleife, sobald eine Zuweisung erfolgt ist

# Umwandlung der Zuweisungen in ein Format, das für den Export geeignet ist
export_data = []
for schueler_index, zuweisungen_info in schueler_zuweisungen.items():
    schueler_row = schuelerwahlen_df.iloc[schueler_index]
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
            'Auf Wunsch zugewiesen': 'Ja' if wunsch else 'Nein'
        })

# Erstelle einen DataFrame für den Export
export_df = pd.DataFrame(export_data)

# Exportiere den DataFrame in eine Excel-Datei
export_df.to_excel(exportpfad, index=False)

print(f"Die Zuweisungen wurden erfolgreich nach {exportpfad} exportiert.")

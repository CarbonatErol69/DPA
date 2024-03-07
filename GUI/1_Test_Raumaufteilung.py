from collections import defaultdict
import numpy as np
import pandas as pd
import os

# Dateipfade
schueler_wahlen_path = 'C://Users//SE//Desktop//IMPORT_BOT2_Wahl.xlsx'
raumliste_path = 'C://Users//SE//Desktop//IMPORT_BOT0_Raumliste.xlsx'
veranstaltungsliste_path = 'C://Users//SE//Desktop//IMPORT_BOT1_Veranstaltungsliste.xlsx'


# Einlesen der Dateien in Dataframes
schuelerwahlen_df = pd.read_excel(schueler_wahlen_path)
raumliste_df = pd.read_excel(raumliste_path)
veranstaltungsliste_df = pd.read_excel(veranstaltungsliste_path)


# Konvertiere 'Frühester Zeitpunkt' in numerische Werte
zeitpunkt_mapping = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5}
veranstaltungsliste_df['Frühester Zeitpunkt Num'] = veranstaltungsliste_df['Frühester Zeitpunkt'].map(zeitpunkt_mapping)


# Anzeigen der ersten Zeilen jeder Tabelle, um die Struktur zu verstehen
print(schuelerwahlen_df.head(), raumliste_df.head(), veranstaltungsliste_df.head())

# Schritt 1: Auszählung wieviele Events benötigt werden
all_choices = pd.concat([schuelerwahlen_df[col] for col in schuelerwahlen_df.columns if 'Wahl' in col])
choice_counts = all_choices.value_counts().sort_index()
choice_counts_df = pd.DataFrame(choice_counts).reset_index()
choice_counts_df.columns = ['VeranstaltungsNr', 'Anzahl Wünsche']
event_requirements_df = pd.merge(choice_counts_df, veranstaltungsliste_df, left_on="VeranstaltungsNr", right_on="Nr. ", how="left")
event_requirements_df['Mindestanzahl Events'] = (event_requirements_df['Anzahl Wünsche'] / event_requirements_df['Max. Teilnehmer']).apply(np.ceil)
print(event_requirements_df[['VeranstaltungsNr', 'Unternehmen', 'Fachrichtung', 'Anzahl Wünsche', 'Max. Teilnehmer', 'Mindestanzahl Events']])


# Schritt 2: Veranstaltungen den Räumen und Zeitslots zuordnen
# Diese Funktion ordnet die Veranstaltungen zu, unter Berücksichtigung der Unternehmens- und Fachrichtungslogik.
def zuordnung_veranstaltungen_zu_raeumen(veranstaltungsliste_df, raumliste_df):
    # Initialisierung der notwendigen Strukturen
    veranstaltung_zu_raum_zeit = []
    raum_verfuegbarkeit = {raum: list(range(1, 6)) for raum in raumliste_df['Raum']}
    unternehmen_zuletzt_im_raum = {}
    
    # Sortiere Veranstaltungen nach Unternehmen und Fachrichtung
    sortierte_veranstaltungen = veranstaltungsliste_df.sort_values(by=['Unternehmen', 'Fachrichtung', 'Frühester Zeitpunkt Num'])

    # Durchgehe alle Veranstaltungen und ordne sie zu
    for _, veranstaltung in sortierte_veranstaltungen.iterrows():
        veranstaltungs_id = veranstaltung['Nr. ']
        unternehmen = veranstaltung['Unternehmen']
        for raum, zeitslots in raum_verfuegbarkeit.items():
            if unternehmen in unternehmen_zuletzt_im_raum and unternehmen_zuletzt_im_raum[unternehmen] != raum:
                continue  # Überspringe, wenn das Unternehmen bereits einem anderen Raum zugeordnet wurde
            for zeitslot in zeitslots:
                # Zuordnung der Veranstaltung zum ersten verfügbaren Zeitslot
                veranstaltung_zu_raum_zeit.append((veranstaltungs_id, raum, zeitslot))
                zeitslots.remove(zeitslot)  # Entferne den zugewiesenen Zeitslot
                unternehmen_zuletzt_im_raum[unternehmen] = raum  # Aktualisiere den zuletzt zugewiesenen Raum für das Unternehmen
                break
            if unternehmen in unternehmen_zuletzt_im_raum:
                break  # Beende, wenn eine Zuordnung erfolgt ist

    return pd.DataFrame(veranstaltung_zu_raum_zeit, columns=['VeranstaltungsNr', 'Raum', 'Zeitslot'])

# Anwendung der Zuordnungsfunktion
veranstaltung_zeitslot_raum_df = zuordnung_veranstaltungen_zu_raeumen(veranstaltungsliste_df, raumliste_df)



# Schritt 3: Schüler den Veranstaltungen zuweisen, unter Berücksichtigung der Anforderung für 4 Veranstaltungen
aktuelle_teilnehmerzahl = defaultdict(int)
schueler_zuweisung = defaultdict(list)

# Zuweisung basierend auf Schülerwahlen
for index, schueler in schuelerwahlen_df.iterrows():
    zugewiesene_events = 0
    for wahl in ['Wahl 1', 'Wahl 2', 'Wahl 3', 'Wahl 4', 'Wahl 5']:
        if pd.isna(schueler[wahl]) or zugewiesene_events >= 4:  # Änderung zu 4 Veranstaltungen
            continue
        veranstaltungs_nr = int(schueler[wahl])
        if aktuelle_teilnehmerzahl[veranstaltungs_nr] < veranstaltungsliste_df.loc[veranstaltungsliste_df['Nr. '] == veranstaltungs_nr, 'Max. Teilnehmer'].values[0]:
            # Zuweisung der Veranstaltung
            schueler_zuweisung[schueler['Name']].append(veranstaltungs_nr)
            aktuelle_teilnehmerzahl[veranstaltungs_nr] += 1
            zugewiesene_events += 1
            if zugewiesene_events >= 4:  # Stoppe nach Zuweisung von 4 Veranstaltungen
                break

# Konvertierung der Zuweisungen in einen DataFrame
schueler_zuweisung_df = pd.DataFrame([(k, v) for k, v in schueler_zuweisung.items()], columns=['Schüler', 'Zugewiesene VeranstaltungsNrn'])
print(schueler_zuweisung_df.head())







# Schritt 4: Exportieren der Daten in Excel-Dateien

# Basispfad für den Export festlegen
export_basispfad = 'C:/Users/SE/Desktop/Laufzettel/'

# 1. Anwesenheitslisten je Veranstaltung
anwesenheitslisten_path = 'C://Users//SE//Schulprojekt//DPA//Data//Export//EXPORT_BOT5_Anwesenheitslisten.xlsx'
anwesenheitslisten = defaultdict(list)
for schueler, veranstaltungsnummern in schueler_zuweisung.items():
    for veranstaltungsnummer in veranstaltungsnummern:
        anwesenheitslisten[veranstaltungsnummer].append(schueler)
anwesenheitslisten_df = pd.DataFrame([(veranstaltungsnummer, ", ".join(schueler)) for veranstaltungsnummer, schueler in anwesenheitslisten.items()], columns=['VeranstaltungsNr', 'Teilnehmende Schüler'])
anwesenheitslisten_df.to_excel(anwesenheitslisten_path, index=False)

# 2. Raum- und Zeitplan
raum_zeitplan_path = 'C://Users//SE//Schulprojekt//DPA//Data//Export//EXPORT_BOT3_Raum_und_Zeitplan.xlsx'
raum_zeitplan_df = veranstaltung_zeitslot_raum_df.merge(veranstaltungsliste_df, left_on='VeranstaltungsNr', right_on='Nr. ', how='left')[['VeranstaltungsNr', 'Raum', 'Zeitslot', 'Unternehmen', 'Fachrichtung']]
raum_zeitplan_df.to_excel(raum_zeitplan_path, index=False)

# 3. Laufzettel für Schüler
# Durchlaufe alle Schülerzuweisungen
for schueler, veranstaltungs_nrn in schueler_zuweisung.items():
    # Hole die Klasseninformation des Schülers
    klasse = schuelerwahlen_df.loc[schuelerwahlen_df['Name'] == schueler, 'Klasse'].values[0]
    # Erstelle den Pfad basierend auf der Klasse
    klassenpfad = os.path.join(export_basispfad, klasse)
    # Erstelle den Ordner, wenn er nicht existiert
    os.makedirs(klassenpfad, exist_ok=True)
    
    # Hole und sortiere Veranstaltungsdetails wie zuvor
    veranstaltungen_details = veranstaltung_zeitslot_raum_df[veranstaltung_zeitslot_raum_df['VeranstaltungsNr'].isin(veranstaltungs_nrn)]
    veranstaltungen_details = veranstaltungen_details.merge(veranstaltungsliste_df, left_on='VeranstaltungsNr', right_on='Nr. ', how='left').sort_values(by=['Zeitslot'])


    # Erstelle Laufzettel für den aktuellen Schüler
    laufzettel = []
    for _, row in veranstaltungen_details.iterrows():
        laufzettel.append([schueler, row['VeranstaltungsNr'], row['Raum'], row['Zeitslot'], row['Unternehmen'], row['Fachrichtung']])
    
    # Konvertiere in DataFrame
    laufzettel_df = pd.DataFrame(laufzettel, columns=['Schüler', 'VeranstaltungsNr', 'Raum', 'Zeitslot', 'Unternehmen', 'Fachrichtung'])
    
    # Exportiere Laufzettel als Excel-Datei in den klassenspezifischen Ordner
    laufzettel_pfad = os.path.join(klassenpfad, f"Laufzettel_{schueler.replace(' ', '_')}.xlsx")
    laufzettel_df.to_excel(laufzettel_pfad, index=False)

print(f"Laufzettel für Schüler wurden in den klassenspezifischen Ordnern gespeichert.")


# Ausgabe der Pfade zu den erstellten Excel-Dateien
print(f"Anwesenheitslisten gespeichert unter: {anwesenheitslisten_path}")
print(f"Raum- und Zeitplan gespeichert unter: {raum_zeitplan_path}")

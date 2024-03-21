# export_functions.py

import pandas as pd

def export_anwesenheitslisten(schueler_zuweisungen, schuelerwahlen_df, export_path):
    # Konvertiere die Zuweisungen in ein DataFrame für eine einfachere Handhabung
    zuweisungs_liste = []
    for schueler, veranstaltungen in schueler_zuweisungen.items():
        for veranstaltung, zeitslot in veranstaltungen:
            zuweisungs_liste.append({'Name': schueler, 'Veranstaltung': veranstaltung, 'Zeitslot': zeitslot})
    zuweisungs_df = pd.DataFrame(zuweisungs_liste)

    # Bereite das DataFrame für die Anwesenheitslisten vor
    anwesenheitslisten_df = zuweisungs_df.pivot_table(index='Name', columns='Zeitslot', values='Veranstaltung', aggfunc='first', fill_value='')

    # Ergänze die Klasse zu jedem Schüler
    anwesenheitslisten_df = anwesenheitslisten_df.merge(schuelerwahlen_df[['Name', 'Klasse']], on='Name', how='left').set_index(['Klasse', 'Name'])

    # Exportiere pro Klasse eine Anwesenheitsliste
    for klasse in anwesenheitslisten_df.index.get_level_values(0).unique():
        klasse_df = anwesenheitslisten_df.xs(klasse, level='Klasse')
        klasse_df.to_excel(f'{export_path}Anwesenheitsliste_{klasse}.xlsx')

from collections import defaultdict
import numpy as np
import pandas as pd
import random

class AlexGPT():
    def __init__(self) -> None:
        #self.main()
        pass

    

    def main(self):
        print("Carbonat Erol ist ein cooler Typ!  ")        
    

    def load_data(schueler_wahlen_path, raumliste_path, veranstaltungsliste_path):
        """Lädt die Daten aus den Excel-Dateien."""
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
        zeitpunkt_mapping = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5}
        veranstaltungsliste_df['Frühester Zeitpunkt Num'] = veranstaltungsliste_df['Frühester Zeitpunkt'].map(zeitpunkt_mapping)
        return veranstaltungsliste_df

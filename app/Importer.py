import os

class Importer():

  def __init__(self, sql):
    self.sql = sql
    self.file_path = 'C:\\Users\\T1485841\\Documents\\Projekte\\Schulprojekt\\DPA\\data'
    self.files = os.listdir(self.file_path)
    self.create_tables()

  def create_tables(self):
    for file in self.files:
      if file.startswith('IMPORT') and file.endswith('.xlsx'):
        if 'Wahl' in file:
          self.sql.create_table_from_excel(self.file_path + "\\" + file, 'Wahl')
        elif 'Veranstaltungsliste' in file:
          self.sql.create_table_from_excel(self.file_path + "\\" + file, 'Veranstaltung')
        elif 'Raumliste' in file:
          self.sql.create_table_from_excel(self.file_path + "\\" + file, 'Räume')
        

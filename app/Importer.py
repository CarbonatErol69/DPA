

class Importer():

  def __init__(self, sql):
    self.sql = sql
    self.file_path = 'C:\\Users\\T1485841\\Documents\\Projekte\\Schulprojekt\\DPA\\data'
    self.create_tables()

  def create_tables(self):
    for file in self.file_path:

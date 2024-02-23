from SQLiteManager import SQLiteManager
from Importer import Importer

if __name__ == '__main__':
  sql = SQLiteManager()
  importer = Importer(sql)

  #sql.create_table_from_excel('C:\\Users\\T1485841\\Documents\\Projekte\\Schulprojekt\\DPA\\data\\IMPORT BOT1_Veranstaltungsliste.xlsx', 'Wahl')

  sql.select('Wahl')
  
  sql.close_connection()
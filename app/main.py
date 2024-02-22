from SQLiteManager import SQLiteManager

if __name__ == '__main__':
  sql = SQLiteManager()

  #sql.create_table_from_excel('C:\\Users\\T1485841\\Documents\\Projekte\\Schulprojekt\\DPA\\data\\BOT1_Veranstaltungsliste.xlsx', 'Wahl')

  sql.select('Wahl')
  
  sql.close_connection()
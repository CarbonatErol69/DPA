import sqlite3
import pandas as pd

class SQLiteManager:
  def __init__(self, db_name='schedulize.db'):
    self.db_name = db_name
    self.connection = None
    self.cursor = None

    self.connect()

  def connect(self):
    try:
      self.connection = sqlite3.connect(self.db_name)
      self.cursor = self.connection.cursor()
      print(f"Connected to {self.db_name} successfully.")
    except sqlite3.Error as e:
      print(f"Error connecting to {self.db_name}: {e}")

  def create_table(self, table_name, columns):
    try:
      create_table_query = f"CREATE TABLE IF NOT EXISTS {table_name} ({columns})"
      self.cursor.execute(create_table_query)
      self.connection.commit()
      print(f"Table {table_name} created successfully.")
    except sqlite3.Error as e:
      print(f"Error creating table {table_name}: {e}")

  def create_table_from_excel(self, excel_file, table_name):
    try:
      # Read Excel file into a DataFrame
      df = pd.read_excel(excel_file)

      # Ensure column names are suitable for SQLite
      df.columns = df.columns.str.replace(' ', '_')

      # Generate the CREATE TABLE query based on DataFrame columns
      columns_definition = ', '.join([f'"{column}" TEXT' for column in df.columns])
      create_table_query = f'CREATE TABLE IF NOT EXISTS {table_name} ({columns_definition})'

      # Create the table in SQLite
      self.cursor.execute(create_table_query)
      self.connection.commit()
      print(f"Table {table_name} created successfully.")

      # Insert data from the DataFrame into the SQLite table
      df.to_sql(table_name, self.connection, index=False, if_exists='replace')
      print(f"Data from Excel file inserted into {table_name} successfully.")
    except sqlite3.Error as e:
      print(f"Error creating table or inserting data: {e}")


  def insert(self, table_name, data):
    try:
      insert_query = f"INSERT INTO {table_name} VALUES ({','.join(['?']*len(data))})"
      self.cursor.execute(insert_query, data)
      self.connection.commit()
      print(f"Data inserted into {table_name} successfully.")
    except sqlite3.Error as e:
      print(f"Error inserting data into {table_name}: {e}")

  def select(self, table_name):
    try:
      select_query = f"SELECT * FROM {table_name}"
      self.cursor.execute(select_query)
      rows = self.cursor.fetchall()
      for row in rows:
        print(row)
    except sqlite3.Error as e:
      print(f"Error querying data from {table_name}: {e}")

  def close_connection(self):
    if self.connection:
      self.connection.close()
      print(f"Connection to {self.db_name} closed.")

  def update(self, table_name, set_clause, where_clause):
    try:
      update_query = f"UPDATE {table_name} SET {set_clause} WHERE {where_clause}"
      self.cursor.execute(update_query)
      self.connection.commit()
      print(f"Data updated in {table_name} successfully.")
    except sqlite3.Error as e:
      print(f"Error updating data in {table_name}: {e}")

  def delete_data(self, table_name, where_clause):
    try:
      delete_query = f"DELETE FROM {table_name} WHERE {where_clause}"
      self.cursor.execute(delete_query)
      self.connection.commit()
      print(f"Data deleted from {table_name} successfully.")
    except sqlite3.Error as e:
      print(f"Error deleting data from {table_name}: {e}")

# Example usage:
# sqlite_manager = SQLiteManager()
# sqlite_manager.connect()
# sqlite_manager.create_table('example_table', 'id INTEGER PRIMARY KEY, name TEXT, age INTEGER')
# sqlite_manager.insert_data('example_table', (1, 'John Doe', 25))
# sqlite_manager.query_data('example_table')
# sqlite_manager.update_data('example_table', 'age = 30', 'id = 1')
# sqlite_manager.delete_data('example_table', 'id = 1')
# sqlite_manager.close_connection()

#TODO: chnage this function (class?!) so that it may be imported into GUI / Logic

#TODO: change to correct language

import sqlite3

#con = object representing connection to on disk db
con = sqlite3.connect("schedulize.db")

#create curser to exec sql statements and fetch reuslts
cur = con.cursor()

#TODO: change to correct langauege
cur.execute("CREATE TABLE students (klasse, name, vorname, wahl1, wahl2, wahl3, wahl4, wahl5, wahl6)")

#querying the master table to see if table has been created

res = cur.execute("SELECT name from sqlite_master")

table_name = res.fetchone()[0]

print(f"Table '{table_name}' has been created.")

def create_connection(db_file):
    try:
        con = sqlite3.connect("schedulize.db")
        return con
    except sqlite3.Error as e:
        print(e)
        return None
    
print("testing")
    
#db_file = "students.db"
#con = create_connection(db_file)
    
'''!!!ALWAYS CLOSE CONNECTION WHEN DONE!!!'''

# created if student table doesnt already exist
def create_student_table(con):

    sql = '''
        CREATE TABLE IF NOT EXISTS Student (
        Klasse TEXT,
        
        )

    '''

    cur = con.cursor()
    cur.execute(sql)
    con.commit()

create_student_table(con)
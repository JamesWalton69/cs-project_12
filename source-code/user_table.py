import mysql.connector as sql

conn = sql.connect(host='localhost', user='root', passwd='manager', database='bank')
cur = conn.cursor()

cur.execute('''
    CREATE TABLE IF NOT EXISTS user_table (
        username VARCHAR(25) PRIMARY KEY,
        passwrd VARCHAR(25) NOT NULL
    )
''')
conn.commit()
print("Table 'user_table' created successfully.")

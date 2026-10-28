import mysql.connector as sql

conn = sql.connect(host='localhost', user='root', passwd='manager', database='bank')
cur = conn.cursor()

cur.execute('''
    CREATE TABLE IF NOT EXISTS transactions (
        acct_no INT,
        date DATE,
        withdrawal_amt FLOAT,
        amount_added FLOAT
    )
''')
conn.commit()
print("Table 'transactions' created successfully.")

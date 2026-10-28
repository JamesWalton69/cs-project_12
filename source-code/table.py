import mysql.connector as sql

conn = sql.connect(host='localhost', user='root', passwd='manager', database='bank')
cur = conn.cursor()

cur.execute('''
    CREATE TABLE IF NOT EXISTS customer_details (
        acct_no INT PRIMARY KEY,
        acct_name VARCHAR(25),
        phone_no BIGINT,
        address VARCHAR(50),
        cr_amt FLOAT
    )
''')
conn.commit()
print("Table 'customer_details' created successfully.")

import datetime as dt
import mysql.connector as sql
import menu

conn = sql.connect(host='localhost', user='root', passwd='manager', database='bank')
cur = conn.cursor()

print('========================= WELCOME TO STARK BANK =========================')
print(dt.datetime.now())
print('1. REGISTER')
print('2. LOGIN')
print('=========================================================================')
print()

n = int(input('enter your choice = '))
print()

if n == 1:
    name = input('Enter a Username = ')
    passwd = input('Enter a 4 DIGIT Password = ')
    query = f"INSERT INTO user_table (username, passwrd) VALUES ('{name}', '{passwd}')"
    cur.execute(query)
    conn.commit()
    print()
    print('USER created successfully!')
    print()
    menu.main_menu()

elif n == 2:
    name = input('Enter your Username = ')
    passwd = input('Enter your 4 DIGIT Password = ')
    query = f"SELECT * FROM user_table WHERE username = '{name}' AND passwrd = '{passwd}'"
    cur.execute(query)
    
    if cur.fetchone() is None:
        print()
        print('Invalid username or password')
    else:
        print()
        print('Login successful!')
        print()
        menu.main_menu()

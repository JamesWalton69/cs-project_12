import datetime as dt
import mysql.connector as sql

conn = sql.connect(host='localhost', user='root', passwd='manager', database='bank')
cur = conn.cursor()

def main_menu():
    while True:
        print()
        print('========================= MAIN MENU =========================')
        print('1. CREATE BANK ACCOUNT')
        print('2. TRANSACTION')
        print('3. CUSTOMER DETAILS')
        print('4. TRANSACTION DETAILS')
        print('5. DELETE ACCOUNT')
        print('6. QUIT')
        print('=============================================================')
        print()

        n = int(input('Enter your CHOICE = '))
        print()

        if n == 1:
            acc_no = int(input('Enter your ACCOUNT NUMBER = '))
            acc_name = input('Enter your ACCOUNT NAME = ')
            ph_no = int(input('Enter your PHONE NUMBER = '))
            add = input('Enter your place = ')
            cr_amt = float(input('Enter your credit amount = '))

            query = f"INSERT INTO customer_details VALUES ({acc_no}, '{acc_name}', {ph_no}, '{add}', {cr_amt})"
            cur.execute(query)
            conn.commit()
            print()
            print('Account Created Successfully!!!!!')

        elif n == 2:
            acct_no = int(input('Enter Your Account Number = '))
            cur.execute(f"SELECT * FROM customer_details WHERE acct_no = {acct_no}")
            data = cur.fetchall()

            if len(data) == 0:
                print()
                print('Account Number Invalid. Please try again.')
            else:
                current_bal = data[0][4]
                print()
                print('1. WITHDRAW AMOUNT')
                print('2. ADD AMOUNT')
                print()
                x = int(input('Enter your CHOICE = '))
                print()
                today_date = dt.date.today()

                if x == 1:
                    amt = float(input('Enter withdrawal amount = '))
                    if amt > current_bal:
                        print()
                        print('Insufficient Balance!')
                    else:
                        cur.execute(f"UPDATE customer_details SET cr_amt = cr_amt - {amt} WHERE acct_no = {acct_no}")
                        cur.execute(f"INSERT INTO transactions VALUES ({acct_no}, '{today_date}', {amt}, 0)")
                        conn.commit()
                        print()
                        print('Account Updated Successfully!!!!!')

                elif x == 2:
                    amt = float(input('Enter amount to be added = '))
                    cur.execute(f"UPDATE customer_details SET cr_amt = cr_amt + {amt} WHERE acct_no = {acct_no}")
                    cur.execute(f"INSERT INTO transactions VALUES ({acct_no}, '{today_date}', 0, {amt})")
                    conn.commit()
                    print()
                    print('Account Updated Successfully!!!!!')

        elif n == 3:
            acct_no = int(input('Enter your account number = '))
            print()
            cur.execute(f"SELECT * FROM customer_details WHERE acct_no = {acct_no}")
            data = cur.fetchall()

            if len(data) == 0:
                print('Invalid Account number')
            else:
                for row in data:
                    print('ACCOUNT NO     =', row[0])
                    print('ACCOUNT NAME   =', row[1])
                    print('PHONE NUMBER   =', row[2])
                    print('ADDRESS        =', row[3])
                    print('CURRENT BALANCE=', row[4])
                    print()

        elif n == 4:
            acct_no = int(input('Enter your account number = '))
            print()
            cur.execute(f"SELECT * FROM customer_details WHERE acct_no = {acct_no}")
            if cur.fetchone() is None:
                print('Invalid Account number')
            else:
                cur.execute(f"SELECT * FROM transactions WHERE acct_no = {acct_no}")
                data = cur.fetchall()
                if len(data) == 0:
                    print('No transactions found for this account.')
                else:
                    print(f"{'ACC NO':<10} {'DATE':<15} {'WITHDRAWAL':<15} {'DEPOSIT':<15}")
                    print('-' * 55)
                    for row in data:
                        print(f"{row[0]:<10} {str(row[1]):<15} {row[2]:<15} {row[3]:<15}")
                    print()

        elif n == 5:
            print('DELETE YOUR ACCOUNT')
            acct_no = int(input('Enter your account number = '))
            cur.execute(f"SELECT * FROM customer_details WHERE acct_no = {acct_no}")
            if cur.fetchone() is None:
                print('Account does not exist!')
            else:
                cur.execute(f"DELETE FROM customer_details WHERE acct_no = {acct_no}")
                cur.execute(f"DELETE FROM transactions WHERE acct_no = {acct_no}")
                conn.commit()
                print()
                print('ACCOUNT DELETED SUCCESSFULLY')

        elif n == 6:
            c = input('DO YOU WANT TO EXIT (y/n) = ')
            if c.lower() == 'y':
                print()
                print('THANK YOU PLEASE VISIT AGAIN')
                break

if __name__ == '__main__':
    main_menu()

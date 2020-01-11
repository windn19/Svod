import pyodbc
from os import getcwd
from pprint import pprint

cur_path = getcwd()
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s\Report.accdb;' % (cur_path))
cursor = conn.cursor()
cursor.execute('select Nom_pact from pos')
numbers = cursor.fetchall()
print(numbers)
for now in numbers:
    cursor.execute(f'''
           select o.Nom_pact, o.Nomer, p.Date_pact, p.Supplier, p.Sum
           from opl o
           left join pos p on o.Nom_pact = p.Nom_pact
           where p.Nom_pact={now[0]}
           ''')
    rows = cursor.fetchall()
    pprint(rows)
    if rows:
        nom_pact, nom_paper, date_pact, supplier, sum_pact = rows[0]
        cursor.execute(f'select Date_pact, Sum from opl where Nom_pact={nom_pact}')
        opls = cursor.fetchall()
        for opl_item in opls:
            print(nom_pact, opl_item)
    else:
        cursor.execute(f'select Nom_pact, Date_pact, Supplier, Sum from pos where Nom_pact={now[0]}')
        post = cursor.fetchall()
        print(post)


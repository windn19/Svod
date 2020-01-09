import pyodbc
from datetime import datetime
from os.path import split
from os import getcwd


conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\python\Svod\Report.accdb;')
cursor = conn.cursor()
# s1 = 43598.0 print(split(opl_file)[1].endswith('xlsx'))
# s2 = 43591.0
# s3 = 'куцй'
s = [515, 520, 513, 525]
# print(s3)
# # print(f'''
# #                      INSERT INTO table3 (First_Name, Last_Name, Age)
# #                      VALUES("{s1}", '{s2}', {s3})
# # 513, 43591.0, куцй, 605681.55
# #                   ''')
# cursor.execute(f'''INSERT INTO Pos (Nom_pact, Date_pact, Supplier, Sum)
#                    VALUES (513, {s2}, '{s3}',  605681.55)
#                   ''')
#
#
#
# conn.commit()
for i in s:
    row = cursor.execute(f'select 1 from pos where Nom_pact={i}')
    print(cursor.fetchall())
    print(list(row))
#if (515) in cursor.fetchall():
#print('yes')

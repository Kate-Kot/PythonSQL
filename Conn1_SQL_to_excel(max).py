import 	pymssql
from pymssql import Error
import pandas as pd
import ConDB

print("Привет")
#Microsoft SQL Server Standard (64-bit)

pot = []
I_max = []
I1 = []
I2 = []
I3 = []
P_max = []
P_avg = []

try:
  conn = pymssql.connect(server=ConDB.server, user=ConDB.user, password=ConDB.password, database=ConDB.database)
  print('Ура! Получилось! Соединение установлено!')

except Error as e:
  print('Ошибка! Соединение не установлено')
  print(e)

cursor = conn.cursor()

Date1 = '2022-12-01 00:00:00.000'
Date2 = '2023-01-01 00:00:00.000'
dd = f'Данные за период с {Date1[:10]} по {Date2[:10]}'
print(dd)

potr1 = ['ТП1_вв1', 'ТП1_вв2', 'ГРЩ_1,1', 'ГРЩ_1,2', 'ТП2_Т1', 'ТП2_Т2', 'ТП3_Т1', 'ТП3_Т2', 'ГРЩ3,1_вв1', 'ГРЩ3,1_вв2', 'ТП4', 'ТП6', 'ГРЩ_4,3', 'Компр_ГП2']
k = 0

while k < len(potr1):
  p = potr1[k]
  if p == "ТП1_вв1":
    Uch1 = 1
    Device1 = 1
  elif p == "ТП1_вв2":
    Uch1 = 1
    Device1 = 2
  elif p == "ГРЩ_1,1":
    Uch1 = 2
    Device1 = 1
  elif p == "ГРЩ_1,2":
    Uch1 = 2
    Device1 = 21
  elif p == "ТП2_Т1":
    Uch1 = 3
    Device1 = 3
  elif p == "ТП2_Т2":
    Uch1 = 3
    Device1 = 4
  elif p == "ТП3_Т1":
    Uch1 = 3
    Device1 = 5
  elif p == "ТП3_Т2":
    Uch1 = 3
    Device1 = 6
  elif p == "ГРЩ3,1_вв1":
    Uch1 = 4
    Device1 = 1
  elif p == "ГРЩ3,1_вв2":
    Uch1 = 4
    Device1 = 6
  elif p == "ТП4":
    Uch1 = 6
    Device1 = 1
  elif p == "ТП6":
    Uch1 = 1
    Device1 = 15
  elif p == "ГРЩ_4,3":
    Uch1 = 6
    Device1 = 16
  elif p == "Компр_ГП2":
    Uch1 = 6
    Device1 = 4
  else:
    print('Ошибка! Неверный потребитель!')

  sql_zapros = str("""SELECT MAX(VAR1)/1000, MAX(VAR2)/1000, MAX(VAR3)/1000, MAX(VAR14)/10000/10, ROUND(AVG(VAR14)/10000/10, 2) AS {}
                FROM dbo.tblPower 
                WHERE [Uch] = {} AND [Device] = {}
                AND [EDate] BETWEEN '{}' AND '{}' 
                AND VAR14/10000 > 1
                """.format(p, Uch1, Device1, Date1, Date2))

  cursor.execute(sql_zapros)
  row = cursor.fetchone()

  while row:
    try:
      print(p, round(max(row[0], row[1], row[2])), row[0], row[1], row[2], row[3], row[4], sep = ' | ')
    except TypeError:
      print(p, 'Нет значений за текущий период')
    pot.append(p)
    try:
      I_max.append(round(max(row[0], row[1], row[2])))
    except TypeError:
      I_max.append(f'{p} Нет значений за текущий период')
    I1.append(row[0])
    I2.append(row[1])
    I3.append(row[2])
    P_max.append(row[3])
    try:
      P_avg.append(round(row[4]))
    except TypeError:
      P_avg.append(f'{p} Нет значений за текущий период')
    
    row = cursor.fetchone()

  k+=1

conn.close()

print(pot, I_max, I1, I2, I3, P_max, P_avg)

dict1 = {'Потребители': pot,
        'I max, A': I_max,
        'I1 max, A': I1,
        'I2 max, A': I2,
        'I3 max, A': I3,
        'P max, кВт': P_max,
        '% загр': P_avg,
        'Date': dd }

print(dict1)

df = pd.DataFrame(dict1)

df.to_excel('/Анализ_нагрузок_по_вводам.xlsx', sheet_name='Потребители', index=False)




print("***УРА***!!!")

#cursor.execute("""
#                SELECT столбцы или * для всех 
#                FROM Таблица
#                SELECT ('столбцы или * для выбора всех столбцов; обязательно')
#                FROM ('таблица; обязательно')
#                WHERE ('условие/фильтрация, например, city = 'Moscow'; необязательно')
#                GROUP BY ('столбец, по которому хотим сгруппировать данные; необязательно')
#                HAVING ('условие/фильтрация на уровне сгруппированных данных; необязательно')
#                ORDER BY ('столбец, по которому хотим отсортировать вывод; необязательно')""")

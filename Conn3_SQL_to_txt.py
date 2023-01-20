import 	pymssql
from pymssql import Error
import ConDB

print("Привет")
#Microsoft SQL Server Standard (64-bit)

try:
  conn = pymssql.connect(server=ConDB.server, user=ConDB.user, password=ConDB.password, database=ConDB.database)
  print('Ура! Получилось! Соединение установлено!')

except Error as e:
  print('Ошибка! Соединение не установлено')
  print(e)

cursor = conn.cursor()

Date1 = '2022-10-01 00:00:00.000'
Date2 = '2022-11-01 00:00:00.000'
print(f'Данные за период с {Date1} по {Date2}')

potr1 = ['ТП1_вв1', 'ТП1_вв2', 'ГРЩ_1,1', 'ГРЩ_1,2', 'ТП2_Т1', 'ТП2_Т2', 'ТП3_Т1', 'ТП3_Т2', 'ГРЩ3,1_вв1', 'ГРЩ3,1_вв2', 'ТП4', 'ТП6', 'ГРЩ_4,3', 'Компр_ГП2']
a = []
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

  sql_zapros = str("""SELECT MAX(VAR1)/1000, MAX(VAR2)/1000, MAX(VAR3)/1000, ROUND(AVG(VAR14)/10000/10, 2) AS {}
                FROM dbo.tblPower 
                WHERE [Uch] = {} AND [Device] = {}
                AND [EDate] BETWEEN '{}' AND '{}' 
                AND VAR14/10000 > 1
                """.format(p, Uch1, Device1, Date1, Date2))

  cursor.execute(sql_zapros)
  row = cursor.fetchone()
  while row:
    try:
      print(p, round(max(row[0], row[1], row[2])), row[0], row[1], row[2], row[3], sep = ' | ')
    except TypeError:
      print(p, 'Нет значений за текущий период')
    a.append(p)
    try:
      a.append(str(round(max(row[0], row[1], row[2]))))
    except TypeError:
      a.append(str(f'{p} Нет значений за текущий период'))
    a.append(str(row[0]))
    a.append(str(row[1]))
    a.append(str(row[2]))
    try:
      a.append(str(round(row[3])))
    except TypeError:
      a.append(str(f'{p} Нет значений за текущий период'))
    
    row = cursor.fetchone()

  k+=1

conn.close()

print(a)
fa = open("/test1.txt", 'w', encoding='utf8')

fa.write(f'Данные за период с {Date1} по {Date2}' + '\n')
for i in a:
  fa.write(i + '\n')

fa.close()

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

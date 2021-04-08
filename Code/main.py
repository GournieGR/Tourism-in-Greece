from matplotlib import pyplot as plt
from urllib.request import urlretrieve as retrieve
from matplotlib.ticker import StrMethodFormatter, MultipleLocator
from tkinter import *
import xlrd
import sqlite3
import csv
import pandas as pd
import heapq
import numpy


url_list = []
urls_to_paths = []
year_list = []
documentID_list = [113865, 113886, 113905, 113925, 198755]  # ID of url

# Store urls from website in url_list.
for i in documentID_list:
    url_list.append("https://www.statistics.gr/el/statistics?p_p_id=documents_WAR_publicationsportlet_INSTANCE_"
                    "VBZOni0vs5VJ&p_p_lifecycle=2&p_p_state=normal&p_p_mode=view&p_p_cacheability=cacheLevelPage&p_p_"
                    "col_id=column-2&p_p_col_count=4&p_p_col_pos=2&_documents_WAR_publicationsportlet_"
                    "INSTANCE_VBZOni0vs5VJ"
                    "_javax.faces.resource=document&_documents_WAR_publicationsportlet_INSTANCE_VBZOni0vs5VJ_"
                    "ln=downloadResources&_"
                    "documents_WAR_publicationsportlet_INSTANCE_VBZOni0vs5VJ_documentID="+str(i)+
                    "&_documents_WAR_publicationsportlet_INSTANCE_VBZOni0vs5VJ_locale=el")

for i in range(2011, 2016):
    urls_to_paths.append("Test"+str(i)+".xls")
    year_list.append(i)

# Loop for transport urls from url_list to excel files who is in urls_to_paths with help of function retrieve.
j = 0
for i in url_list:
    retrieve(i, urls_to_paths[j])
    j += 1

Quarter_list = []  # Store population of tourists per three months.
data_population = []  # Store population of tourists in period 2011-2015
Edit_list = []
data_air = []
data_train = []
data_ship = []
data_car = []
country_list = []
total_list1 = []
total_list2 = []
total_list3 = []
total_list4 = []
total_list5 = []
total_list = [total_list1, total_list2, total_list3, total_list4, total_list5]
k = 0
sum = 0

# The data cleaning process begins with help of list.
for i in urls_to_paths:
  inputWorkbook = xlrd.open_workbook(i)
  inputWorksheet = inputWorkbook.sheet_by_index(11)
  rows = inputWorksheet.nrows
  row = 4

  for count in range(2, 12):
      if count == 2:
          sheet = inputWorkbook.sheet_by_index(count)
          rows = sheet.nrows
          for row1 in range(80, rows):
              if sheet.cell_value(row1, 1) == 'ΓΕΝΙΚΟ ΣΥΝΟΛΟ':
                  Quarter_population = sheet.cell_value(row1, 6)
                  Quarter_list.append(int(Quarter_population))
                  break
      else:
          sheet = inputWorkbook.sheet_by_index(count)
          rows = sheet.nrows
          for row1 in range(rows):
              if sheet.cell_value(row1, 0) == '* Πηγή: Έρευνα Συνόρων της Τράπεζας της Ελλάδος':
                  Edit_list.append(int(sheet.cell_value(row1 - 1, 6)))
                  break
          if count == 5 or count == 8 or count == 11:
              for num in Edit_list:
                  sum += num
              Quarter_list.append(sum)
              Edit_list.clear()
              sum = 0


  for row in range(rows):
      if inputWorksheet.cell_value(row, 6) == 'ΣΥΝΟΛΟ':
          start = row
      elif inputWorksheet.cell_value(row, 1) == 'ΓΕΝΙΚΟ ΣΥΝΟΛΟ':
          end = row

  for row in range(start+2, end):
      if inputWorksheet.cell_value(row, 0) == 'ΙΙ. ΧΩΡΕΣ ΑΣΙΑΣ':
          continue
      elif inputWorksheet.cell_value(row, 0) == 'Ι.ΧΩΡΕΣ ΕΥΡΩΠΗΣ':
          continue
      elif inputWorksheet.cell_value(row, 0) == ' - ΧΩΡΕΣ ΕΥΡΩΠΑΪΚΗΣ ΕΝΩΣΗΣ':
          continue
      elif inputWorksheet.cell_value(row, 1) == 'από τΙς οποίες:':
          continue
      elif inputWorksheet.cell_value(row, 0) == 'ΙΙΙ. ΧΩΡΕΣ ΑΦΡΙΚΗΣ':
          continue
      elif inputWorksheet.cell_value(row, 1) == 'Κροατία (2)':
          continue
      elif inputWorksheet.cell_value(row, 1) == 'Κροατία (1)':
          continue
      elif inputWorksheet.cell_value(row, 1) == 'Κροατία':
          continue
      elif inputWorksheet.cell_value(row, 1) == 'Μη προσδιορίσιμες χώρες ταξιδιωτών':
          continue
      elif inputWorksheet.cell_value(row, 0) == 'ΙV. ΧΩΡΕΣ ΑΜΕΡΙΚΗΣ':
          continue
      elif inputWorksheet.cell_value(row, 1) == 'ΠΓΔΜ':
          continue
      elif inputWorksheet.cell_value(row, 0) == 'V. ΧΩΡΕΣ ΩΚΕΑΝΙΑΣ':
          continue
      country = inputWorksheet.cell_value(row, 1)
      total = int(inputWorksheet.cell_value(row, 6))

      if i == urls_to_paths[0]:
          country_list.append(country)
      total_list[k].append(total)
  k += 1

  data_population.append(int(inputWorksheet.cell_value(end, 6)))
  data_air.append(int(inputWorksheet.cell_value(end, 2)))
  data_train.append(int(inputWorksheet.cell_value(end, 3)))
  data_ship.append(int(inputWorksheet.cell_value(end, 4)))
  data_car.append(int(inputWorksheet.cell_value(end, 5)))

result = []
for i in range(0, len(total_list1)):
    result.append(total_list1[i] + total_list2[i] + total_list3[i] + total_list4[i] + total_list5[i])

biggers_result_list = heapq.nlargest(5, result)

final_list = []
for i in range(0, 5):
    j = result.index(biggers_result_list[i])
    final_list.append(country_list[j])

# Export from excel files and store them to database.
connection = sqlite3.connect(':memory:')
cursor1 = connection.cursor()
cursor2 = connection.cursor()
cursor3 = connection.cursor()
cursor4 = connection.cursor()

cursor1.execute("""CREATE TABLE Arrivals_Tourists (                                                                                    
                    Year integer,
                    Population integer 
                    )""")

cursor2.execute("""CREATE TABLE Transport_Tourists (
                    Year integer,
                    Airplane integer,
                    Train integer,
                    Ship integer,
                    Car integer)""")

cursor3.execute("""CREATE TABLE Country_Arrivals_Tourists(
                     Country text,
                     Total_Population integer)""")

cursor4.execute("""CREATE TABLE Quarter_Arrivals_Tourists(
                    Quarter_Population integer,
                    Years integer)""")

for k in range(0, 5):
    cursor1.execute("INSERT INTO Arrivals_Tourists VALUES (?, ?)", (year_list[k], data_population[k]))
    cursor2.execute("INSERT INTO Transport_Tourists VALUES (?, ?, ?, ?, ?)", (year_list[k], data_air[k], data_train[k],
                                                                              data_ship[k], data_car[k]))
    cursor3.execute("INSERT INTO Country_Arrivals_Tourists VALUES (?, ?)", (final_list[k], biggers_result_list[k]))
connection.commit()

k = 0
for count in range(len(Quarter_list)):
    cursor4.execute("INSERT INTO Quarter_Arrivals_Tourists VALUES (?, ?)", (Quarter_list[count], year_list[k]))
    if count == 3 or count == 7 or count == 11 or count == 15:
        k += 1

cursor1.execute("SELECT * FROM Arrivals_Tourists")
cursor2.execute("SELECT * FROM Transport_Tourists")
cursor3.execute("SELECT * FROM Country_Arrivals_Tourists ")
cursor4.execute("SELECT * FROM Quarter_Arrivals_Tourists")


# Export from database and store them to csv_files.
csv_files = ['Arrivals_Tourists_2011-2015.csv', 'Transport_Tourists_2011-2015.csv', 'Countries_Arrivals_2011-2015.csv',
             'Quarter_Arrivals_2011-2015.csv']
for filename in csv_files:
    with open(filename, 'w', newline='', encoding='utf-8-sig') as csv_file:
        csv_writer = csv.writer(csv_file, delimiter=',')
        if filename == csv_files[0]:
            csv_writer.writerow([n[0] for n in cursor1.description])
            csv_writer.writerows(cursor1)
        elif filename == csv_files[1]:
            csv_writer.writerow([n[0] for n in cursor2.description])
            csv_writer.writerows(cursor2)
        elif filename == csv_files[2]:
            csv_writer.writerow([n[0] for n in cursor3.description])
            csv_writer.writerows(cursor3)
        else:
            csv_writer.writerow([n[0] for n in cursor4.description])
            csv_writer.writerows(cursor4)
connection.close()

# Functions for plots data cleaning.
def TouristArrivals():
    df = pd.read_csv(csv_files[0], delimiter=',', encoding='utf-8-sig')
    plt.figure(figsize=(8, 5))
    plt.style.use('seaborn-notebook')
    plt.bar(df.Year, df.Population, width=0.4, label='Population', color='Brown')
    plt.gca().yaxis.set_major_formatter(StrMethodFormatter('{x:,.0f}'))
    plt.title("Tourist arrivals in Greece", fontdict={'fontsize': 18})
    plt.xlabel("Years")
    plt.ylabel("Populations")
    plt.legend()
    plt.tight_layout()
    plt.show()

def PerMetaphor():
    df = pd.read_csv(csv_files[1], delimiter=',', encoding='utf-8-sig')
    plt.style.use("seaborn")
    plt.title("Tourist arrivals per metaphor", fontdict={'fontsize': 18})
    plt.xlabel("Years")
    plt.ylabel("Populations")
    plt.plot(df.Year, df.Airplane, label='Airplane', color='blue', marker='H')
    plt.plot(df.Year, df.Ship, label='Ship', color='green', marker='H')
    plt.plot(df.Year, df.Car, label='Car', color='magenta', marker='H')
    plt.plot(df.Year, df.Train, label='Train', color='red', marker='H')
    plt.gca().xaxis.set_major_locator(MultipleLocator(1))
    plt.gca().yaxis.set_major_formatter(StrMethodFormatter('{x:,.0f}'))
    plt.gca().xaxis.set_major_formatter(StrMethodFormatter('{x:.0f}'))
    plt.legend()
    plt.tight_layout()
    plt.show()

def PerCountries():
    df = pd.read_csv(csv_files[2], delimiter=',', encoding='utf-8-sig')
    plt.figure(figsize=(8, 5))
    plt.style.use("seaborn")
    plt.bar(df.Country, df.Total_Population, width=0.4)
    plt.gca().yaxis.set_major_formatter(StrMethodFormatter('{x:,.0f}'))
    plt.title("The biggest arrivals of countries", fontdict={'fontsize': 15})
    plt.xlabel("Countries")
    plt.ylabel("Total Populations")
    plt.tight_layout()
    plt.show()

def PerThreeMonths():
    df = pd.read_csv(csv_files[3], delimiter=',', encoding='utf-8-sig')
    x = numpy.arange(2011, 2016, 0.25)
    plt.figure(figsize=(8, 5))
    plt.style.use("seaborn")
    plt.title("Tourist arrivals per three months", fontdict={'fontsize': 18})
    plt.xlabel("Years")
    plt.ylabel("Populations")
    plt.plot(x, df.Quarter_Population, color='black', marker='o')
    plt.gca().yaxis.set_major_formatter(StrMethodFormatter('{x:,.0f}'))
    plt.tight_layout()
    plt.show()

# Graphic User Interface.
root = Tk()
root.title('Tourism')
root.geometry("500x360")
Header_Label = Label(root, text="Tourism in Greece in the period 2011-2015", font=("Helvetica", 18))
l1 = Label(root, text="Tourist arrivals:", font=("Helvetica", 13))
b1 = Button(root, text='appear', command=TouristArrivals, relief='groove')
frame = LabelFrame(root, text='Tourist arrivals... ', padx=5, pady=5)
l2 = Label(frame, text="Per metaphor:", font=("Helvetica", 13))
b2 = Button(frame, text='appear', relief='groove', command=PerMetaphor)
l3 = Label(frame, text="Per countries:", font=("Helvetica", 13))
b3 = Button(frame, text='appear', relief='groove', command=PerCountries)
l4 = Label(frame, text="Per three months:", font=("Helvetica", 13))
b4 = Button(frame, text='appear', relief='groove', command=PerThreeMonths)
button_quit = Button(root, text='Exit Program', command=root.quit, relief='groove')
Header_Label.pack()
l1.pack(pady=3)
b1.pack()
frame.pack(padx=10, pady=10)
l2.pack(pady=3)
b2.pack()
l3.pack(pady=3)
b3.pack()
l4.pack(pady=3)
b4.pack()
button_quit.pack(pady=5)
root.mainloop()









# Projets-Yoan
# Extraction des prix Epex Spot français sur une période désirée.

import pandas as pd
from openpyxl import load_workbook
from datetime import timedelta, date
import datetime
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import style
import time

start_date = date(2016, 12, 8)
end_date = date(2016, 12, 9)

d = start_date
delta = datetime.timedelta(days=7)


book = load_workbook('foo.xlsx')

df_main = pd.DataFrame([])
df_main_2 = pd.DataFrame([])

writer = pd.ExcelWriter("foo.xlsx", engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)  

#CREATION DE LA BDD :
    
while d <= end_date:    
    epex = pd.read_html("https://www.epexspot.com/en/market-data/dayaheadauction/auction-table/"+d.strftime("%Y-%m-%d")+"/FR") 
 
    df = epex[2]

    df_prix = df.loc[[0, 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39, 41, 43, 45, 47], :]
    df_quant = df.loc[::2, :]
    
    df_prix_t = df_prix.transpose()
    df_quant_t = df_quant.transpose()  
    
    df_prix_t_drop = df_prix_t.drop(df_prix_t.index[[0, 1]])
    df_quant_t_drop = df_quant_t.drop(df_quant_t.index[[0, 1]])
    
    df_main = df_main.append(df_prix_t_drop)    
    df_main_2 = df_main_2.append(df_quant_t_drop)
    
    d += delta
    time.sleep(1) #Delais entre les requetes au server EPEX, modifiable. 


#ORDONNEMENT DE LA BDD EN MODE SERIE TEMP :
rows = len(df_main.index)
i = 0

df_main.index = range(len(df_main.index)) #Label corectement les prix
df_main_2.index = range(len(df_main_2.index)) #label correctement les quantités
df_main = df_main.convert_objects(convert_numeric=True) #transforme mes données en numérique
df_main_2 = df_main_2.convert_objects(convert_numeric=True) #transforme mes données en numérique

vertical1 = np.array([[0]])
vertical2 = np.array([[0]])
df_vertical1 = pd.DataFrame(vertical1, columns = [1], dtype = "object")
df_vertical2 = pd.DataFrame(vertical2, columns = [1], dtype = "object")

while i < int(rows):
    df_main_li = df_main.loc[i, 1:]
    df_main_li = df_main_2.loc[i, 1:]

    df_vertical1 = np.vstack((df_vertical1, df_main_li[:, None])) #jsais pas pourquoi il faut faire [:, None] -_-' mais ça marche
    df_vertical2 = np.vstack((df_vertical2, df_main_li[:, None]))
    i += 1


df_temp1 = pd.DataFrame(df_vertical1)  
df_temp2 = pd.DataFrame(df_vertical2)  
'''
#SORTIE GRAPHIQUE
print(df_temp.plot())
plt.plot(df_temp)
plt.xlabel("heures")
plt.ylabel("prix en euro par MW")
plt.title("Prix depuis 2015")
style.use('ggplot')
'''

#STAT DESCRIPTIVE




#df.to_excel(writer, d.strftime("%Y-%m-%d"))
df_main.to_excel(writer, "new_prix")
df_main_2.to_excel(writer, "new_quant")
df_temp1.to_excel(writer, "new_seriep")
df_temp2.to_excel(writer, "new_serieq")
writer.save()





# Importer les packages
from datetime import date
from tkinter import *
from tkinter.filedialog import askopenfilename, asksaveasfilename

import pandas as pd
from tabulate import tabulate

#Filedialog
Tk().withdraw()
chemin1 = askopenfilename()
print(chemin1)

#Importer le fichier excel avec le nom de la page
df1 = pd.read_excel(chemin1,
  sheet_name='DATA_carte_inter_hors_SUAP', engine='openpyxl', usecols=[0,1,2], header= 5-1, skipfooter=1,)
print(tabulate(df1.head(10), headers='keys', tablefmt='psql', showindex = False))

#Incorporer un compte des entités
number1 = df1.shape[0]

#Print le compte des entités
print('compter commune total :', number1)

#Selectionner les lignes supérieures ou égales à 95000
df1 = df1.loc[df1['INSEE'] >= 95000]

#Print le tableau
print(tabulate(df1.head(10), headers='keys', tablefmt='psql', showindex = False))

#Incorporer un compte des entités
number2 = df1.shape[0]

#Print le tableau
print('compter commune 95 :', number2)

#L'index devient INSEE
df1 = df1.set_index('INSEE')

#Exporter en csv
Tk().withdraw()
chemin2 = asksaveasfilename(defaultextension = '.csv', filetypes = [('csv','*.csv')])
df1.to_csv (chemin2)
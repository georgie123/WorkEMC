import os
from tabulate import tabulate as tab

import pandas as pd

import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.collections import PatchCollection
from mpl_toolkits.basemap import Basemap

import numpy as np

from PIL import Image, ImageOps

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

shp_simple_countries = r'C:/Users/Georges/PycharmProjects/data/simple_countries/simple_countries'

workDirectory = r'C:/Users/Georges/Downloads/'

segmentFileName = 'VCS_2021_Delegates&ExhibitorsAllEditions'

outputExcelFile = workDirectory+segmentFileName+'_Counts.xlsx'

# Excel import
inputExcelFile = workDirectory+segmentFileName+'.xlsx'
df = pd.read_excel(inputExcelFile, sheet_name='Sheet1', engine='openpyxl',
                   usecols=['Email Address', 'Job Title', 'Company', 'City', 'State or Province', 'Country Name'])


# COUNT CITY
df['City'] = df['City'].str.upper().str.replace(r'\d+', '').str.title().str.replace(r'[,;.:()%]', '', regex=True)
df['City'] = df['City'].replace(r'^\s+$', np.nan, regex=True)
df_City_count = pd.DataFrame(df.groupby(['Country Name', 'City'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Country Name', 'Total'], ascending=[True, False]).reset_index()
df_City_count = df_City_count.fillna('Unknow')

df_City_count['Percent'] = (df_City_count['Total'] / df_City_count['Total'].sum()) * 100
df_City_count['Percent'] = df_City_count['Percent'].round(decimals=2)


# TERMINAL OUTPUTS AND TESTS
print(tab(df_City_count, headers='keys', tablefmt='psql'))
print("OK, export done!")
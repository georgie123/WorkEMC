import pandas as pd
from tabulate import tabulate
import numpy as np

df_FichierTas = pd.read_csv(r'C:/Users/Georges/Downloads/Georges report on on-demand_AMS purchases (TAS) .csv', sep=',', usecols=[
    'I understand that On-Demand includes access to content on AMS site,...', 'First Name', 'Last Name', 'Email Address', 'CC Email Address', 'Job Title'])

df_FichierMcs = pd.read_csv(r'C:/Users/Georges/Downloads/Georges report on on-demand_AMS purchases (MCS) .csv', sep=',', usecols=[
    'Would you like to purchase On-Demand with AMS Premium Membership?', 'First Name', 'Last Name', 'Email Address', 'CC Email Address', 'Practice Type'])

df_FichierVcs = pd.read_excel(r'C:/Users/Georges/Downloads/Georges report on on-demand_AMS purchases (VCS).xlsx',
                   sheet_name='Georges report on on-demandAMS', engine='openpyxl', usecols=['Parent Attendee: I understand that Vegas Cosmetic Surgery On-Demand...', 'First Name', 'Last Name', 'Email Address', 'Job Title'])


# ADD MISSING FIELD TAS
df_FichierTas['source'] = 'TAS 2021'

# ADD MISSING FIELD MCS
df_FichierMcs['source'] = 'MCS 2021'

# ADD MISSING FIELD VCS
df_FichierVcs['CC Email Address'] = np.NaN
df_FichierVcs['source'] = 'VCS 2021'


TAS = df_FichierTas.loc[df_FichierTas['I understand that On-Demand includes access to content on AMS site,...'] == 1]
MCS = df_FichierMcs.loc[df_FichierMcs['Would you like to purchase On-Demand with AMS Premium Membership?'] == 'Yes, I would like to purchase On-Demand with AMS Premium Membership']
VCS = df_FichierVcs.loc[df_FichierVcs['Parent Attendee: I understand that Vegas Cosmetic Surgery On-Demand...'] == 1]


# Merge the selections
df_Merged = pd.concat([TAS, MCS, VCS], axis=0)


# Re-Organize
df_Merged = df_Merged[['source', 'First Name', 'Last Name', 'Email Address', 'CC Email Address', 'Job Title']]

# Fix fields
df_Merged['First Name'] = df_Merged['First Name'].str.title()
df_Merged['Last Name'] = df_Merged['Last Name'].str.title()
df_Merged['Email Address'] = df_Merged['Email Address'].str.lower()
df_Merged['CC Email Address'] = df_Merged['CC Email Address'].str.lower()

df_Merged.loc[df_Merged['CC Email Address'] == df_Merged['Email Address'], 'CC Email Address'] = np.NaN


print(tabulate(df_Merged.head(10), headers='keys', tablefmt='psql', showindex=False))

# COUNT
number = df_Merged.shape[0]
print('Total:', number)

CountTAS = df_Merged[df_Merged['source'] == 'TAS 2021'].count()['source']
print('From TAS: ', CountTAS)
CountMCS = df_Merged[df_Merged['source'] == 'MCS 2021'].count()['source']
print('From MCS: ', CountMCS)
CountVCS = df_Merged[df_Merged['source'] == 'VCS 2021'].count()['source']
print('From VCS: ', CountVCS)
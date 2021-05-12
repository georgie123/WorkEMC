from datetime import date
import pandas as pd
from tabulate import tabulate

AemFile = 'C:/Users/Georges/Downloads/20210511-AEM-Emails.txt'
file = open(AemFile, 'r')

year = date.today().year

listSubject = []
listFirstname = []
listLastname = []
listEmail = []
listSpeciality = []
listCountry = []

lines = file.readlines()

for num, x in enumerate(lines):
    if x == 'Subject:\tFACE 2021 Program Download - Delegate Prospect Lead\n':
        listSubject.append('AEM FACE DELEGATE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', ''))
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', ''))
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', ''))
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', ''))
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])


MergeLists = list(zip(listSubject, listFirstname, listLastname, listEmail, listSpeciality, listCountry))

df = pd.DataFrame(MergeLists, columns=['source', 'firstname', 'lastname', 'email', 'speciality', 'country'])

print(tabulate(df, headers='keys', tablefmt='psql', showindex=False))

number = df.shape[0]
print(number)

# lines = file.readlines()
# print(lines[2])
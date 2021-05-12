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
listPhone = []
listCompany = []

lines = file.readlines()

for num, x in enumerate(lines):
    if x == 'Subject:\tFACE 2021 Program Download - Delegate Prospect Lead\n':
        listSubject.append('AEM FACE DELEGATE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', ''))
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', ''))
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', ''))
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', ''))
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')

    if x == 'Subject:\tContact us - FACE\n':
        listSubject.append('AEM FACE DELEGATE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', ''))
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', ''))
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', ''))
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', ''))
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')

    if x == 'Subject:\tEUROGIN - Webinar Re-watch Contacts\n':
        listSubject.append('AEM EUROGIN DELEGATE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', ''))
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', ''))
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', ''))
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', ''))
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')

    if x == 'Subject:\tEUROGIN Contact Us 2021\n':
        listSubject.append('AEM EUROGIN DELEGATE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', ''))
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', ''))
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', ''))
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', ''))
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')

    if x == 'Subject:\tFrancophone Workshop - Inscription\n':
        listSubject.append('AEM EUROGIN DELEGATE '+str(year))
        listFirstname.append(lines[num+6].replace('First Name \t', '').replace('\n', '').replace(' \t', ''))
        listLastname.append(lines[num+7].replace('Surname \t', '').replace('\n', '').replace(' \t', ''))
        listEmail.append(lines[num+8].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', ''))
        listSpeciality.append('Other Gynecology Specialty')
        listCountry.append(lines[num+9].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')

    if x == 'Subject:\tAMWC Asia 2021 Program Download - Delegate Prospect Lead\n':
        listSubject.append('AEM AMWC-ASIA DELEGATE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', ''))
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', ''))
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', ''))
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', ''))
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')

    if x == 'Subject:\tAMWC 2021 - I would like to be contacted - Exhibitor\n':
        listSubject.append('AEM AMWC EXHIBITOR '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', ''))
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', ''))
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', ''))
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', ''))
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append(lines[num+24].replace('Mobile (personal) \t', '').replace('\n', '').replace(' \t', ''))
        listCompany.append(lines[num+18].replace('Organization Alias \t', '').replace('\n', '').replace(' \t', ''))

    if x == 'Subject:\tAMWC Monaco 2021 - Contact Us\n':
        listSubject.append('AEM AMWC DELEGATE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', ''))
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', ''))
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', ''))
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', ''))
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')

    if x == 'Subject:\tAMWC Global Contest - Fotona\n':
        listSubject.append('AEM AMWC DELEGATE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', ''))
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', ''))
        listEmail.append(lines[num+9].replace('Email Address(business) \t', '').replace('\n', '').replace(' \t', ''))
        listSpeciality.append('Other')
        listCountry.append(lines[num+10].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')


MergeLists = list(zip(listSubject, listFirstname, listLastname, listEmail, listSpeciality, listCountry, listPhone, listCompany))

df = pd.DataFrame(MergeLists, columns=['source', 'firstname', 'lastname', 'email', 'speciality', 'country', 'phone', 'company'])

print(tabulate(df, headers='keys', tablefmt='psql', showindex=False))

number = df.shape[0]
print(number)

# lines = file.readlines()
# print(lines[2])
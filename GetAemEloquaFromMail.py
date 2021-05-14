import warnings
from pandas.core.common import SettingWithCopyWarning

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill

from datetime import date
import pandas as pd
from tabulate import tabulate

warnings.simplefilter(action='ignore', category=SettingWithCopyWarning)

AemFile = 'C:/Users/Georges/Downloads/AEM-Emails.txt'
file = open(AemFile, 'r')

today = date.today()
myToday = str(today).replace('-', '')

year = date.today().year

listSource = []
listFirstname = []
listLastname = []
listEmail = []
listSpeciality = []
listCountry = []
listPhone = []
listCompany = []

listTheme = []
listType = []

lines = file.readlines()

for num, x in enumerate(lines):
    if x == 'Subject:\tFACE 2021 Program Download - Delegate Prospect Lead\n':
        listSource.append('AEM FACE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', '').title())
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', '').title())
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', '').lower())
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', '').title())
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')
        listTheme.append('AA')
        listType.append('DELEGATE')

    if x == 'Subject:\tContact us - FACE\n':
        listSource.append('AEM FACE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', '').title())
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', '').title())
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', '').lower())
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', '').title())
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')
        listTheme.append('AA')
        listType.append('DELEGATE')

    if x == 'Subject:\tFACE exhibitor pack Download - Lead\n':
        listSource.append('AEM FACE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', '').title())
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', '').title())
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', '').lower())
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', '').title())
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append(lines[num+24].replace('Landline (business) \t', '').replace('\n', '').replace(' \t', ''))
        listCompany.append(lines[num+18].replace('Organization Alias \t', '').replace('\n', '').replace(' \t', ''))
        listTheme.append('AA')
        listType.append('EXHIBITOR')

    if x == 'Subject:\tVISAGE 2021 Exhibitor pack Download - Lead\n':
        listSource.append('AEM VISAGE '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', '').title())
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', '').title())
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', '').lower())
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', '').title())
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append(lines[num+24].replace('Landline (business) \t', '').replace('\n', '').replace(' \t', ''))
        listCompany.append(lines[num+18].replace('Organization Alias \t', '').replace('\n', '').replace(' \t', ''))
        listTheme.append('AA')
        listType.append('EXHIBITOR')

    if x == 'Subject:\tEUROGIN - Webinar Re-watch Contacts\n':
        listSource.append('AEM EUROGIN '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', '').title())
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', '').title())
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', '').lower())
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', '').title())
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')
        listTheme.append('GYN')
        listType.append('DELEGATE')

    if x == 'Subject:\tEUROGIN Contact Us 2021\n':
        listSource.append('AEM EUROGIN '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', '').title())
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', '').title())
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', '').lower())
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', '').title())
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')
        listTheme.append('GYN')
        listType.append('DELEGATE')

    if x == 'Subject:\tFrancophone Workshop - Inscription\n':
        listSource.append('AEM EUROGIN FRENCH-WS '+str(year))
        listFirstname.append(lines[num+6].replace('First Name \t', '').replace('\n', '').replace(' \t', '').title())
        listLastname.append(lines[num+7].replace('Surname \t', '').replace('\n', '').replace(' \t', '').title())
        listEmail.append(lines[num+8].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', '').lower())
        listSpeciality.append('Other Gynecology Specialty')
        listCountry.append(lines[num+9].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')
        listTheme.append('GYN')
        listType.append('DELEGATE')

    if x == 'Subject:\tAMWC Asia 2021 Program Download - Delegate Prospect Lead\n':
        listSource.append('AEM AMWC-ASIA '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', '').title())
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', '').title())
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', '').lower())
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', '').title())
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')
        listTheme.append('AA')
        listType.append('DELEGATE')

    if x == 'Subject:\tAMWC 2021 - I would like to be contacted - Exhibitor\n':
        listSource.append('AEM AMWC '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', '').title())
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', '').title())
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', '').lower())
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', '').title())
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append(lines[num+24].replace('Mobile (personal) \t', '').replace('\n', '').replace(' \t', ''))
        listCompany.append(lines[num+18].replace('Organization Alias \t', '').replace('\n', '').replace(' \t', ''))
        listTheme.append('AA')
        listType.append('EXHIBITOR')

    if x == 'Subject:\tAMWC Monaco 2021 - Contact Us\n':
        listSource.append('AEM AMWC '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', '').title())
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', '').title())
        listEmail.append(lines[num+9].replace('Email Address(personal) \t', '').replace('\n', '').replace(' \t', '').lower())
        listSpeciality.append(lines[num+10].replace('Job Title \t', '').replace('\n', '').replace(' \t', '').title())
        listCountry.append(lines[num+11].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')
        listTheme.append('AA')
        listType.append('DELEGATE')

    if x == 'Subject:\tAMWC Global Contest - Fotona\n':
        listSource.append('AEM AMWC '+str(year))
        listFirstname.append(lines[num+7].replace('First Name \t', '').replace('\n', '').replace(' \t', '').title())
        listLastname.append(lines[num+8].replace('Surname \t', '').replace('\n', '').replace(' \t', '').title())
        listEmail.append(lines[num+9].replace('Email Address(business) \t', '').replace('\n', '').replace(' \t', '').lower())
        listSpeciality.append('Other')
        listCountry.append(lines[num+10].replace('Country \t', '').replace('\n', '').replace(' \t', '')[:-3])
        listPhone.append('')
        listCompany.append('')
        listTheme.append('AA')
        listType.append('DELEGATE')


MergeLists = list(zip(listSource, listTheme, listType, listFirstname, listLastname, listEmail, listSpeciality, listCountry, listPhone, listCompany))

df = pd.DataFrame(MergeLists, columns=['source', 'theme', 'type', 'firstname', 'lastname', 'email', 'speciality', 'country', 'phone', 'company'])

# DEDUPING
dfDeduped = df.drop_duplicates()
dfDeduped['JoomlaFixQuery'] = 'UPDATE joo_acymailing_subscriber SET eloqua = "Yes" WHERE email LIKE "' + dfDeduped['email'] + '" ;'

# CLEANING
ind_drop = dfDeduped[dfDeduped['email'].apply(lambda x: x.endswith('@informa.com'))].index
dfDeduped = dfDeduped.drop(ind_drop)
ind_drop = dfDeduped[dfDeduped['email'].apply(lambda x: x.endswith('@euromedicom.com'))].index
dfDeduped = dfDeduped.drop(ind_drop)
ind_drop = dfDeduped[dfDeduped['email'].apply(lambda x: x == ('georges.hinot@gmail.com'))].index
dfDeduped = dfDeduped.drop(ind_drop)
ind_drop = dfDeduped[dfDeduped['email'].apply(lambda x: x == ('sauveneelaurent@gmail.com'))].index
dfDeduped = dfDeduped.drop(ind_drop)

# PRINT
print(tabulate(dfDeduped, headers='keys', tablefmt='psql', showindex=False))

# COUNT
number = df.shape[0]
print('Original:', number)
numberFinal = dfDeduped.shape[0]
print('Final:', numberFinal)

# EXPORT & EXCEL
outputExcelFile = r'C:/Users/Georges/Downloads/INTEGRATION/'+myToday+'_AEM-Eloqua.xlsx'

dfDeduped.to_excel(outputExcelFile, index=False, sheet_name='AEM-Eloqua', header=['source', 'theme', 'type', 'firstname', 'lastname', 'email', 'speciality', 'country', 'phone', 'company', 'JoomlaFixQuery'])
workbook = openpyxl.load_workbook(outputExcelFile)
worksheet = workbook['AEM-Eloqua']
FullRange = 'A1:' + get_column_letter(worksheet.max_column) + str(worksheet.max_row)
worksheet.auto_filter.ref = FullRange
worksheet.freeze_panes = 'A2'
sheetsLits = workbook.sheetnames
workbook['AEM-Eloqua'].column_dimensions['A'].width = 30
workbook['AEM-Eloqua'].column_dimensions['D'].width = 15
workbook['AEM-Eloqua'].column_dimensions['E'].width =15
workbook['AEM-Eloqua'].column_dimensions['F'].width = 30
workbook['AEM-Eloqua'].column_dimensions['G'].width = 30
workbook['AEM-Eloqua'].column_dimensions['H'].width = 15
workbook['AEM-Eloqua'].column_dimensions['K'].width = 90
for sheet in sheetsLits:
    worksheet = workbook[sheet]
    for cell in workbook[sheet][1]:
        worksheet[cell.coordinate].fill = PatternFill(fgColor = 'FFC6C1C1', fill_type = 'solid')
workbook.save(outputExcelFile)
import xml.etree.ElementTree as ET
from selenium import webdriver

chrome_path = 'C:/Users/Georges/PycharmProjects/chromedriver.exe'
browser = webdriver.Chrome(chrome_path)

tree = ET.parse('C:/Users/Georges/Downloads/NoEloqua.xml')
root = tree.getroot()

# Voir (et enrichir) les groupes de pays tout en bas de ce document.

# Groupe de pays avec optin en simple checkbox (Orthographe Eloqua)
# for NoEloqua in root.findall('NoEloqua'):
#     xml_title = NoEloqua.find('title').text
#     xml_FirstName = NoEloqua.find('FirstName').text
#     xml_LastName = NoEloqua.find('LastName').text
#     xml_Email = NoEloqua.find('Email').text
#     xml_JobTitle = NoEloqua.find('JobTitle').text
#     xml_Country = NoEloqua.find('Country').text
#
#     browser.get('https://www.euromedicom.com/amwc-asia/en/contact/receive-amwc-asia-updates.html')
#
#     try:
#         browser.find_elements_by_class_name('cookieButton')[0].click()
#     except Exception:
#         pass
#
#     title = browser.find_element_by_name('title')
#     firstname = browser.find_element_by_id('FirstName')
#     lastname = browser.find_element_by_name('surname')
#     email = browser.find_element_by_name('personalemail')
#     specialty = browser.find_element_by_name('jobTitle')
#     country = browser.find_element_by_name('countryOfResidence')
#
#     title.send_keys(xml_title)
#     firstname.send_keys(xml_FirstName)
#     lastname.send_keys(xml_LastName)
#     email.send_keys(xml_Email)
#     specialty.send_keys(xml_JobTitle)
#     country.send_keys(xml_Country)
#     browser.find_element_by_name('subscriptionrelaxedcombination').click()
#
#     browser.find_element_by_id('formSubmit').click()

# Groupe spécial (Orthographe Eloqua)
# United States, State Florida par défaut
for NoEloqua in root.findall('NoEloqua'):
    xml_title = NoEloqua.find('title').text
    xml_FirstName = NoEloqua.find('FirstName').text
    xml_LastName = NoEloqua.find('LastName').text
    xml_Email = NoEloqua.find('Email').text
    xml_JobTitle = NoEloqua.find('JobTitle').text
    xml_Country = NoEloqua.find('Country').text

    browser.get('https://www.euromedicom.com/amwc-asia/en/contact/receive-amwc-asia-updates.html')

    try:
        browser.find_elements_by_class_name('cookieButton')[0].click()
    except Exception:
        pass

    title = browser.find_element_by_name('title')
    firstname = browser.find_element_by_id('FirstName')
    lastname = browser.find_element_by_name('surname')
    email = browser.find_element_by_name('personalemail')
    specialty = browser.find_element_by_name('jobTitle')
    country = browser.find_element_by_name('countryOfResidence')
    state = browser.find_element_by_name('stateOfResidence')

    title.send_keys(xml_title)
    firstname.send_keys(xml_FirstName)
    lastname.send_keys(xml_LastName)
    email.send_keys(xml_Email)
    specialty.send_keys(xml_JobTitle)
    country.send_keys(xml_Country)
    state.send_keys('Florida')
    browser.find_element_by_name('subscriptionrelaxedcombination').click()

    browser.find_element_by_id('formSubmit').click()

# Groupe de pays sans optin obligatoire (Orthographe Eloqua)
# for NoEloqua in root.findall('NoEloqua'):
#     xml_title = NoEloqua.find('title').text
#     xml_FirstName = NoEloqua.find('FirstName').text
#     xml_LastName = NoEloqua.find('LastName').text
#     xml_Email = NoEloqua.find('Email').text
#     xml_JobTitle = NoEloqua.find('JobTitle').text
#     xml_Country = NoEloqua.find('Country').text
#
#     browser.get('https://www.euromedicom.com/amwc-asia/en/contact/receive-amwc-asia-updates.html')
#
#     try:
#         browser.find_elements_by_class_name('cookieButton')[0].click()
#     except Exception:
#         pass
#
#     title = browser.find_element_by_name('title')
#     firstname = browser.find_element_by_id('FirstName')
#     lastname = browser.find_element_by_name('surname')
#     email = browser.find_element_by_name('personalemail')
#     specialty = browser.find_element_by_name('jobTitle')
#     country = browser.find_element_by_name('countryOfResidence')
#
#     title.send_keys(xml_title)
#     firstname.send_keys(xml_FirstName)
#     lastname.send_keys(xml_LastName)
#     email.send_keys(xml_Email)
#     specialty.send_keys(xml_JobTitle)
#     country.send_keys(xml_Country)
#
#     browser.find_element_by_id('formSubmit').click()


# Groupe de pays avec les optin en dropdown (Orthographe Eloqua)
# for NoEloqua in root.findall('NoEloqua'):
#     xml_title = NoEloqua.find('title').text
#     xml_FirstName = NoEloqua.find('FirstName').text
#     xml_LastName = NoEloqua.find('LastName').text
#     xml_Email = NoEloqua.find('Email').text
#     xml_JobTitle = NoEloqua.find('JobTitle').text
#     xml_Country = NoEloqua.find('Country').text
#
#     browser.get('https://www.euromedicom.com/amwc-asia/en/contact/receive-amwc-asia-updates.html')
#
#     try:
#         browser.find_elements_by_class_name('cookieButton')[0].click()
#     except Exception:
#         pass
#
#     title = browser.find_element_by_name('title')
#     firstname = browser.find_element_by_id('FirstName')
#     lastname = browser.find_element_by_name('surname')
#     email = browser.find_element_by_name('personalemail')
#     specialty = browser.find_element_by_name('jobTitle')
#     country = browser.find_element_by_name('countryOfResidence')
#     thirdpaty = browser.find_element_by_name('subscriptiongdprthirdparty')
#
#     title.send_keys(xml_title)
#     firstname.send_keys(xml_FirstName)
#     lastname.send_keys(xml_LastName)
#     email.send_keys(xml_Email)
#     specialty.send_keys(xml_JobTitle)
#     country.send_keys(xml_Country)
#     thirdpaty.send_keys('Yes')
#
#     browser.find_element_by_id('formSubmit').click()

####################################################################
# Groupe de pays avec les optin en dropdown (Orthographe Eloqua)
# France
# United Kingdom
# Belgium
# Italy
# Germany
# Netherlands

# Groupe de pays sans optin obligatoire (Orthographe Eloqua)
# Taiwan Region
# South Korea
# Australia
# China
# Afghanistan
# Hong Kong SAR
# Myanmar
# Iran, Islamic Republic of
# Singapore
# Vietnam
# Malaysia
# Iraq
# Japan

# Groupe de pays avec optin en simple checkbox (Orthographe Eloqua)
# Philippines
# Indonesia
# Thailand

# Groupe spécial (Orthographe Eloqua)
# United States
####################################################################

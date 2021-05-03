import xml.etree.ElementTree as ET
from selenium import webdriver

chrome_path = 'C:/Users/Georges/PycharmProjects/chromedriver.exe'
browser = webdriver.Chrome(chrome_path)

tree = ET.parse('C:/Users/Georges/Downloads/NoEloqua.xml')
root = tree.getroot()

for NoEloqua in root.findall('NoEloqua'):
    xml_title = NoEloqua.find('title').text
    xml_FirstName = NoEloqua.find('First_x0020_Name').text
    xml_LastName = NoEloqua.find('Last_x0020_Name').text
    xml_Email = NoEloqua.find('Email').text
    xml_JobTitle = NoEloqua.find('Job_x0020_Title').text

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
    thirdpaty = browser.find_element_by_name('subscriptiongdprthirdparty')

    title.send_keys(xml_title)
    firstname.send_keys(xml_FirstName)
    lastname.send_keys(xml_LastName)
    email.send_keys(xml_Email)
    specialty.send_keys(xml_JobTitle)
    country.send_keys('France')
    thirdpaty.send_keys('Yes')

    browser.find_element_by_id('formSubmit').click()
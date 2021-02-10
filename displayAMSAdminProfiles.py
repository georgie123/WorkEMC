import webbrowser

myIds = [4658665,3866554,4658629,4658670,4658700,4658691,4658650,4652310,4658687,3833831,4658634,4658709,4658672,4552886,4658712,4658605,4658616,4658705,4658611,4658676,4658627,4658710,4658715,3709861,4658713,4658602,4658640,4658684,4658714,4002188,3886191,4658656,4658655,4658688,3846814,4658716,4658613,4658711,4658628,4658641]

prefixURL = 'https://multispecialtysociety.com/backoffice/networks/528/users/'
membershipUrlTab = '/edit#tab-memberships'

chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'

for i in myIds:
    #webbrowser.get(chrome_path).open(prefixURL+str(i)+membershipUrlTab, new=0, autoraise=True)
    print(prefixURL+str(i)+membershipUrlTab)
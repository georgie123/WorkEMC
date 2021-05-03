import webbrowser

myIds = [3871014,5010502,3926117,3887412,4632145,5015182]

prefixURL = 'https://multispecialtysociety.com/backoffice/networks/528/users/'
suffixURL = '/edit'

chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'

for i in myIds:
    webbrowser.get(chrome_path).open(prefixURL+str(i)+suffixURL, new=0, autoraise=True)
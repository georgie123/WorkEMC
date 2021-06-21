import webbrowser
import tkinter as tk

chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'

root = tk.Tk()

# Create a Tkinter variable
tkvar = tk.StringVar(root)

# options
choices = ['Including or excluding all AMS users',
           'Including or excluding Premium AMS users',
           'Including or excluding Russians']
tkvar.set('See the list') # set the default option

def on_selection(value):
    global choice
    choice = value
    root.destroy()

popupMenu = tk.OptionMenu(root, tkvar, *choices, command=on_selection)
tk.Label(root, text="Please choose a group of segments").grid(row=0, column=0, padx=10, pady=5)
popupMenu.grid(row=1, column=0, padx=10, pady=5)

root.mainloop()

SegmentsAllAmsUsers = [
'https://secure.p06.eloqua.com/Main.aspx#segments&id=31304',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=36217',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=36504',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37353',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37358',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37372',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37385',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37453',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=38223',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=39424'
]

SegmentsPremiumAmsUsers = [
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37149',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37217',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37382',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37388',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37517',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37524'
]

SegmentsIncludeExcludeRussians = [
'https://secure.p06.eloqua.com/Main.aspx#segments&id=32494',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=32495',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37406',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37407',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=37353',
'https://secure.p06.eloqua.com/Main.aspx#segments&id=38223'
]

if choice == 'Including or excluding all AMS users':
    for i in SegmentsAllAmsUsers:
        webbrowser.get(chrome_path).open(i, new=0, autoraise=True)

if choice == 'Including or excluding Premium AMS users':
    for i in SegmentsPremiumAmsUsers:
        webbrowser.get(chrome_path).open(i, new=0, autoraise=True)

if choice == 'Including or excluding Russians':
    for i in SegmentsIncludeExcludeRussians:
        webbrowser.get(chrome_path).open(i, new=0, autoraise=True)
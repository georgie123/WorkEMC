import tkinter as tk
root = tk.Tk()

# Create a Tkinter variable
tkvar = tk.StringVar(root)

# options
choices = ['Excluding all AMS users',
           'Lasagne','Fries','Fish','Potatoe']
tkvar.set('See the list') # set the default option

def on_selection(value):
    global choice
    choice = value
    root.destroy()

popupMenu = tk.OptionMenu(root, tkvar, *choices, command=on_selection)
tk.Label(root, text="Please choose a group of segments").grid(row=0, column=0, padx=10, pady=5)
popupMenu.grid(row=1, column =0, padx=10, pady=5)

root.mainloop()



if choice == 'Excluding all AMS users':
    print('gogogogo Excluding all AMS users')

elif choice != 'Excluding all AMS users':
  print('On attend')
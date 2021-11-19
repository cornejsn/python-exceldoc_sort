from typing import runtime_checkable
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

########################################################################################################################
# This program removes several rows from a given Excel spreadsheet and makes a new spreadsheet containing desired cars #
########################################################################################################################

# Allows rows from a runlist that contain a whitelisted make to be added to the final list 
# (any makes not in this list are heavily limited or not considered at all)
def isWhitelistedMake(string): 
    make_whitelist = ['ACURA', 'BMW', 'HONDA', 'HYUNDAI', 'RAM', 'SCION', 'SUBARU', 'TOYOTA', 'VOLKSWAGEN']
    return str(string).upper() in make_whitelist

# Allows certain models from limited makes to be added to the final list
def isWhitelistedModel(string):
    model_whitelist = ["1500", "2500", "CRUZE", "EQUINOX", "EXPRESS", "MALIBU", "SONIC", "AVENGER", "RAM 1500", "RAM 2500", "RAM 3500", "F-150", "RANGER", "ACADIA", "TERRAIN", 
        "COMPASS", 'LIBERTY', 'PATRIOT', 'WRANGLER', 'SORENTO', 'SOUL', 'IS250', 'IS350' 'RX350', 'MAZDA3', 'MAZDA6', 
        'C-CLASS', 'ALTIMA', 'ROGUE', 'VERSA']
    return str(string).upper() in model_whitelist

# Main function that sorts a given runlist based on given parameters
def SortRunlist():
    # Filter the initial lists by Make, Model, Odometer, Year and then outputs a final sorted list based on Make and Model
    file = filedialog.askopenfilename() 
    init_runlist = pd.read_excel(file)

    min_year = int(t1.get())
    max_year = int(t2.get())
    min_mile = int(t3.get())
    max_mile = int(t4.get())

    make_filter = init_runlist["Make"].apply(isWhitelistedMake)
    filtered_by_make = init_runlist[make_filter]
    model_filter = init_runlist["Model"].apply(isWhitelistedModel)
    filtered_by_model = init_runlist[model_filter]

    frames = [filtered_by_make, filtered_by_model]
    filtered_list = pd.concat(frames)

    final_list = filtered_list.loc[(filtered_list["Odometer"] >= min_mile) & (filtered_list["Odometer"] <= max_mile) 
                                    & (filtered_list["Year"] >= min_year) & (filtered_list["Year"] <= max_year)]
    final_list = final_list.sort_values(by=['Make', 'Model'])
                        

    # create excel writer object
    writer = pd.ExcelWriter('output.xlsx')
    # write dataframe to excel
    final_list.to_excel(writer)
    # save the excel
    writer.save()
    messagebox.showinfo('Success!','Sorted list was written to "output.xlsx".')

# Using tkinter to record user input for given runlist, beginning year, and ending year
tk.Tk().withdraw()
win= tk.Tk()
win.title('Runlist Sorter')

l1=tk.Label(win,text="Min. Year:")
t1=tk.Entry(win)
l2=tk.Label(win,text="Max. Year")
t2=tk.Entry(win)
l3=tk.Label(win,text="Min. Mileage")
t3=tk.Entry(win)
l4=tk.Label(win,text="Max. Mileage")
t4=tk.Entry(win)
b1=tk.Button(win,text="Submit",command=SortRunlist)

l1.grid(row=0,column=0, padx=(5,5), pady=5)
t1.grid(row=0,column=1, pady=5)
l2.grid(row=1,column=0, padx=(5,5), pady=5)
t2.grid(row=1,column=1, pady=5)
l3.grid(row=2,column=0, padx=(5,5), pady=5)
t3.grid(row=2,column=1, pady=5)
l4.grid(row=3,column=0, padx=(5,5), pady=5)
t4.grid(row=3,column=1, pady=5)
b1.grid(row=4,column=1, padx=50, pady=5)

def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        win.destroy()

win.protocol("WM_DELETE_WINDOW", on_closing)
win.mainloop()
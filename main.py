from typing import runtime_checkable
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# This program removes several rows from a given Excel spreadsheet and makes a new spreadsheet containing desired cars

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

# Record user input for given runlist, beginning year, and ending year
# Using tkinter to recieve file input
tk.Tk().withdraw()
file = filedialog.askopenfilename()

# #Create an instance of Tkinter frame
# win= tk.Tk()

# #Set the geometry of Tkinter frame
# win.geometry("500x500")

# def display_text():
#    global entry
#    string= entry.get()
#    label.configure(text=string)

# #Initialize a Label to display the User Input
# label= tk.Label(win, text="", font=("Courier 22 bold"))
# label.pack()

# #Create an Entry widget to accept User Input
# entry= tk.Entry(win, width= 40)
# entry.focus_set()
# entry.pack()

# #Create a Button to validate Entry Widget
# tk.Button(win, text= "Okay",width= 20, command= display_text).pack(pady=20)

# win.mainloop()

begin_year = int(input("Enter beginning year: "))
end_year = int(input("Enter end year: "))

# Filter the initial lists by Make, Model, Odometer, Year and then outputs a final sorted list based on Make and Model 
init_runlist = pd.read_excel(file)

make_filter = init_runlist["Make"].apply(isWhitelistedMake)
filtered_by_make = init_runlist[make_filter]
model_filter = init_runlist["Model"].apply(isWhitelistedModel)
filtered_by_model = init_runlist[model_filter]

frames = [filtered_by_make, filtered_by_model]
filtered_list = pd.concat(frames)

final_list = filtered_list.loc[(filtered_list["Odometer"] >= 50000) & (filtered_list["Odometer"] <= 180000) 
                                & (filtered_list["Year"] >= begin_year) & (filtered_list["Year"] <= end_year)]
final_list = final_list.sort_values(by=['Make', 'Model'])
                    

# create excel writer object
writer = pd.ExcelWriter('output.xlsx')
# write dataframe to excel
final_list.to_excel(writer)
# save the excel
writer.save()
print('DataFrame is written successfully to "output.xlsx".')
from typing import runtime_checkable
import pandas as pd

# Removes all rows from an Excel Sheet that contain a blacklisted make
def isBlacklistedMake(string): 
    make_blacklist = ["AUDI", "BUICK", "CADILLAC", "CHRYSLER", "HUMMER", "INFINITI", "JAGUAR", "LAND ROVER", "LINCOLN", 
        "MINI", "MITSUBISHI", "PORSCHE", "SATURN", "SMART", "SUZUKI", "VOLVO"]
    return not (string.upper() in make_blacklist)

filename = input("Enter file name: ")

runlist = pd.read_excel(filename + '.xlsx')
filter = runlist["Make"].apply(isBlacklistedMake)
filtered_df = runlist[filter]

over50k_under180k = runlist.loc[(runlist["Odometer"] >= 50000) & (runlist["Odometer"] <= 180000)]

# create excel writer object
writer = pd.ExcelWriter('output.xlsx')
# write dataframe to excel
filtered_df.to_excel(writer)
# save the excel
writer.save()
print('DataFrame is written successfully to Excel File.')
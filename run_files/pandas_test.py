import os
import pandas as pd

sle_data = pd.read_csv(
    "C:\\Users\\erin.asilo\\Documents\\Erin Automation\\Mist Scrape\\SLE List.csv"
)
sle_data = sle_data.reset_index()
# connect_column = sle_data["Successful Connect"].to_list()
# site_column = sle_data["Site"].to_list()
# print(connect_column)
sites = []
values = []
sle_data["Site"] = sle_data["Site"].astype("string")
sle_data["Successful Connect"] = sle_data["Successful Connect"].astype("string")
for index, row in sle_data.iterrows():
    if len(row["Site"]) <= 3:
        sites.append(row["Site"])
        values.append(row["Successful Connect"])
        # data[row["Site"]] = row["Successful Connect"]

print(sites)
print(values)
# print(len(row["Site"]))

# if len(row["Site"])

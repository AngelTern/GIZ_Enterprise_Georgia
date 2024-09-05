# -*- coding: utf-8 -*-


import pandas as pd
import os
import re
import json


directory = "C:/Users/georg/OneDrive/Desktop/Katsadze_data/data_test"

dataframes = []
filenames = []

for filename in os.listdir(directory):
    # Check if the file is an Excel file
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        filenames.append(filename)
        # Read the Excel file into a DataFrame
        filepath = os.path.join(directory, filename)
        df = pd.read_excel(filepath)
        # Append the DataFrame to the list
        dataframes.append(df)

unique_lineitems_in_seperate_eng = []

for dataframe in dataframes:
    for index,row in dataframe.iterrows():
        if row["LineItemENG"] not in unique_lineitems_in_seperate_eng:
            unique_lineitems_in_seperate_eng.append(row["LineItemENG"])

unique_lineitems_in_seperate_geo = []

for dataframe in dataframes:
    for index,row in dataframe.iterrows():
        if row["LineItemGEO"] not in unique_lineitems_in_seperate_geo:
            unique_lineitems_in_seperate_geo.append(row["LineItemGEO"])

'''
def update_lineitem_geo(lineitem_geo):
    if lineitem_geo == "Retained earnings (Accumulated deficit)":
        print(f"{lineitem_geo} changed")
        return "Retained earnings / (Accumulated deficit)"
    elif lineitem_geo == "Impairment loss/reversal of  financial assets":
        print(f"{lineitem_geo} changed")
        return "Impairment (loss)/reversal of financial assets"
    elif lineitem_geo == "Total comprehensive income" or lineitem_geo == "Total comprehensive income(loss)":
        print(f"{lineitem_geo} changed")
        return "Total comprehensive income / (loss)"
    else:
        return lineitem_geo

# Apply the function to the 'LineItemGEO' column for each DataFrame in the list
for dataframe in dataframes:
    dataframe['LineItemENG'] = dataframe['LineItemENG'].apply(update_lineitem_geo)'''

replace_dict ={
    "Retained earnings (Accumulated deficit)": "Retained earnings / (Accumulated deficit)",
    "Impairment loss/reversal of  financial assets": "Impairment (loss)/reversal of financial assets",
    "Total comprehensive income": "Total comprehensive income / (loss)",
    "Total comprehensive income(loss)": "Total comprehensive income / (loss)"
    }

for i, df in enumerate(dataframes):
    dataframes[i]["LineItemENG"] = df["LineItemENG"].replace(replace_dict)

all_variables = [
    'Cash and cash equivalents', 'Current Inventory', 'Non current inventory', 'Trade receivables',
    'Biological assets', 'Other current assets', 'Other non current assets', 'Property, plant and equipment',
    'Total assets', 'Trade payables', 'Provisions for liabilities and charges', 'Total liabilities',
    'Share premium', 'Treasury shares', 'Retained earnings / (Accumulated deficit)', 'Other reserves',
    'Total equity', 'Total liabilities and equity', 'Inventories', 'Net Revenue', 'Cost of goods sold', 
    'Gross profit', 'Other operating income', 'Personnel expense', 'Rental expenses', 'Depreciation and amortisation',
    'Other administrative and operating expenses', 'Operating income', 'Impairment (loss)/reversal of financial assets', 
    'Net gain (loss) from foreign exchange operations', 'Dividends received', 'Other net operating income/(expense)', 
    'Profit/(loss) before tax from continuing operations', 'Income tax', 'Profit/(loss)', 
    'Revaluation reserve of property, plant and equipment', 'Other (include Share of associates and joint ventures in revaluation reserve of property, plant and equipment and defined benefit obligation)',
    'Total other comprehensive (loss) income', 'Total comprehensive income / (loss)', 
    'Net cash from operating activities', 'Net cash used in investing activities', 'Net cash raised in financing activities', 
    'Net cash inflow for the year', 'Effect of exchange rate changes on cash and cash equivalents', 
    'Cash at the beginning of the year', 'Cash at the end of the year '
]

def flatten_lineitem(input):
    processed = re.sub(r'[^a-zA-Z]', '', str(input))
    return processed.lower()                           

#all_variables = list(map(flatten_lineitem, all_variables))



corresponding_lineitems_for_df = {}
geo_lineitems_for_df = {}

for name in filenames:
    corresponding_lineitems_for_df[name] = {}                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              


for i, dataframe in enumerate(dataframes):
    
    if "LineItemENG" not in dataframe.columns:
        dataframe = dataframe.rename(columns={'LineItem': 'LineItemGEO'})
        dataframe["LineItemENG"] = None
    
    current_df_key = list(corresponding_lineitems_for_df.keys())[i]
    if current_df_key.startswith(("2021", "2022")):
        for index, row in dataframe.iterrows():
            lineitem_eng = row['LineItemENG'].rstrip()
            lineitem_geo = row['LineItemGEO']
            if lineitem_eng not in list(corresponding_lineitems_for_df[current_df_key].keys()) and lineitem_eng in all_variables:
                corresponding_lineitems_for_df[current_df_key][lineitem_eng] = []
            elif lineitem_eng in list(corresponding_lineitems_for_df[current_df_key].keys()) and lineitem_geo not in corresponding_lineitems_for_df[current_df_key][lineitem_eng]:
                corresponding_lineitems_for_df[current_df_key][lineitem_eng].append(lineitem_geo)
                

for i in all_variables:
    if i not in list(corresponding_lineitems_for_df["2021 Lineitems Cat IV, part 5.xlsx"].keys()):
        print(i)

print(filenames)

found_retained_earnings = 0



dataframes[1]['LineItemENG'] = dataframes[1]["LineItemENG"].replace(replace_dict)

for index, row in dataframes[1].iterrows():
    if row["LineItemENG"] == "Retained earnings / (Accumulated deficit)":
        found_retained_earnings += 1

print(found_retained_earnings)
        
def choose_lineitem_geo(df, lineitem_to_change, lineitem_main):    
    df['LineItemGEO'] = df['LineItemGEO'].replace(lineitem_to_change, lineitem_main)


with open('lineitem_data/2021.json', 'w', encoding='utf-8') as file:
    json.dump(corresponding_lineitems_for_df, file, ensure_ascii=False, indent=4)




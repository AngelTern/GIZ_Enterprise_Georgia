# -*- coding: utf-8 -*-
"""
Created on Tue Sep  3 21:48:29 2024

@author: george
"""

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


for index, value in enumerate(dataframes):
    dataframes[index] = value[(value["CategoryMain"] != "III ჯგუფი") & (value["FormName"] != "ფინანსური ინსტიტუტებისთვის (გარდა მზღვეველებისა)") 
                              & (value["LineItemGEO"] != "მარაგების გაუფასურების (ხარჯი) / აღდგენა")]


unique_items_in_2022 = []
for index,row in dataframes[1].iterrows():
    if row["LineItemGEO"] not in unique_items_in_2022:
        unique_items_in_2022.append(row["LineItemGEO"])

unique_items_in_2022_eng = []
for index,row in dataframes[1].iterrows():
    if row["LineItemENG"] not in unique_items_in_2022_eng:
        unique_items_in_2022_eng.append(row["LineItemENG"])

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

'''for key, df in enumerate(dataframes):
    dataframes[key] = df[df['CategoryMain'] != "III ჯგუფი"]'''
    


all_variables = [
    'Cash and cash equivalents', 'Current Inventory', 'Non current inventory', 'Trade receivables',
                                          'Biological assets', 'Other current assets', 'Other non current assets', 'Property, plant and equipment',
                                          'Total assets', 'Trade payables', 'Provisions for liabilities and charges', 'Total liabilities',
                                          'Share premium', 'Treasury shares', 'Retained earnings / (Accumulated deficit)', 'Other reserves',
                                          'Total equity', 'Total liabilities and equity', 'Cash advances made to other parties', 'Investment property',
                                          'Investments in subsidiaries', 'Goodwill', 'Other intangible assets', 'Finance lease payable', 'Unearned income',
                                          'Current borrowings', 'Non current borrowings', 'Received grants', 'Total current assets', 'Total current liabilities',
                                          'Share capital', 'Net Revenue', 'Cost of goods sold', 'Gross profit', 'Other operating income',
                                          'Personnel expense', 'Rental expenses', 'Depreciation and amortisation',
                                          'Other administrative and operating expenses', 'Operating income', 
                                          'Impairment (loss)/reversal of financial assets', 'Inventories',
                                          'Net gain (loss) from foreign exchange operations', 'Dividends received',
                                          'Other net operating income/(expense)', 'Profit/(loss) before tax from continuing operations',
                                          'Income tax', 'Profit/(loss)', 'Revaluation reserve of property, plant and equipment',
                                          'Other (include Share of associates and joint ventures in revaluation reserve of property, plant and equipment and defined benefit obligation)',
                                          'Total other comprehensive (loss) income', 'Total comprehensive income / (loss)', 'Net cash from operating activities', 'Net cash used in investing activities',
                                        'Net cash raised in financing activities', 'Net cash inflow for the year',
                                        'Effect of exchange rate changes on cash and cash equivalents',
                                        'Cash at the beginning of the year', 'Cash at the end of the year'
]


replace_dict_eng ={
    "Retained earnings (Accumulated deficit)": "Retained earnings / (Accumulated deficit)",
    "Impairment loss/reversal of  financial assets": "Impairment (loss)/reversal of financial assets",
    "Total comprehensive income": "Total comprehensive income / (loss)",
    "Total comprehensive income(loss)": "Total comprehensive income / (loss)",
    "Prepayments": "Cash advances made to other parties",
    "Cash advances to other parties": "Cash advances made to other parties",
    'Share capital (in case of Limited Liability Company - "capital", in case of cooperative entity - "unit capital"': "Share capital",
    "    - inventories": "Inventories"
    }

replace_dict_geo ={
    "ამონაგები": "ნეტო ამონაგები",
    "სხვა პირებზე ავანსებად და სესხებად გაცემული ფულადი სახსრები": "სხვა მხარეებზე ავანსებად გაცემული ფულადი სახსრები"
    #"ფინანსური აქტივების გაუფასურების (ხარჯი) / აღდგენა" : "გაუფასურების (ხარჯი) / აღდგენა ფინანსურ აქტივებზე" 
    }

for i, df in enumerate(dataframes):
    if "LineItemENG" in df.columns:
        dataframes[i]["LineItemENG"] = df["LineItemENG"].replace(replace_dict_eng)
    else: pass
    if "LineItemGEO" in df.columns:
        dataframes[i]["LineItemGEO"] = df["LineItemGEO"].replace(replace_dict_geo)
    elif "LineItem" in df.columns:
        dataframes[i]["LineItem"] = df["LineItem"].replace(replace_dict_geo)


filtered_df = dataframes[1][(dataframes[1]["LineItemGEO"] == "მარაგები") 
                           ]







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
    if i not in list(corresponding_lineitems_for_df["2022 Lineitems Cat III.xlsx"].keys()):
        print(i)   

geo_unique_2022 = []

for key, value in corresponding_lineitems_for_df["2022 Lineitems Cat III.xlsx"].items():
    for i in value:
        geo_unique_2022.append(i)

geo_unique_2021 = []

for key, value in corresponding_lineitems_for_df["2021 Lineitems Cat III.xlsx"].items():
    for i in value:
        geo_unique_2021.append(i)



with open('lineitem_data/geo_lineitems.json', 'w', encoding='utf-8') as file:
    json.dump(geo_unique_2021, file, ensure_ascii=False, indent=4)
    
with open('lineitem_data/corresponding_lineitems.json', 'w', encoding='utf-8') as file:
    json.dump(corresponding_lineitems_for_df['2021 Lineitems Cat III.xlsx'], file, ensure_ascii=False, indent=4)
    

dataframes_new = []
filenames_new = []
directory_new = "C:/Users/georg/OneDrive/Desktop/Katsadze_data/data_test_2"

for filename in os.listdir(directory_new):
    # Check if the file is an Excel file
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        filenames_new.append(filename)
        # Read the Excel file into a DataFrame
        filepath = os.path.join(directory_new, filename)
        df = pd.read_excel(filepath)
        # Append the DataFrame to the list
        dataframes_new.append(df)

lineitem_geo = geo_unique_2021

unique_lineitems_in_seperate_new= []

for dataframe in dataframes_new:
    for index,row in dataframe.iterrows():
        if row["LineItem"] not in unique_lineitems_in_seperate_new:
            unique_lineitems_in_seperate_new.append(row["LineItem"])

for i in lineitem_geo:
    if i not in unique_lineitems_in_seperate_new:
        print(i)


#df.rename(columns={'LineItem': 'LineItemGEO'}, inplace=True)



for index, value in enumerate(dataframes_new):
    dataframes_new[index] = value.rename(columns={'LineItem': 'LineItemGEO', "Category": "CategoryMain"})
    print(value.columns)
    dataframes_new[index] = value[(value["CategoryMain"] != "III ჯგუფი") & (value["FormName"] != "ფინანსური ინსტიტუტებისთვის (გარდა მზღვეველებისა)") 
                              & (value["LineItemGEO"] != "მარაგების გაუფასურების (ხარჯი) / აღდგენა")]
    dataframes_new[index]["LineItemGEO"] = value["LineItemGEO"].replace(replace_dict_geo)



'''
unique_geo = []

for df in dataframes:
    for value in df["LineItemGEO"].tolist():
        if value not in unique_geo:
            unique_geo.append(value)'''

print(corresponding_lineitems_for_df.keys())







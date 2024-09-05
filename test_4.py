# -*- coding: utf-8 -*-
"""
Created on Mon Sep  2 01:38:12 2024

@author: george
"""

import pandas as pd
import os
import re
import json


directory = "C:/Users/georg/OneDrive/Desktop/Katsadze_data/data_test_2"

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


for i, df in enumerate(dataframes):
    dataframes[i].rename(columns={"LineItem": "LineItemGEO"}, inplace=True)


'''unique_lineitems_in_seperate_eng = []

for dataframe in dataframes:
    for index,row in dataframe.iterrows():
        if row["LineItemENG"] not in unique_lineitems_in_seperate_eng:
            unique_lineitems_in_seperate_eng.append(row["LineItemENG"])'''

unique_lineitems_in_seperate_geo = []

for dataframe in dataframes:
    for index,row in dataframe.iterrows():
        if row["LineItemGEO"] not in unique_lineitems_in_seperate_geo:
            unique_lineitems_in_seperate_geo.append(row["LineItemGEO"])
            
with open('lineitem_data/2021.json', 'r', encoding='utf-8') as file:
    data_2021 = json.load(file)
    
with open('lineitem_data/2022.json', 'r', encoding='utf-8') as file:
    data_2022 = json.load(file)

data_together = {**data_2021, **data_2022}

main_list = []

for main_dict in data_together.values():
    for value in main_dict.values():
        for i in value:
            if i not in main_list:
                main_list.append(i)



values_not_found = []

for i in main_list:
    if i not in unique_lineitems_in_seperate_geo:
        values_not_found.append(i)


    
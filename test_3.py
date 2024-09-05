# -*- coding: utf-8 -*-
"""
Created on Tue Aug 27 14:11:29 2024

@author: george
"""

import pandas as pd
import os
import json

with open('lineitem_data/2021.json', 'r', encoding='utf-8') as file:
    data_2021 = json.load(file)
    
with open('lineitem_data/2022.json', 'r', encoding='utf-8') as file:
    data_2022 = json.load(file)
    
for data in data_2021.values():
    to_pirnt = data.values()
    if len(to_pirnt) > 2:
        print(to_pirnt)
        
        
first_key = next(iter(data_2022))
first_value = data_2022[first_key]

for value in first_value.values():
    if len(value) > 1:
        print(value)

main_list = list()

for main_dict in data_2022.values():
    for value in main_dict.values():
        if len(value) > 1:
            main_list.append(value)

unique_pairs_set = {tuple(pair) for pair in main_list}
unique_pairs_list = list(unique_pairs_set)

with open('lineitem_data/unique_pairs_2022.json', 'w', encoding='utf-8') as file:
    json.dump(unique_pairs_list, file, ensure_ascii=False, indent=4)
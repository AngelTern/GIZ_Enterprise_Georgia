import pandas as pd
import os
import re

directory = "C:/Users/georg/OneDrive/Desktop/Katsadze_data/data_test"

dataframes = []


for filename in os.listdir(directory):
    # Check if the file is an Excel file
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        # Read the Excel file into a DataFrame
        filepath = os.path.join(directory, filename)
        df = pd.read_excel(filepath)
        # Append the DataFrame to the list
        dataframes.append(df)

for i, df in enumerate(dataframes):
    print(f"DataFrame {i+1}:")
    print(df.head(), "\n")
    
combined_df = pd.concat(dataframes, ignore_index=True)

def flatten_lineitem(input, column_name=None):
    # If the input is a pandas DataFrame
    if isinstance(input, pd.DataFrame):
        # Check if the column name is provided and exists in the DataFrame
        if column_name is None or column_name not in input.columns:
            raise ValueError("Provide a valid column name")
        
        # Clean the specified column in the DataFrame
        input[column_name] = input[column_name].apply(lambda x: re.sub(r'[^a-zA-Z]', '', str(x)).lower())
        
        return input
    
    # If the input is a list or another iterable
    elif isinstance(input, (list, pd.Series)):
        # Clean each element in the list or series
        cleaned_input = [re.sub(r'[^a-zA-Z]', '', str(item)).lower() for item in input]
        
        return cleaned_input
    
    else:
        raise TypeError("Input must be either a DataFrame, a list, or a Series")


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
    'Cash at the beginning of the year', 'Cash at the end of the year'
]

all_variables = flatten_lineitem(all_variables)
combined_df = flatten_lineitem(combined_df, column_name="LineItemENG")


unique_values_dict = {var: [] for var in all_variables}

# Iterate over each variable and collect unique corresponding values from the 'Category' column
for var in all_variables:
    corresponding_values = combined_df[combined_df['LineItemENG'] == var]['LineItemGEO'].unique()
    unique_values_dict[var] = corresponding_values.tolist()


found_variables = set(combined_df['LineItemENG'])
not_found_variables = set(all_variables) - found_variables

print("Unique values in the 'Category' column:")
print(unique_values_dict)

# Print the variables that were not found
print("\nVariables not found in the DataFrame:")
print(not_found_variables)

abc = combined_df[combined_df['LineItemENG'] == 'operatingincome']

a = set(abc['LineItemGEO'].to_list())

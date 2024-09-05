import os
import pandas as pd

# Directory containing Excel files
directory = "C:/Users/georg/OneDrive/Desktop/Katsadze_data/data_main"

# Directory to save CSV files
csv_directory = "C:/Users/georg/OneDrive/Desktop/Katsadze_data/data_main_csv"

# Create the CSV directory if it doesn't exist
if not os.path.exists(csv_directory):
    os.makedirs(csv_directory)

# Iterate over all files in the directory
for filename in os.listdir(directory):
    # Check if the file is an Excel file
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        # Read the Excel file into a DataFrame
        filepath = os.path.join(directory, filename)
        df = pd.read_excel(filepath)
        
        # Generate the path for the CSV file
        csv_filename = os.path.splitext(filename)[0] + '.csv'
        csv_filepath = os.path.join(csv_directory, csv_filename)
        
        # Save the DataFrame to a CSV file
        df.to_csv(csv_filepath, index=False)

print("All Excel files have been converted to CSV format and saved in the data_main_csv folder.")

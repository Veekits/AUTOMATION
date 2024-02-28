import os
import pandas as pd
from sklearn.preprocessing import LabelEncoder
import re
import openpyxl

base_dir = 'C:/Users/VMUKITA/OneDrive - Goodlife Pharmacy/Desktop/MY PROJECTS/Automation/AUTOMATION/OUTPUT/Excel Attachments'
output_excel_path = 'C:/Users/VMUKITA/OneDrive - Goodlife Pharmacy/Desktop/MY PROJECTS/Automation/AUTOMATION/SPECIAL ORDER TEMPLATE.csv'
sheet_name = 'order'

def clean_quantity(quantity):
    # Remove non-numeric characters and specific strings
    cleaned_quantity = re.sub(r'[^0-9.]', '', str(quantity))
    return int(float(cleaned_quantity)) if cleaned_quantity else None

def extract_codes_and_quantities(directory, output_excel_path, sheet_name):
    # Create a list to store the file paths
    file_paths = []

    # Create an empty DataFrame to store the results
    result_df = pd.DataFrame(columns=['Code', 'Quantity'])

    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)
                file_paths.append(file_path)

    # Sort the file paths based on their index in the folder name
    file_paths.sort(key=lambda x: int(os.path.basename(os.path.dirname(x)).split('_')[0]))

    # Initialize a label encoder for categorical encoding
    label_encoder = LabelEncoder()

    for file_path in file_paths:
        try:
            # Read the Excel file
            df = pd.read_excel(file_path)

            # Identify columns with codes and quantities
            code_columns = [col for col in df.columns if 'code' in str(col).lower() or 'item no' in str(col).lower()]
            quantity_columns = [col for col in df.columns if 'quantity' in str(col).lower() or 'qty' in str(col).lower()]

            if code_columns and quantity_columns:
                # Extract data from identified columns
                codes_raw = df[code_columns].values.flatten()
                quantities_raw = df[quantity_columns].values.flatten()

                # Clean codes and quantities
                codes = [int(code) if pd.notna(code) else None for code in codes_raw]
                quantities = [clean_quantity(q) for q in quantities_raw]

                # Filter out entries where the cleaned quantity is not None
                non_empty_entries = pd.DataFrame({'Code': [code for code in codes if code is not None],
                                                  'Quantity': [quantity for quantity in quantities if quantity is not None]})

                # Append the results to the DataFrame
                result_df = pd.concat([result_df, non_empty_entries], ignore_index=True)

                # Print the results
                folder_name = os.path.basename(os.path.dirname(file_path))
                print(f"Folder: {folder_name}, File: {os.path.basename(file_path)}")
                for code, quantity in zip(non_empty_entries['Code'], non_empty_entries['Quantity']):
                    print(f"Code: {code}, Quantity: {quantity}")

        except Exception as e:
            print(f"Error processing {file_path}: {e}")

    # Write the results to the Excel file
    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='w') as writer:
        result_df[['Code', 'Quantity']].to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)


# Call the function with the base directory, output path, and sheet name
extract_codes_and_quantities(base_dir, output_excel_path, sheet_name)

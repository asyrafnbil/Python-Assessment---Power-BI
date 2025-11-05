import pandas as pd
import os
import glob

# Set the folder path
folder_path = r'C:\Users\Asyraf Nabil\Documents\Python-Assessment'

# Find Excel file (ignore already processed files)
excel_files = [f for f in glob.glob(os.path.join(folder_path, '*.xls*')) 
               if not f.endswith(('combined_sales.xlsx', 
                                  'combined_sales_with_product.xlsx',
                                  'final_sales_data.xlsx'))]

if not excel_files:
    raise FileNotFoundError("No Excel files found in the folder.")

file_path = excel_files[0]  # Use the first Excel file found
print(f"Using Excel file: {os.path.basename(file_path)}")

# Sheets to combine
sales_sheets = ['2022 Sales', '2021 Sales', '2020 Sales']

# Read and combine sales sheets
df_list = []
for sheet in sales_sheets:
    try:
        df = pd.read_excel(file_path, sheet_name=sheet)
        df['Year'] = sheet.split()[0]  # Add 'Year' column from sheet name
        df_list.append(df)
    except Exception as e:
        print(f"Could not read sheet '{sheet}': {e}")

# Combine all sales into one DataFrame
combined_sales_df = pd.concat(df_list, ignore_index=True)

# Read Product sheet
try:
    product_df = pd.read_excel(file_path, sheet_name='Products')
except Exception as e:
    raise ValueError(f"Could not read 'Products' sheet: {e}")

# Merge sales with Product data on 'Product ID'
merged_df = pd.merge(combined_sales_df, product_df, on='Product ID', how='left')

# Read Locations sheet
try:
    locations_df = pd.read_excel(file_path, sheet_name='Locations')
    merged_df = pd.merge(merged_df, locations_df, on='Location ID', how='left')
except Exception as e:
    print(f"Could not join 'Locations': {e}")

# Read Customers sheet
try:
    customers_df = pd.read_excel(file_path, sheet_name='Customers')
    merged_df = pd.merge(merged_df, customers_df, on='Customer ID', how='left')
except Exception as e:
    print(f"Could not join 'Customers': {e}")

# Format 'Order Date' to dd/mm/yyyy
if 'Order Date' in merged_df.columns:
    merged_df['Order Date'] = pd.to_datetime(merged_df['Order Date'], errors='coerce').dt.strftime('%d/%m/%Y')
else:
    print("'Order Date' column not found.")

# Drop unwanted columns
columns_to_drop = [
    'Year', 'Latitude', 'Longitude', 'Area Code', 'Population', 
    'Households', 'Land Area', 'Water Area', 'Time Zone'
]

# Ensure columns exist before attempting to drop
merged_df = merged_df.drop(columns=[col for col in columns_to_drop if col in merged_df.columns])

print(f"Columns {columns_to_drop} dropped successfully.")

# Save the final DataFrame to a new Excel file with the sheet name 'Dataset'
output_file = os.path.join(folder_path, 'Final File.xlsx')
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:  
    merged_df.to_excel(writer, index=False, sheet_name='Dataset')

print(f"✅ Final sales dataset saved at: {output_file}")

import pandas as pd

# Load the Excel file
file_path = f'C:/Users/2816624M/Desktop/horizontal_tr.xlsx'  # Replace with your Excel file path
xl = pd.ExcelFile(file_path)

# Read the 'Nopeo' sheet
nopeo_df = xl.parse('Nopeo')
# Calculate the column-wise average
column_averages = nopeo_df.mean()

# Initialize an empty DataFrame to store the normalized data
normalized_data_combined = pd.DataFrame()

# Normalize data for each sheet
for sheet in xl.sheet_names:
    # Read the current sheet
    df = xl.parse(sheet)
    
    # Normalize the data by dividing by the 'Nopeo' column averages
    normalized_df = df.div(column_averages, axis='columns')
    
    # Add a column with the name of the sheet to keep track of the origin of data
    normalized_df['Situation'] = sheet
    
    # Combine the normalized data
    normalized_data_combined = pd.concat([normalized_data_combined, normalized_df], axis=0, ignore_index=True)

# Create a new Excel writer object and write the combined DataFrame to a new sheet
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    normalized_data_combined.to_excel(writer, sheet_name='Standardized_Data', index=False)

print('Normalization complete and data combined in a new sheet called "Standardized_Data".')

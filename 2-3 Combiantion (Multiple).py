import pandas as pd
import os

base_path = r'C:\Users\2816624M\Desktop\lab2TX_te'  # The file path of the original files stored
file_names_Nopeo = ['Nopeo_1.xlsx', 'Nopeo_2.xlsx']
file_names_Zone1 = ['Zone1_1.xlsx', 'Zone1_2.xlsx']
file_names_Zone2 = ['Zone2_1.xlsx', 'Zone2_2.xlsx']
file_names_Zone3 = ['Zone3_1.xlsx', 'Zone3_2.xlsx']
file_names_Zone4 = ['Zone4_1.xlsx', 'Zone4_2.xlsx']
file_names_Zone12 = ['Zone12_1.xlsx', 'Zone12_2.xlsx']
file_names_Zone13 = ['Zone13_1.xlsx', 'Zone13_2.xlsx']
file_names_Zone14 = ['Zone14_1.xlsx', 'Zone14_2.xlsx']
file_names_Zone23 = ['Zone23_1.xlsx', 'Zone23_2.xlsx']
file_names_Zone24 = ['Zone24_1.xlsx', 'Zone24_2.xlsx']
file_names_Zone34 = ['Zone34_1.xlsx', 'Zone34_2.xlsx']

# Skipping rows configuration
skip_first_nopeo = 50
skip_last_nopeo = 50
skip_first_Zone1 = 150
skip_last_Zone1 = 150
skip_first_Zone2 = 100
skip_last_Zone2 = 100
skip_first_Zone3 = 100
skip_last_Zone3 = 100
skip_first_Zone4 = 100
skip_last_Zone4 = 100
skip_first_Zone12 = 100
skip_last_Zone12 = 100
skip_first_Zone13 = 100
skip_last_Zone13 = 100
skip_first_Zone14 = 100
skip_last_Zone14 = 100
skip_first_Zone23 = 100
skip_last_Zone23 = 100
skip_first_Zone24 = 100
skip_last_Zone24 = 100
skip_first_Zone34 = 100
skip_last_Zone34 = 100

file_names = ['Nopeo.xlsx', 'Zone1.xlsx', 'Zone2.xlsx', 'Zone3.xlsx', 'Zone4.xlsx','Zone12.xlsx', 'Zone13.xlsx', 'Zone14.xlsx', 'Zone23.xlsx', 'Zone24.xlsx', 'Zone34.xlsx']
sheet_names_combine = ['Nopeo', 'Zone1', 'Zone2', 'Zone3', 'Zone4', 'Zone12', 'Zone13', 'Zone14', 'Zone23', 'Zone24', 'Zone34']
replacements = {'Nopeo': '0', 'Zone1': '1', 'Zone2': '2', 'Zone3': '3', 'Zone4': '4',
                'Zone12': '5', 'Zone13': '6', 'Zone14': '7', 'Zone23': '8', 'Zone24': '9', 'Zone34': '10'}

file_names_combined = []

# ... (Keep the function definitions here)

def combine_excel_files(file_paths, skip_start_first=0, skip_end_last=0):
    """
    Combines multiple Excel files into one DataFrame.
    Args:
    - file_paths: List of paths to Excel files.
    - skip_start_first: Number of rows to skip from the start of the first file, after the first row.
    - skip_end_last: Number of rows to skip from the end of the last file.
    Returns:
    - Combined DataFrame.
    """
    all_data = []

    for i, file in enumerate(file_paths):
        # Skip rows logic
        if i == 0:  # First file
            skip_rows = list(range(1, skip_start_first + 1))  # Skip rows after the first
            df = pd.read_excel(file, skiprows=skip_rows)
        elif i == len(file_paths) - 1:  # Last file
            df = pd.read_excel(file, skipfooter=skip_end_last)
        else:  # Middle files
            df = pd.read_excel(file)  # Skip only the header row

        all_data.append(df)

    # Combine all DataFrames
    combined_df = pd.concat(all_data, ignore_index=True)
    return combined_df

def combine_specific_files(directory_path, file_names, sheet_names, output_file):
    """
    Combine specified files in a given directory into a single Excel file with multiple sheets.

    :param directory_path: Path to the directory containing files.
    :param file_names: List of file names to be combined.
    :param sheet_names: List of sheet names for each file.
    :param output_file: Name of the output Excel file.
    """
    if len(file_names) != len(sheet_names):
        raise ValueError("The number of file names and sheet names must be the same.")

    file_paths = [os.path.join(directory_path, file_name) for file_name in file_names]

    sheets_added = 0
    with pd.ExcelWriter(os.path.join(directory_path, output_file), engine='openpyxl') as writer:
        for file_path, sheet_name in zip(file_paths, sheet_names):
            if not os.path.isfile(file_path):
                print(f"File not found: {file_path}")
                continue

            try:
                df = pd.read_excel(file_path)  # Adjust this if using a different file format
                if df.empty:
                    print(f"Empty DataFrame for file: {file_path}")
                    continue

                if len(sheet_name) > 31 or set('*:/\\?[]').intersection(sheet_name):
                    raise ValueError(f"Invalid sheet name: {sheet_name}")

                df.to_excel(writer, sheet_name=sheet_name, index=False)
                sheets_added += 1
            except Exception as e:
                print(f"Error processing file {file_path}: {e}")

    if sheets_added == 0:
        raise ValueError("No sheets were added. Ensure files are not empty and exist.")


zone_file_paths = {
    'Nopeo': (file_names_Nopeo, skip_first_nopeo, skip_last_nopeo),
    'Zone1': (file_names_Zone1, skip_first_Zone1, skip_last_Zone1),
    'Zone2': (file_names_Zone2, skip_first_Zone2, skip_last_Zone2),
    'Zone3': (file_names_Zone3, skip_first_Zone3, skip_last_Zone3),
    'Zone4': (file_names_Zone4, skip_first_Zone4, skip_last_Zone4),
    'Zone12': (file_names_Zone12, skip_first_Zone12, skip_last_Zone12),
    'Zone13': (file_names_Zone13, skip_first_Zone13, skip_last_Zone13),
    'Zone14': (file_names_Zone14, skip_first_Zone14, skip_last_Zone14),
    'Zone23': (file_names_Zone23, skip_first_Zone23, skip_last_Zone23),
    'Zone24': (file_names_Zone24, skip_first_Zone24, skip_last_Zone24),
    'Zone34': (file_names_Zone34, skip_first_Zone34, skip_last_Zone34),
}

directory_path = r'C:\Users\2816624M\Desktop'  # Your directory path
for zone, (file_names, skip_start, skip_end) in zone_file_paths.items():
    file_paths = [os.path.join(base_path, name) for name in file_names]
    combined_df = combine_excel_files(file_paths, skip_start_first=skip_start, skip_end_last=skip_end)
    output_file = os.path.join(directory_path, f'{zone}.xlsx')
    combined_df.to_excel(output_file, index=False)
    print(f"{zone} data saved to {output_file}")


for zone, (file_names, skip_start, skip_end) in zone_file_paths.items():
    file_paths = [os.path.join(base_path, name) for name in file_names]
    combined_df = combine_excel_files(file_paths, skip_start_first=skip_start, skip_end_last=skip_end)
    
    # Define the name of the output file for the combined data
    combined_file_name = f'{zone}.xlsx'
    output_file = os.path.join(directory_path, combined_file_name)
    combined_df.to_excel(output_file, index=False)
    print(f"{zone} data saved to {output_file}")
    
    # Add the name of the combined file to the list
    file_names_combined.append(combined_file_name)

# Use file_names_combined for the combine_specific_files function
output_file = 'horizontal_tr.xlsx'
combine_specific_files(directory_path, file_names_combined, sheet_names_combine, output_file)
print(f"Combination file saved as {output_file}")

# Load the Excel file
file_path = f'{directory_path}/horizontal_tr.xlsx'
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

with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    normalized_data_combined.to_excel(writer, sheet_name='Standardized_Data', index=False)

print('Normalization complete and data combined in a new sheet called "Standardized_Data".')

file_path = f'{directory_path}/horizontal_tr.xlsx'  # Replace with your file path
excel_data = pd.read_excel(file_path, sheet_name=None)  # Load all sheets
excel_data['Standardized_Data']['Situation'] = excel_data['Standardized_Data']['Situation'].replace(replacements)
with pd.ExcelWriter(file_path) as writer:
    for sheet_name, df in excel_data.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
print("File has been modified and saved.")
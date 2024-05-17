import pandas as pd
import os

base_path = r'C:\Users\' #The file path of the original files stored
file_names_Nopeo = ['Nopeo_1.xlsx', 'Nopeo_2.xlsx']  # Replace with your actual file names
file_names_Zone1 = ['Zone1_1.xlsx', 'Zone1_2.xlsx']  # Replace with your actual file names
file_names_Zone2 = ['Zone2_1.xlsx', 'Zone2_2.xlsx']  # Replace with your actual file names
file_names_Zone3 = ['Zone3_1.xlsx', 'Zone3_2.xlsx']  # Replace with your actual file names
file_names_Zone4 = ['Zone4_1.xlsx', 'Zone4_2.xlsx']  # Replace with your actual file names
skip_first_nopeo = 200
skip_last_nopeo = 200
skip_first_Zone1 = 100
skip_last_Zone1 = 100
skip_first_Zone2 = 100
skip_last_Zone2 = 100
skip_first_Zone3 = 100
skip_last_Zone3 = 100
skip_first_Zone4 = 100
skip_last_Zone4 = 100
file_names = ['Nopeo.xlsx','Zone1.xlsx','Zone2.xlsx','Zone3.xlsx','Zone4.xlsx']      # Replace with your file names
sheet_names_combine = ['Nopeo','Zone1','Zone2','Zone3','Zone4']           # Replace with your desired sheet names
replacements = {'Nopeo': '0', 'Zone1': '1', 'Zone2': '2', 'Zone3': '3','Zone4':'4'}
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


file_paths_Nopeo = [os.path.join(base_path, name) for name in file_names_Nopeo]
combined_df = combine_excel_files(file_paths_Nopeo, skip_start_first=skip_first_nopeo, skip_end_last=skip_last_nopeo)
output_file = r'C:\Users\2816624M\Desktop\Nopeo.xlsx'  # Modify this path and file name as needed
combined_df.to_excel(output_file, index=False)
print(f"Nopeo data saved to {output_file}")

# List of file names, adjust these as per your actual file names
file_paths_Zone1 = [os.path.join(base_path, name) for name in file_names_Zone1]
combined_df = combine_excel_files(file_paths_Zone1, skip_start_first=skip_first_Zone1, skip_end_last=skip_last_Zone1)
output_file = r'C:\Users\2816624M\Desktop\Zone1.xlsx'  # Modify this path and file name as needed
combined_df.to_excel(output_file, index=False)
print(f"Zone1 data saved to {output_file}")


file_paths_Zone2 = [os.path.join(base_path, name) for name in file_names_Zone2]
combined_df = combine_excel_files(file_paths_Zone2, skip_start_first=skip_first_Zone2, skip_end_last=skip_last_Zone2)
output_file = r'C:\Users\2816624M\Desktop\Zone2.xlsx'  # Modify this path and file name as needed
combined_df.to_excel(output_file, index=False)
print(f"Zone2 data saved to {output_file}")


file_paths_Zone3 = [os.path.join(base_path, name) for name in file_names_Zone3]
combined_df = combine_excel_files(file_paths_Zone3, skip_start_first=skip_first_Zone3, skip_end_last=skip_last_Zone3)
output_file = r'C:\Users\2816624M\Desktop\Zone3.xlsx'  # Modify this path and file name as needed
combined_df.to_excel(output_file, index=False)
print(f"Zone3 data saved to {output_file}")

file_paths_Zone4 = [os.path.join(base_path, name) for name in file_names_Zone4]
combined_df = combine_excel_files(file_paths_Zone4, skip_start_first=skip_first_Zone4, skip_end_last=skip_last_Zone4)
output_file = r'C:\Users\2816624M\Desktop\Zone4.xlsx'  # Modify this path and file name as needed
combined_df.to_excel(output_file, index=False)
print(f"Zone4 data saved to {output_file}")

directory_path = r'C:\Users\2816624M\Desktop' # Your directory path
output_file = 'horizontal_tr.xlsx'
combine_specific_files(directory_path, file_names, sheet_names_combine, output_file)
print(f"Combination file saved as {output_file}")


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

with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    normalized_data_combined.to_excel(writer, sheet_name='Standardized_Data', index=False)

print('Normalization complete and data combined in a new sheet called "Standardized_Data".')

file_path = f'C:/Users/2816624M/Desktop/horizontal_tr.xlsx'  # Replace with your file path
excel_data = pd.read_excel(file_path, sheet_name=None)  # Load all sheets
excel_data['Standardized_Data']['Situation'] = excel_data['Standardized_Data']['Situation'].replace(replacements)
with pd.ExcelWriter(file_path) as writer:
    for sheet_name, df in excel_data.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
print("File has been modified and saved.")

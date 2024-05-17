import pandas as pd
import os

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


# Example usage
directory_path = r'C:\Users\2816624M\Desktop' # Your directory path
file_names = ['Nopeo.xlsx','Zone1.xlsx','Zone2.xlsx','Zone3.xlsx','Zone4.xlsx']      # Replace with your file names
sheet_names = ['Nopeo','Zone1','Zone2','Zone3','Zone4']           # Replace with your desired sheet names
output_file = 'horizontal_tr.xlsx'

try:
    combine_specific_files(directory_path, file_names, sheet_names, output_file)
except Exception as e:
    print(f"An error occurred: {e}")



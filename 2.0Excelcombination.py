import pandas as pd
import os

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

# Define the base path
base_path = r'C:\Users\2816624M\Desktop\Area4'


# List of file names, adjust these as per your actual file names
file_names = ['area4_1.xlsx', 'area4_2.xlsx', 'area4_3.xlsx','area4_4.xlsx', 'area4_5.xlsx','area4_6.xlsx']  # Replace with your actual file names

# Combine the base path with each file name
file_paths = [os.path.join(base_path, name) for name in file_names]

# Example usage
combined_df = combine_excel_files(file_paths, skip_start_first=500, skip_end_last=500)

# Specify the output file path and name
output_file = r'C:\Users\2816624M\Desktop\Zone4.xlsx'  # Modify this path and file name as needed

# Save the combined DataFrame to an Excel file
combined_df.to_excel(output_file, index=False)

print(f"Combined data saved to {output_file}")

import serial
import pandas as pd
import openpyxl
from openpyxl import Workbook

# Configuration
serial_port = 'COM4'  # Change this to your serial port
baud_rate = 9600  # Change this to your baud rate
timeout = 1  # Timeout for reading the serial port
data_chunk_size = 1000  # Number of data points per Excel file

# Initialize the serial connection
ser = serial.Serial(serial_port, baud_rate, timeout=timeout)

# Function to save a chunk of data to an Excel file
def save_data_to_excel(data_list, file_number):
    df = pd.DataFrame(data_list, columns=['V1', 'V2', 'V3', 'V4','V5', 'V6', 'V7', 'V8'])
    output_excel_file = f'Zone4_{file_number}.xlsx'  # Unique file name
    df.to_excel(output_excel_file, index=False)
    print(f"Data saved to {output_excel_file}")

# Read serial port data and save to Excel in chunks
def read_serial_data_to_excel(ser):
    # Open the serial connection if it's closed
    if not ser.is_open:
        ser.open()
    
    data_list = []  # List to hold the data
    file_number = 1  # Counter for file naming
    
    try:
        print("Starting data collection. Press Ctrl+C to stop.")
        while True:
            # Read a line from the serial port
            line = ser.readline().decode('utf-8').strip()
            if line:
                # Assume that the data is separated by commas
                values = line.split(',')
                if len(values) == 8:
                    data_list.append(values)
                    print(f"Data received: {values}")
                    # Save to Excel after every 1000 data points
                    if len(data_list) >= data_chunk_size:
                        save_data_to_excel(data_list, file_number)
                        data_list = []  # Reset the list
                        file_number += 1  # Increment the file number
                else:
                    print(f"Line skipped: {line}")

    except KeyboardInterrupt:
        print("\nStopping data collection.")
        if data_list:
            # Save any remaining data to Excel
            save_data_to_excel(data_list, file_number)

    finally:
        # Close the serial connection
        ser.close()

# Call the function to start reading data
read_serial_data_to_excel(ser)

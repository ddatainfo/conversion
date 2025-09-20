import pandas as pd
import numpy as np
import os
from excel_extraction import extract_excel_data
from extract_measurements import extract_measurements

def merge_data(excel_file_path, txt_file_path):
    # Extract data from Excel file
    excel_data = extract_excel_data(excel_file_path)

    # Extract measurements from the TXT file
    measurements = extract_measurements(txt_file_path)

    merged_data = []
    for mes in measurements:
        # Handle cases where dimension does not include a '#' character
        if '#' in mes['dimension']:
            dimension_number = int(mes['dimension'].split('#')[1].split()[0])
        else:
            print(f"Skipping dimension without '#': {mes['dimension']}")
            continue

        if dimension_number in excel_data:
            exc = excel_data[dimension_number]

            # Assign values from mes to exc
            exc['TOLERANCE Max'] = float(mes['+tol'])
            exc['Min'] = float(mes['-tol'])
            exc['Measured'] = float(mes['measured'])
            exc['Deviation'] = float(mes['deviation'])
            exc['Out of Tolerance'] = float(mes['outtol'])

            # Add updated exc to merged data
            merged_data.append(exc)

    return merged_data

if __name__ == "__main__":
    excel_file_path = "/mnt/c/Users/admin/Desktop/conversion/TXT/report/901/PDIR-DAI S10 -901.xlsx"  # Path to Excel file
    txt_folder_path = "/mnt/c/Users/admin/Desktop/conversion/TXT/901.TXT"  # Path to folder containing TXT files

    merged_data = merge_data(excel_file_path, txt_folder_path)

    # Print merged data
    for data in merged_data:
        print(data)

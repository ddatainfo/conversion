import pandas as pd
import numpy as np
import sys
import re
import logging
sys.path.append("../")  # Add parent directory to sys.path for relative imports
from api.utils.excel_extraction import extract_excel_data
from api.utils.extract_measurements import extract_measurements

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def final_data(excel_file_path, txt_file_path, output_file_path):
    logging.info(f"Starting data merging process.")
    logging.debug(f"Excel file path: {excel_file_path}")
    logging.debug(f"TXT file path: {txt_file_path}")

    # Extract data from Excel file
    pre_header, excel_data = extract_excel_data(excel_file_path)

    # Standardize column names to uppercase for consistency
    pre_header.columns = [col.upper() for col in pre_header.columns]
    for k, v in excel_data.items():
        excel_data[k] = {col.upper(): val for col, val in v.items()}

    logging.debug(f"Extracted pre_header columns: {pre_header.columns.tolist()}")
    logging.debug(f"Extracted Excel data keys: {list(excel_data.keys())}")

    # Reset index for pre_header to ensure unique indices
    pre_header = pre_header.reset_index(drop=True)
    logging.debug("Pre-header index reset.")

    # Extract measurements from the TXT file
    logging.info("Extracting measurements from TXT file.")
    measurements = extract_measurements(txt_file_path)
    logging.debug(f"Extracted measurements: {measurements}")

    merged_data = []
    for mes in measurements:
        logging.debug(f"Processing measurement: {mes}")
        # Handle cases where dimension does not include a '#' character
        if '#' in mes['dimension']:
            try:
                # Extract the part after '#', split by non-digit characters, and take the first integer
                dimension_part = mes['dimension'].split('=')[0]
                logging.debug(f"Dimension part after '=': {dimension_part}")
                dimension = re.search(r'#(\d+)', dimension_part)
                logging.debug(f"Extracted dimension number: {dimension}")
                if dimension:
                    dimension_number =int(dimension.group(1))
                else:
                    logging.warning(f"Skipping as re can't find dimension format: {mes['dimension']} - Error: {e}")
                    continue
            except ValueError as e:
                logging.warning(f"Skipping invalid dimension format: {mes['dimension']} - Error: {e}")
                continue
        else:
            logging.warning(f"Skipping dimension without '#': {mes['dimension']}")
            continue

        if dimension_number in excel_data:
            exc = excel_data[dimension_number]
            logging.debug(f"Matching Excel data found for dimension {dimension_number}.")

            # Assign values from mes to exc using uppercase keys
            exc['TOLERANCE MAX'] = float(mes['+tol'])
            exc['TOLERANCE MIN'] = float(mes['-tol'])  # Rename 'Min' to 'Tolerance Min'
            exc['MEASURED'] = float(mes['measured'])
            exc['DEVIATION'] = float(mes['deviation'])
            exc['OUT OF TOLERANCE'] = float(mes['outtol'])
            
            # Add updated exc to merged data
            merged_data.append(exc)
            logging.debug(f"Updated Excel data: {exc}")

    # Convert merged data to DataFrame
    logging.info(f"Data frame:{merged_data}")
    merged_df = pd.DataFrame(merged_data)
    logging.info("Converted merged data to DataFrame.")
    logging.debug(f"Merged DataFrame columns: {merged_df.columns.tolist()}")

    # Reset index for merged_df to ensure unique indices
    merged_df = merged_df.reset_index(drop=True)
    logging.debug("Merged DataFrame index reset.")

    # Filter columns to include only those with matching names
    common_columns = list(set(pre_header.columns) & set(merged_df.columns))
    logging.debug(f"Common columns between pre_header and merged_df: {common_columns}")
    pre_header = pre_header[common_columns]
    merged_df = merged_df[common_columns + ['MEASURED']]  # Add 'Measured' column back

    # Write to Excel file
    merged_df.to_excel(output_file_path, index=False)
    logging.info(f"Merged data written to {output_file_path}")

if __name__ == "__main__":
    excel_file_path = "/mnt/c/Users/admin/Desktop/conversion/TXT/report/901/PDIR-DAI S10 -901.xlsx"  # Path to Excel file
    txt_file_path = "/mnt/c/Users/admin/Desktop/conversion/TXT/901.TXT"  # Path to folder containing TXT files
    output_file_path = "merged_output.xlsx"  # Path to output Excel file

    final_data(excel_file_path, txt_file_path, output_file_path)


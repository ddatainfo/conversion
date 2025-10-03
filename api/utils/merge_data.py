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
    pre_header, excel_data , header_data= extract_excel_data(excel_file_path)
    print("****************")
    logging.debug(f"Header data: {header_data}")
    print("****************")

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
    # Collect measurements which don't match any excel_data dimension
    unmatched_data = []
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
        else:
            # Record unmatched measurement for later review/export
            logging.debug(f"No matching Excel entry for dimension {dimension_number}; saving to unmatched list.")
            unmatched_record = {
                'DIMENSION_NUMBER': dimension_number,
                'DIMENSION': mes.get('dimension'),
                'MEASURED': mes.get('measured'),
                'TOLERANCE_MAX': mes.get('+tol'),
                'TOLERANCE_MIN': mes.get('-tol'),
                'DEVIATION': mes.get('deviation'),
                'OUT_OF_TOLERANCE': mes.get('outtol')
            }
            unmatched_data.append(unmatched_record)

    # Convert merged data to DataFrame
    logging.info(f"Data frame:{merged_data}")
    merged_df = pd.DataFrame(merged_data)
    logging.info("Converted merged data to DataFrame.")
    logging.info(f"Unmatched data records: {len(unmatched_data)}")
    logging.info(f"Unmatched data records: {unmatched_data}")
    logging.debug(f"Merged DataFrame columns: {merged_df.columns.tolist()}")

    # Reset index for merged_df to ensure unique indices
    merged_df = merged_df.reset_index(drop=True)
    logging.debug("Merged DataFrame index reset.")

    # Filter columns to include only those with matching names, in the order of merged_df
    # This preserves the order seen in the merged DataFrame (important for output layout)
    # Build common columns in merged_df order, but exclude any NaN-like labels
    common_columns = [col for col in merged_df.columns if col in pre_header.columns]
    # Remove columns that are actual NaN or the literal string 'NAN'
    removed_nan_cols = [c for c in common_columns if (pd.isna(c) or (isinstance(c, str) and c.strip().upper() == 'NAN'))]
    if removed_nan_cols:
        logging.warning(f"Removing NaN-like columns from common_columns: {removed_nan_cols}")
        common_columns = [c for c in common_columns if c not in removed_nan_cols]
    logging.debug(f"Common columns in merged_df order: {common_columns}")

    # If pre_header has duplicate column labels, reindex will fail. Remove duplicates (keep first).
    dup_cols = pre_header.columns[pre_header.columns.duplicated()].tolist()
    if dup_cols:
        logging.warning(f"Found duplicate columns in pre_header, dropping duplicates (keep first): {dup_cols}")
        # Keep first occurrence of each column label
        pre_header = pre_header.loc[:, ~pre_header.columns.duplicated()]

    # Reindex pre_header to the common columns (safe - will insert NaN for missing)
    pre_header = pre_header.reindex(columns=common_columns)

    # Prepare merged_df keep list, ensure 'MEASURED' is appended if present
    merged_keep = list(common_columns)
    if 'MEASURED' in merged_df.columns and 'MEASURED' not in merged_keep:
        merged_keep.append('MEASURED')
    merged_df = merged_df.reindex(columns=merged_keep)

    # Combine header_data and merged_df
    logging.info("Combining header data and merged DataFrame.")
    header_df = pd.DataFrame(header_data)
    header_df.loc[len(header_df)] = ['head'] * len(header_df.columns)
    # Remove columns from merged_df
    #merged_df = merged_df.drop(columns=merged_df.columns, errors='ignore')
    logging.debug(f"Meged_Df: {merged_df}")

    # Create a replacement row for header_df with merged_df column names
    n_header_cols = len(header_df.columns)
    merged_col_names = [str(c) for c in merged_df.columns]

    # Build a row matching header_df column count: place merged column names left-to-right, pad with empty strings
    replacement_row = [""] * n_header_cols
    for i, name in enumerate(merged_col_names):
        if i < n_header_cols:
            replacement_row[i] = name
        else:
            break

    # Assign the replacement row (safe length)
    header_df.iloc[-1] = replacement_row

    # Now build merged_data_only as rows matching header_df columns.
    n_merged_cols = merged_df.shape[1]
    rows = []
    for arr in merged_df.values:
        rowvals = list(arr)
        if n_merged_cols < n_header_cols:
            # pad with NaN to match header width
            rowvals = rowvals + [np.nan] * (n_header_cols - n_merged_cols)
        elif n_merged_cols > n_header_cols:
            # truncate extra merged columns to fit into header width
            rowvals = rowvals[:n_header_cols]
        rows.append(rowvals)

    merged_data_only = pd.DataFrame(rows, columns=header_df.columns)

    # Combine header_df and merged_data_only
    combined_df = pd.concat([header_df, merged_data_only], ignore_index=True)
    logging.debug(f"Combined DataFrame shape: {combined_df.shape}")

    # Write combined DataFrame to Excel file. If there are unmatched records, write them to a separate sheet.
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='combined', index=False)
        if unmatched_data:
            unmatched_df = pd.DataFrame(unmatched_data)
            unmatched_df.to_excel(writer, sheet_name='unmatched', index=False)

    logging.info(f"Combined data written to {output_file_path} (sheets: combined{', unmatched' if unmatched_data else ''})")

if __name__ == "__main__":
    excel_file_path = "/mnt/c/Users/admin/Desktop/conversion/TXT/report/901/PDIR-DAI S10 -901.xlsx"  # Path to Excel file
    txt_file_path = "/mnt/c/Users/admin/Desktop/conversion/TXT/901.TXT"  # Path to folder containing TXT files
    output_file_path = "merged_output.xlsx"  # Path to output Excel file

    final_data(excel_file_path, txt_file_path, output_file_path)

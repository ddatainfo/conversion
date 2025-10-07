import os
import pandas as pd
import numpy as np
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')



def remove_rows_after(file_path, row_number):
    """
    Reads an Excel file, removes rows after a specific row number, and saves the result as an Excel file.

    :param file_path: Path to the Excel file.
    :param row_number: The row number after which rows should be removed.
    :param output_file: Path to save the modified Excel file.
    """
    logging.info(f"Reading Excel file: {file_path}")

    # Read the Excel file using a compatible engine
    df = _safe_read_excel(file_path, header=None)
    logging.debug(f"Original DataFrame shape: {df.shape}")

    # Remove rows after the specified row number
    df = df.iloc[:row_number]
    logging.info(f"Rows after row {row_number} have been removed.")
    logging.debug(f"Modified DataFrame shape: {df.shape}")

    # Save the modified DataFrame to an Excel file
    # df.to_excel(output_file, index=False, header=False)
    # logging.info(f"Modified DataFrame saved to {output_file}")
    return df


def extract_excel_data(file_path):
    logging.info(f"Starting extraction of data from Excel file: {file_path}")

    # Read the entire sheet without headers using a compatible engine
    df = _safe_read_excel(file_path, header=None)
    # Find the row containing 'Print No'
    header_row_idx = None
    for idx, row in df.iterrows():
        if any(str(cell).strip().lower() == 'print no' for cell in row):
            header_row_idx = idx
            logging.info(f"Header row index found at: {header_row_idx}")
            break
    if header_row_idx is None:
        raise KeyError("Could not find a row containing 'Print No'.")

    # Extract parent and sub-columns
    headers = remove_rows_after(file_path,header_row_idx)
   
    parent_columns = df.iloc[header_row_idx]
    sub_columns = df.iloc[header_row_idx + 1]

    # Combine parent and sub-columns
    combined_columns = []
    for parent, sub in zip(parent_columns, sub_columns):
        parent_str = str(parent).strip() if not pd.isna(parent) else ''
        sub_str = str(sub).strip() if not pd.isna(sub) else ''
        if parent_str and sub_str:
            combined_columns.append(f"{parent_str} {sub_str}")
        elif parent_str:
            combined_columns.append(parent_str)
        elif sub_str:
            combined_columns.append(sub_str)
        else:
            combined_columns.append('nan')

    logging.debug(f"Combined columns before renaming: {combined_columns}")
    # Rename 'Min' to 'Tolerance Min'
    combined_columns = ['TOLERANCE MIN' if col.lower() == 'min' else col for col in combined_columns]

    df.columns = combined_columns
    logging.debug(f"Columns after renaming: {df.columns.tolist()}")

    # Extract rows above header_row_idx while preserving the exact format
    pre_header_df = df.iloc[:header_row_idx].copy()
    pre_header_df.reset_index(drop=True, inplace=True)
    logging.debug("Extracted pre-header data with exact format:")
    logging.debug(pre_header_df)

    # Drop header rows
    df = df.drop(range(header_row_idx + 2))
    df = df.reset_index(drop=True)

    # Convert 'Print No' as key (generic, case-insensitive)
    # Find the column that contains both 'print' and 'no' (case-insensitive)
    key_col = None
    for col in df.columns:
        col_str = str(col).lower().replace(' ', '')
        if 'print' in col_str and 'no' in col_str:
            key_col = col
            break
    if key_col is None:
        raise KeyError("Could not find a column containing both 'print' and 'no'.")

    data_dict = {}
    for _, row in df.iterrows():
        key = row[key_col]
        value = row.to_dict()
        data_dict[key] = value
        #logging.debug(f"Extracted data for {key_col} {key}: {value}")

    logging.info(f"Extraction completed. Total items extracted: {len(data_dict)}")
    return pre_header_df, data_dict, headers


def _safe_read_excel(file_path, **kwargs):
    """Read Excel using an explicit engine based on file extension.

    .xlsx -> openpyxl
    .xls  -> xlrd
    If engine is unavailable, raises ImportError with actionable message.
    """
    ext = os.path.splitext(file_path)[1].lower()
    engine = None
    if ext == '.xls':
        engine = 'xlrd'
    elif ext in ('.xlsx', '.xlsm', '.xltx', '.xltm'):
        engine = 'openpyxl'
    # If extension unknown, let pandas try but prefer openpyxl
    try:
        if engine:
            return pd.read_excel(file_path, engine=engine, **kwargs)
        return pd.read_excel(file_path, **kwargs)
    except ImportError as e:
        # Provide a helpful message about installing the required package
        if engine == 'xlrd':
            raise ImportError("xlrd is required to read .xls files. Install it with: pip install xlrd==1.2.0") from e
        if engine == 'openpyxl':
            raise ImportError("openpyxl is required to read .xlsx files. Install it with: pip install openpyxl") from e
        raise

if __name__ == "__main__":
    file_path = "final_inscpection.xlsx"  # Path to Excel file
    pre_header_df, extracted_data = extract_excel_data(file_path)
    print("Pre-header DataFrame:")
    print(pre_header_df)
    print("Extracted Data:")
    for print_no, data in extracted_data.items():
        print(f"Print No: {print_no}")
        print(data)
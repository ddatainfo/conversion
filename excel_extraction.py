import pandas as pd
import numpy as np

def extract_excel_data(file_path):
    # Read the entire sheet without headers
    df = pd.read_excel(file_path, header=None)
    # Find the row containing 'Print No'
    header_row_idx = None
    for idx, row in df.iterrows():
        if any(str(cell).strip().lower() == 'print no' for cell in row):
            header_row_idx = idx
            print("Header row index:", header_row_idx)
            break
    if header_row_idx is None:
        raise KeyError("Could not find a row containing 'Print No'.")

    # Extract parent and sub-columns
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

    df.columns = combined_columns
    print("Columns:", df.columns.tolist())

    # Drop header rows
    df = df.drop(range(header_row_idx + 2))
    df = df.reset_index(drop=True)

    # Convert 'TOLERANCE Max' and 'TOLERANCE Min' columns to float
    for column in ['TOLERANCE Max', 'TOLERANCE Min']:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors='coerce')  # Convert to float, set invalid values to NaN

    # Now use 'Print No' as key
    data_dict = {}
    for _, row in df.iterrows():
        key = row['Print No']
        value = row.to_dict()
        data_dict[key] = value

    return data_dict

if __name__ == "__main__":
    file_path = "/mnt/c/Users/admin/Desktop/conversion/TXT/report/901/PDIR-DAI S10 -901.xlsx"
    extracted_data = extract_excel_data(file_path)
    for print_no, data in extracted_data.items():
        #print(f"Print No: {print_no}")
        print(data)
        pass
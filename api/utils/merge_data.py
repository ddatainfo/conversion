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


def _try_float(v):
    """Safely convert v to float; return np.nan on failure."""
    try:
        if v is None:
            return np.nan
        return float(v)
    except Exception:
        return np.nan

def final_data(excel_file_path, txt_file_path, output_file_path):
    logging.info("Starting data merging process.")

    # Extract data from Excel file
    pre_header, excel_data, header_data = extract_excel_data(excel_file_path)
    logging.debug(f"excel_Data keys: {list(excel_data)}")

    # Normalize keys inside excel_data templates to uppercase so they match pre_header columns
    for k, v in list(excel_data.items()):
        excel_data[k] = { (col.upper() if isinstance(col, str) else col): val for col, val in v.items() }

    # Normalize pre_header column names to uppercase for matching
    pre_header.columns = [col.upper() for col in pre_header.columns]
    pre_header = pre_header.reset_index(drop=True)

    # Extract measurements
    measurements = extract_measurements(txt_file_path)

    # Group measurements by parsed dimension_number
    measurements_by_dim = {}
    for mes in measurements:
        if '#' in mes.get('dimension', ''):
            try:
                dimension_part = mes.get('dimension', '').split('=')[0]
                dimension = re.search(r'#(\d+)', dimension_part)
                if dimension:
                    dim_no = int(dimension.group(1))
                    measurements_by_dim.setdefault(dim_no, []).append(mes)
            except Exception:
                continue

    logging.debug(f"Collected measurements by dimension: {measurements_by_dim}")

    # Build merged_data by iterating excel_data templates; populate with measurement values when available
    merged_data = []
    unmatched_data = []

    for key, template in excel_data.items():
        # Work on a shallow copy
        excel_template = template.copy()
        mes_list = measurements_by_dim.get(key)
        if mes_list:
            for mes in mes_list:
                exc = excel_template.copy()
                # Overwrite template fields when measurement provides values
                if mes.get('+tol') is not None:
                    exc['TOLERANCE MAX'] = _try_float(mes.get('+tol'))
                if mes.get('-tol') is not None:
                    exc['TOLERANCE MIN'] = _try_float(mes.get('-tol'))
                if mes.get('measured') is not None:
                    exc['MEASURED'] = _try_float(mes.get('measured'))
                else:
                    # Ensure MEASURED key exists
                    exc.setdefault('MEASURED', np.nan)
                exc['DEVIATION'] = _try_float(mes.get('deviation'))
                exc['OUT OF TOLERANCE'] = _try_float(mes.get('outtol'))
                merged_data.append(exc)
        else:
            # No measurements for this template: ensure MEASURED key exists (empty)
            row = {}
            inserted = False
            for k, v in excel_template.items():
                row[k] = v
                if not inserted and k == 'INSTRUMENT':
                    row['MEASURED'] = ''
                    inserted = True
            if not inserted:
                row['MEASURED'] = ''
            merged_data.append(row)

    # Any measurement keys not present in excel_data are unmatched
    unmatched_keys = set(measurements_by_dim.keys()) - set(excel_data.keys())
    for uk in unmatched_keys:
        for mes in measurements_by_dim.get(uk, []):
            unmatched_record = {
                'DIMENSION_NUMBER': uk,
                'DIMENSION': mes.get('dimension'),
                'MEASURED': mes.get('measured'),
                'TOLERANCE_MAX': mes.get('+tol'),
                'TOLERANCE_MIN': mes.get('-tol'),
                'DEVIATION': mes.get('deviation'),
                'OUT_OF_TOLERANCE': mes.get('outtol')
            }
            unmatched_data.append(unmatched_record)

    logging.info(f"Built merged_data rows: {len(merged_data)}; unmatched: {len(unmatched_data)}")

    # Convert merged_data to DataFrame and drop columns that are all NaN
    merged_df = pd.DataFrame(merged_data)
    merged_df = merged_df.dropna(axis=1, how='all')
    logging.debug(f"Merged DataFrame columns after dropping all-NaN: {merged_df.columns.tolist()}")

    # Continue pipeline: filter common columns in merged_df order
    common_columns = [col for col in merged_df.columns if col in pre_header.columns]
    removed_nan_cols = [c for c in common_columns if (pd.isna(c) or (isinstance(c, str) and c.strip().upper() == 'NAN'))]
    if removed_nan_cols:
        logging.warning(f"Removing NaN-like columns from common_columns: {removed_nan_cols}")
        common_columns = [c for c in common_columns if c not in removed_nan_cols]

    # Deduplicate pre_header columns if needed
    if pre_header.columns.duplicated().any():
        dup_cols = pre_header.columns[pre_header.columns.duplicated()].tolist()
        logging.warning(f"Found duplicate columns in pre_header, dropping duplicates (keep first): {dup_cols}")
        pre_header = pre_header.loc[:, ~pre_header.columns.duplicated()]

    pre_header = pre_header.reindex(columns=common_columns)

    # Ensure MEASURED preserved in merged_df
    merged_keep = list(common_columns)
    if 'MEASURED' in merged_df.columns and 'MEASURED' not in merged_keep:
        merged_keep.append('MEASURED')
    merged_df = merged_df.reindex(columns=merged_keep)

    # Build header_df and append merged rows (with safe padding/truncation)
    header_df = pd.DataFrame(header_data)
    header_df.loc[len(header_df)] = ['head'] * len(header_df.columns)

    n_header_cols = len(header_df.columns)
    merged_col_names = [str(c) for c in merged_df.columns]
    replacement_row = [""] * n_header_cols
    for i, name in enumerate(merged_col_names):
        if i < n_header_cols:
            replacement_row[i] = name
        else:
            break
    header_df.iloc[-1] = replacement_row

    # Build rows from merged_df matching header width
    n_merged_cols = merged_df.shape[1]
    rows = []
    for arr in merged_df.values:
        rowvals = list(arr)
        if n_merged_cols < n_header_cols:
            rowvals = rowvals + [np.nan] * (n_header_cols - n_merged_cols)
        elif n_merged_cols > n_header_cols:
            rowvals = rowvals[:n_header_cols]
        rows.append(rowvals)
    merged_data_only = pd.DataFrame(rows, columns=header_df.columns)

    combined_df = pd.concat([header_df, merged_data_only], ignore_index=True)

    # Write to Excel with unmatched sheet if any
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='combined', index=False)
        if unmatched_data:
            pd.DataFrame(unmatched_data).to_excel(writer, sheet_name='unmatched', index=False)

    logging.info(f"Wrote combined and unmatched (if any) to {output_file_path}")

    # # Convert merged data to DataFrame
    # logging.info(f"Data frame:{merged_data}")
    # merged_df = pd.DataFrame(merged_data)
    # logging.info("Converted merged data to DataFrame.")
    # logging.info(f"Unmatched data records: {len(unmatched_data)}")
    # logging.debug(f"Unmatched data samples: {unmatched_data[:5]}")
    # logging.debug(f"Merged DataFrame columns: {merged_df.columns.tolist()}")

    # # Reset index for merged_df to ensure unique indices
    # merged_df = merged_df.reset_index(drop=True)
    # logging.debug("Merged DataFrame index reset.")

    # # Filter columns to include only those with matching names, in the order of merged_df
    # common_columns = [col for col in merged_df.columns if col in pre_header.columns]
    # # Remove columns that are actual NaN or the literal string 'NAN'
    # removed_nan_cols = [c for c in common_columns if (pd.isna(c) or (isinstance(c, str) and c.strip().upper() == 'NAN'))]
    # if removed_nan_cols:
    #     logging.warning(f"Removing NaN-like columns from common_columns: {removed_nan_cols}")
    #     common_columns = [c for c in common_columns if c not in removed_nan_cols]
    # logging.debug(f"Common columns in merged_df order: {common_columns}")

    # # If pre_header has duplicate column labels, reindex will fail. Remove duplicates (keep first).
    # dup_cols = pre_header.columns[pre_header.columns.duplicated()].tolist()
    # if dup_cols:
    #     logging.warning(f"Found duplicate columns in pre_header, dropping duplicates (keep first): {dup_cols}")
    #     pre_header = pre_header.loc[:, ~pre_header.columns.duplicated()]

    # # Reindex pre_header to the common columns (safe - will insert NaN for missing)
    # pre_header = pre_header.reindex(columns=common_columns)

    # # Prepare merged_df keep list, ensure 'MEASURED' is appended if present
    # merged_keep = list(common_columns)
    # if 'MEASURED' in merged_df.columns and 'MEASURED' not in merged_keep:
    #     merged_keep.append('MEASURED')
    # merged_df = merged_df.reindex(columns=merged_keep)

    # # Combine header_data and merged_df
    # logging.info("Combining header data and merged DataFrame.")
    # header_df = pd.DataFrame(header_data)
    # header_df.loc[len(header_df)] = ['head'] * len(header_df.columns)
    # logging.debug(f"Meged_Df: {merged_df}")

    # # Create a replacement row for header_df with merged_df column names
    # n_header_cols = len(header_df.columns)
    # merged_col_names = [str(c) for c in merged_df.columns]

    # # Build a row matching header_df column count: place merged column names left-to-right, pad with empty strings
    # replacement_row = [""] * n_header_cols
    # for i, name in enumerate(merged_col_names):
    #     if i < n_header_cols:
    #         replacement_row[i] = name
    #     else:
    #         break

    # # Assign the replacement row (safe length)
    # header_df.iloc[-1] = replacement_row

    # # Now build merged_data_only as rows matching header_df columns.
    # n_merged_cols = merged_df.shape[1]
    # rows = []
    # for arr in merged_df.values:
    #     rowvals = list(arr)
    #     if n_merged_cols < n_header_cols:
    #         # pad with NaN to match header width
    #         rowvals = rowvals + [np.nan] * (n_header_cols - n_merged_cols)
    #     elif n_merged_cols > n_header_cols:
    #         # truncate extra merged columns to fit into header width
    #         rowvals = rowvals[:n_header_cols]
    #     rows.append(rowvals)

    # merged_data_only = pd.DataFrame(rows, columns=header_df.columns)

    # # Combine header_df and merged_data_only
    # combined_df = pd.concat([header_df, merged_data_only], ignore_index=True)
    # logging.debug(f"Combined DataFrame shape: {combined_df.shape}")

    # # Write combined DataFrame to Excel file. If there are unmatched records, write them to a separate sheet.
    # with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    #     combined_df.to_excel(writer, sheet_name='combined', index=False)
    #     if unmatched_data:
    #         unmatched_df = pd.DataFrame(unmatched_data)
    #         unmatched_df.to_excel(writer, sheet_name='unmatched', index=False)

    # logging.info(f"Combined data written to {output_file_path} (sheets: combined{', unmatched' if unmatched_data else ''})")

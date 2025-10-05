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

def final_data(excel_file_path, txt_file_paths, output_file_path):
    """Merge Excel templates with one or more TXT measurement files.

    txt_file_paths may be a single path (str) or a list of paths. When multiple TXT files
    are provided, measured values are written into columns named MEASURED-1, MEASURED-2, ...
    """
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

    # Accept either a single path or a list of paths
    if isinstance(txt_file_paths, (str, bytes)):
        txt_file_paths = [txt_file_paths]

    # For each TXT file, extract measurements and build a mapping dim->first_measurement
    per_file_maps = []  # list of dicts: [{dim: measurement, ...}, ...]
    for txt_path in txt_file_paths:
        file_meas = extract_measurements(txt_path)
        mmap = {}
        for mes in file_meas:
            if '#' in mes.get('dimension', ''):
                try:
                    dp = mes.get('dimension', '').split('=')[0]
                    d = re.search(r'#(\d+)', dp)
                    if d:
                        dn = int(d.group(1))
                        # keep first measurement for this dimension in this file
                        if dn not in mmap:
                            mmap[dn] = mes
                except Exception:
                    continue
        per_file_maps.append(mmap)

    logging.debug(f"Per-file measurement maps count: {len(per_file_maps)}")

    # Build merged_data by iterating excel_data templates; create MEASURED-N columns for each file
    merged_data = []
    unmatched_data = []

    multi_files = len(per_file_maps) > 1

    for key, template in excel_data.items():
        base = template.copy()
        # For multiple files, add MEASURED-1..N; for single file, use 'MEASURED'
        if multi_files:
            for idx, mmap in enumerate(per_file_maps, start=1):
                mes = mmap.get(key)
                colname = f"MEASURED-{idx}"
                if mes is not None:
                    base[colname] = _try_float(mes.get('measured'))
                else:
                    base[colname] = np.nan
            # keep original DEVIATION/OUT OF TOLERANCE empty (or could compute from first file)
            merged_data.append(base)
        else:
            # single file behavior: populate MEASURED, DEVIATION, OUT OF TOLERANCE if available
            mmap = per_file_maps[0] if per_file_maps else {}
            mes = mmap.get(key)
            if mes is not None:
                if mes.get('+tol') is not None:
                    base['TOLERANCE MAX'] = _try_float(mes.get('+tol'))
                if mes.get('-tol') is not None:
                    base['TOLERANCE MIN'] = _try_float(mes.get('-tol'))
                base['MEASURED'] = _try_float(mes.get('measured'))
                base['DEVIATION'] = _try_float(mes.get('deviation'))
                base['OUT OF TOLERANCE'] = _try_float(mes.get('outtol'))
            else:
                # ensure MEASURED exists
                base.setdefault('MEASURED', '')
            merged_data.append(base)

    # Any measurement keys not present in excel_data are unmatched
    all_keys_in_files = set().union(*[set(m.keys()) for m in per_file_maps]) if per_file_maps else set()
    unmatched_keys = all_keys_in_files - set(excel_data.keys())
    for uk in unmatched_keys:
        for mmap in per_file_maps:
            mes = mmap.get(uk)
            if mes:
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
    # Remove IDENTIFICATION NO from common_columns (case-insensitive)
    id_cols = [c for c in common_columns if isinstance(c, str) and c.strip().upper() == 'IDENTIFICATION NO']
    if id_cols:
        logging.info(f"Removing identification columns from common_columns: {id_cols}")
        common_columns = [c for c in common_columns if c not in id_cols]

    # Deduplicate pre_header columns if needed
    if pre_header.columns.duplicated().any():
        dup_cols = pre_header.columns[pre_header.columns.duplicated()].tolist()
        logging.warning(f"Found duplicate columns in pre_header, dropping duplicates (keep first): {dup_cols}")
        pre_header = pre_header.loc[:, ~pre_header.columns.duplicated()]

    pre_header = pre_header.reindex(columns=common_columns)

    # Ensure MEASURED preserved in merged_df
    merged_keep = list(common_columns)
    logging.debug(f"Common columns in merged_df order: {common_columns}")
    # Always preserve any measured-like columns (MEASURED, MEASURED-1, MEASURED-2, ...)
    measured_cols = [c for c in merged_df.columns if isinstance(c, str) and c.upper().startswith('MEASURED')]
    for mc in measured_cols:
        if mc not in merged_keep:
            merged_keep.append(mc)
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

import pandas as pd
import numpy as np
import sys
import re
import logging
import os
import openpyxl
from typing import Dict, List, Union
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
try:
    from openpyxl.drawing.image import Image
except ImportError:
    try:
        from openpyxl.drawing import Image
    except ImportError:
        Image = None
sys.path.append("../")  # Add parent directory to sys.path for relative imports
from api.utils.excel_extraction import extract_excel_data, copy_cell_format
from api.utils.extract_measurements import extract_measurements

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
COLUMN_TO_RMV = ['OUT OF TOLERANCE', 'DEVIATION', 'OUT_OF_TOLERANCE','IDENTIFICATION NO']  # replace with your list

def merge_excel_with_header(output_file_path, header_file_path, final_output_path, header_row_idx):
    """
    Append data workbook into header workbook starting after header_row_idx,
    preserving header formatting from header_file_path.
    
    Args:
        output_file_path: Path to the data Excel file
        header_file_path: Path to the header template Excel file
        final_output_path: Path where the final merged file will be saved
        header_row_idx: Row index where the header ends (1-based)
    """
    logging.info("Starting Excel merge - appending data after header with format preservation")
    logging.info(f"Data file: {output_file_path}, Header file: {header_file_path}, header_row_idx: {header_row_idx}")

    try:
        # Load workbooks
        header_wb = openpyxl.load_workbook(header_file_path)
        data_wb = openpyxl.load_workbook(output_file_path)

        # Choose sheets (first sheet for header, 'combined' or first for data)
        header_sheet = header_wb[header_wb.sheetnames[0]]
        data_sheet_name = 'combined' if 'combined' in data_wb.sheetnames else data_wb.sheetnames[0]
        data_sheet = data_wb[data_sheet_name]

        logging.info(f"Using data from sheet: {data_sheet_name}")
        logging.info(f"Using header format from sheet: {header_wb.sheetnames[0]}")

        # Determine column range to handle
        max_cols = max(header_sheet.max_column, data_sheet.max_column)

        # Set column widths (prefer header widths, fallback to data widths)
        for col in range(1, max_cols + 1):
            col_letter = get_column_letter(col)
            if col_letter in header_sheet.column_dimensions and header_sheet.column_dimensions[col_letter].width:
                width = header_sheet.column_dimensions[col_letter].width
            elif col_letter in data_sheet.column_dimensions and data_sheet.column_dimensions[col_letter].width:
                width = data_sheet.column_dimensions[col_letter].width
            else:
                width = 10  # Default width
            header_sheet.column_dimensions[col_letter].width = width

        # Insert empty rows after header_row_idx to accommodate all data rows (skip header row in data)
        data_rows = data_sheet.max_row - 1  # Exclude header row
        if data_rows > 0:
            insert_at = header_row_idx + 1
            header_sheet.insert_rows(insert_at, amount=data_rows)
            logging.info(f"Inserted {data_rows} rows starting at row {insert_at}")

            # Copy data rows directly by position
            for data_row in range(2, data_sheet.max_row + 1):  # Start from row 2 (skip header)
                target_row = insert_at + (data_row - 2)
                
                # Copy all columns by position
                for col in range(1, max_cols + 1):
                    source_cell = data_sheet.cell(row=data_row, column=col) if col <= data_sheet.max_column else None
                    target_cell = header_sheet.cell(row=target_row, column=col)
                    
                    # Extract ONLY the raw value (not the cell object)
                    cell_value = source_cell.value if source_cell is not None else None
                    
                    # Set the value first
                    target_cell.value = cell_value
                    
                    # Always set black font FIRST (before any other formatting)
                    target_cell.font = Font(
                        name='Calibri',
                        size=11,
                        bold=False,
                        italic=False,
                        underline=None,
                        strike=False,
                        color="FF000000"  # BLACK text - explicit with alpha
                    )
                    
                    # Then apply other formatting from template (if exists)
                    template_cell = header_sheet.cell(row=header_row_idx, column=col) if col <= header_sheet.max_column else None
                    if template_cell is not None and template_cell.has_style:
                        try:
                            # Copy non-font formatting from template
                            if template_cell.border:
                                target_cell.border = template_cell.border
                            if template_cell.fill and template_cell.fill.fill_type:
                                target_cell.fill = template_cell.fill
                            if template_cell.alignment:
                                #target_cell.alignment = template_cell.alignment
                                  target_cell.alignment = Alignment(
                                    horizontal=source_cell.alignment.horizontal,
                                    vertical=source_cell.alignment.vertical,
                                    text_rotation=source_cell.alignment.text_rotation,
                                    wrap_text=True,
                                    shrink_to_fit=source_cell.alignment.shrink_to_fit,
                                    indent=source_cell.alignment.indent
                                )
                            
                            if template_cell.number_format:
                                target_cell.number_format = template_cell.number_format
                        except Exception as e:
                            logging.debug(f"Failed to copy template style for row {target_row} col {col}: {str(e)}")

                # Set row height
                if target_row not in header_sheet.row_dimensions:
                    header_sheet.row_dimensions[target_row].height = 15  # Default height

            # Second pass: Force all data cells to have BLACK font color
            logging.info("Second pass: Ensuring all data cells have black font color")
            for target_row in range(insert_at, insert_at + data_rows):
                for col in range(1, max_cols + 1):
                    cell = header_sheet.cell(row=target_row, column=col)
                    if cell.value is not None:  # Only process cells with values
                        try:
                            # Get current font properties and override color to black
                            current_font = cell.font
                            cell.font = Font(
                                name=current_font.name if current_font.name else 'Calibri',
                                size=current_font.size if current_font.size else 11,
                                bold=False,
                                italic=False,
                                underline=None,
                                strike=False,
                                color="FF000000"  # Explicit black with alpha channel
                            )
                        except Exception as e:
                            logging.debug(f"Failed to set black font for row {target_row} col {col}: {str(e)}")

        # Save the final merged workbook
        header_wb.save(final_output_path)
        logging.info(f"Successfully saved merged file to {final_output_path}")

    except Exception as e:
        logging.error(f"Error during Excel merge: {str(e)}", exc_info=True)
        raise


def move_measured_columns_to_end(df):
    """
    Reorder DataFrame columns so all columns starting with 'MEASURED' (case-insensitive)
    appear at the end, preserving the order of other columns.
    Returns a new DataFrame with reordered columns.
    """
    cols = df.columns.tolist()
    measured_cols = [col for col in cols if str(col).strip().upper().startswith("MEASURED")]
    other_cols = [col for col in cols if not str(col).strip().upper().startswith("MEASURED")]
    return df[other_cols + measured_cols]

def _get_merged_cell_value(sheet, row, col):
    """Get cell value, handling merged cells properly."""
    cell = sheet.cell(row=row, column=col)
    for range_ in sheet.merged_cells.ranges:
        if cell.coordinate in range_:
            # Return value from the top-left cell of the merged range
            return sheet.cell(row=range_.min_row, column=range_.min_col).value
    return cell.value

def get_data_sheet_columns(sheet, header_row=1):
    """Get column headers from sheet, handling merged cells and stripping whitespace."""
    max_col = sheet.max_column
    headers = []
    for col in range(1, max_col + 1):
        value = _get_merged_cell_value(sheet, header_row, col)
        if value is not None:
            # Strip whitespace and trailing dots
            value = str(value).strip().rstrip('.')
        headers.append(value)
    return headers

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

    txt_file_paths may be:
    - a single path (str)
    - a list of paths
    - a dict organized by folders: {"folder_name": [list_of_paths]}
    
    When dict is provided, files are processed folder-wise while preserving folder context.
    When multiple TXT files are provided, measured values are written into columns named MEASURED-1, MEASURED-2, ...
    """
    logging.info("Starting data merging process.")

    # Extract data from Excel file
    excel_data, header_file_path,header_row_idx = extract_excel_data(excel_file_path)
    logging.debug(f"excel_Data keys: {list(excel_data)}")

    # Normalize keys inside excel_data templates to uppercase so they match pre_header columns
    for k, v in list(excel_data.items()):
        excel_data[k] = { (col.upper() if isinstance(col, str) else col): val for col, val in v.items() }

    #logging.debug(f"Normalized excel_Data keys: {excel_data.values()}")
    # # Normalize pre_header column names to uppercase for matching
    # pre_header.columns = [col.upper() for col in pre_header.columns]
    # pre_header = pre_header.reset_index(drop=True)

    # Handle different input formats and preserve folder context
    folder_to_files = {}
    
    if isinstance(txt_file_paths, dict):
        # Already organized by folders - preserve structure
        folder_to_files = txt_file_paths
        logging.info(f"Processing files organized by folders: {list(folder_to_files.keys())}")
        for folder_name, files in folder_to_files.items():
            logging.info(f"  Folder '{folder_name}': {len(files)} file(s)")
    elif isinstance(txt_file_paths, (str, bytes)):
        # Single file path
        folder_to_files = {"default": [txt_file_paths]}
    elif isinstance(txt_file_paths, list):
        # List of paths - organize by folder from file path
        folder_to_files = {}
        for path in txt_file_paths:
            # Extract folder name from path (e.g., /uploads/1/file.txt -> "1")
            path_parts = path.split(os.sep)
            if len(path_parts) > 1 and path_parts[-2] not in ["uploads", ""]:
                folder_name = path_parts[-2]
            else:
                folder_name = "default"
            
            if folder_name not in folder_to_files:
                folder_to_files[folder_name] = []
            folder_to_files[folder_name].append(path)
        logging.info(f"Organized files by folder: {list(folder_to_files.keys())}")
    else:
        raise ValueError(f"Unsupported txt_file_paths type: {type(txt_file_paths)}")
    
    # Flatten into a single list for processing while keeping folder info
    txt_file_paths_flat = []
    folder_info = {}  # Maps file path to folder name
    for folder_name, paths in folder_to_files.items():
        for path in paths:
            txt_file_paths_flat.append(path)
            folder_info[path] = folder_name
    
    logging.info(f"Total files to process: {len(txt_file_paths_flat)} across {len(folder_to_files)} folder(s)")

    # For each TXT file, extract measurements and build a mapping dim->first_measurement
    # Group files by folder first, then merge within each folder
    files_by_folder = {}
    for txt_path in txt_file_paths_flat:
        folder_name = folder_info.get(txt_path, "default")
        if folder_name not in files_by_folder:
            files_by_folder[folder_name] = []
        files_by_folder[folder_name].append(txt_path)
    
    logging.info(f"Grouped files by folder: {list(files_by_folder.keys())}")
    
    # Process each folder and merge files within it
    per_file_maps = []  # list of dicts: [{dim: measurement, ...}, ...] - one merged dict per folder
    
    for folder_name, file_paths in files_by_folder.items():
        logging.info(f"\n=== Processing Folder '{folder_name}' with {len(file_paths)} file(s) ===")
        merged_mmap = {}  # Merged measurements for this folder
        
        for file_index, txt_path in enumerate(file_paths, start=1):
            logging.info(f"Processing file {file_index}/{len(file_paths)} from folder '{folder_name}': {os.path.basename(txt_path)}")
            file_meas = extract_measurements(txt_path)
            logging.info(f"  Extracted {len(file_meas)} measurements from this file")
            
            for mes in file_meas:
                if '#' in mes.get('dimension', ''):
                    try:
                        dp = mes.get('dimension', '').split('=')[0]
                        d = re.search(r'#(\d+)', dp)
                        if d:
                            dn = int(d.group(1))
                            # Only add if not already in merged_mmap (keep first file's data)
                            if dn not in merged_mmap:
                                merged_mmap[dn] = mes
                                logging.debug(f"  ✓ Added measurement #{dn} from file {file_index}")
                            else:
                                logging.debug(f"  ⊘ Skipped duplicate measurement #{dn} (keeping data from file 1)")
                    except Exception as e:
                        logging.debug(f"  Error processing measurement: {str(e)}")
                        continue
        
        if merged_mmap:
            per_file_maps.append(merged_mmap)
            logging.info(f"Folder '{folder_name}': ✓ Merged {len(file_paths)} file(s) → {len(merged_mmap)} unique measurements\n")
        else:
            logging.warning(f"Folder '{folder_name}': ⚠ No valid measurements found\n")

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
                
                base['DEVIATION'] = _try_float(mes.get('deviation'))
                base['OUT OF TOLERANCE'] = _try_float(mes.get('outtol'))
                base['MEASURED'] = _try_float(mes.get('measured'))
            else:
                # ensure MEASURED exists
                base.setdefault('MEASURED', '')
            merged_data.append(base)

    # Any measurement keys not present in excel_data are unmatched
    logging.debug(f"MEger_data count before unmatched: {merged_data}")
    all_keys_in_files = set().union(*[set(m.keys()) for m in per_file_maps]) if per_file_maps else set()
    logging.info(f"All keys in measurement files: {all_keys_in_files}")
    unmatched_keys = all_keys_in_files - set(excel_data.keys())
    logging.info(f"Unmatched keys count: {unmatched_keys}")
    for uk in unmatched_keys:
        for mmap in per_file_maps:
            mes = mmap.get(uk)
            if mes:
                unmatched_record = {
                    'DIMENSION_NUMBER': uk,
                    'DIMENSION': mes.get('dimension'),
                    'TOLERANCE_MAX': mes.get('+tol'),
                    'TOLERANCE_MIN': mes.get('-tol'),
                    'DEVIATION': mes.get('deviation'),
                    'OUT_OF_TOLERANCE': mes.get('outtol'),
                    'MEASURED': mes.get('measured')
                }
                unmatched_data.append(unmatched_record)

    logging.debug(f"unmatxhed_data count: {unmatched_data}")
    logging.info(f"Built merged_data rows: {len(merged_data)}; unmatched: {len(unmatched_data)}")

    # Convert merged_data to DataFrame and drop columns that are all NaN
    merged_df = pd.DataFrame(merged_data)
    merged_df = merged_df.dropna(axis=1, how='all')
    logging.debug(f"Merged DataFrame columns before dropping specified columns: {merged_df.columns.tolist()}")
    cols_to_rmv = [c for c in COLUMN_TO_RMV if c in merged_df.columns]
    logging.debug(f"Columns to be removed: {cols_to_rmv}")
    if cols_to_rmv:
       merged_df.drop(columns=cols_to_rmv, inplace=True)
    merged_df = move_measured_columns_to_end(merged_df)
    logging.debug(f"Merged DataFrame columns after dropping all-NaN and reordering: {merged_df.columns.tolist()}")
    # Save the data to a temporary Excel file first
    # temp_output = output_file_path
    temp_output = output_file_path
    logging.info(f"Writing to temporary file: {temp_output}")
    logging.debug(f"Merged DataFrame preview:\n{merged_df.head()}")
    
    temp_data_file = "data.xlsx"
    
    # Write DataFrame to Excel using ExcelWriter to control formatting
    with pd.ExcelWriter(temp_data_file, engine='openpyxl') as writer:
        merged_df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    logging.info(f"Data written to temporary file: {temp_data_file}")
    logging.info(f"header_file_path: {header_file_path}, header_row_idx: {header_row_idx}")
    logging.info("Data merging process completed successfully.")

    # # Merge the temporary file with the header file while preserving formatting
    # try:
    #     logging.info("Merging temporary file with header file to preserve formatting.")
    #     merge_excel_with_header(temp_data_file, header_file_path, temp_output, header_row_idx)
    #     logging.info(f"Final formatted data saved to {temp_output}")
    
    # except Exception as e:
    #     logging.error(f"Failed to create excel file: {str(e)}")
    #     raise
    
    # finally:
    #     # Clean up temporary files
    #     try:
    #         if os.path.exists(temp_data_file):
    #             os.remove(temp_data_file)
    #             logging.info(f"Cleaned up temporary file: {temp_data_file}")
    #         if os.path.exists("temp_modified.xlsx"):
    #             os.remove("temp_modified.xlsx")
    #             logging.info("Cleaned up temporary header file: temp_modified.xlsx")
    #     except Exception as e:
    #         logging.warning(f"Failed to clean up temporary files: {str(e)}")
    # Align all data cells in the center
    logging.info("Aligning all data cells in the center...")
    try:
        data_workbook = openpyxl.load_workbook(temp_data_file)
        data_worksheet = data_workbook.active

        # Iterate through ALL cells and set alignment to center
        for row in data_worksheet.iter_rows(min_row=1, max_row=data_worksheet.max_row, 
                                           min_col=1, max_col=data_worksheet.max_column):
            for cell in row:
                if cell.value is not None:
                    # Force center alignment for every cell with a value
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        data_workbook.save(temp_data_file)
        logging.info("Successfully aligned all data cells to center")
    except Exception as e:
        logging.warning(f"Failed to align data cells to center: {str(e)}")

    append_excel_data(temp_data_file, header_file_path, header_row_idx, output_file_path)

def append_excel_data(temp_data_file, header_file_path, header_row_idx, output_file_path=None):
    """
    Create a new Excel file that merges header data and temp data.
    Preserves formatting from header file while copying data values.
    
    Args:
        temp_data_file: Path to the Excel file containing data to merge
        header_file_path: Path to the Excel file with header data
        header_row_idx: Row index where header data ends (1-based)
        output_file_path: Path for the new merged Excel file (optional, defaults to temp_data_file)
    """
    if output_file_path is None:
        output_file_path = temp_data_file
    
    logging.info(f"Creating new merged Excel file: {output_file_path}")
    logging.info(f"Merging header from {header_file_path} (up to row {header_row_idx}) with data from {temp_data_file}")

    try:
        # Load both workbooks
        header_wb = openpyxl.load_workbook(header_file_path)
        temp_wb = openpyxl.load_workbook(temp_data_file)

        # Get the active sheets
        header_sheet = header_wb.active
        temp_sheet = temp_wb.active

        logging.info(f"Header sheet: {header_sheet.title}, Temp sheet: {temp_sheet.title}")
        
        # Create new workbook for merged result
        merged_wb = openpyxl.Workbook()
        merged_sheet = merged_wb.active
        merged_sheet.title = "Merged_Data"

        # First, copy header data with formatting (from row 1 to header_row_idx)
        header_max_cols = header_sheet.max_column
        logging.info(f"Copying header data with formatting: rows 1 to {header_row_idx}, columns 1 to {header_max_cols}")
        
        # Copy column dimensions from header sheet
        for col_letter, dimension in header_sheet.column_dimensions.items():
            merged_sheet.column_dimensions[col_letter].width = dimension.width
        
        # Copy row dimensions and cell data with formatting
        for row in range(1, header_row_idx + 1):
            # Copy row height if it exists
            if row in header_sheet.row_dimensions:
                merged_sheet.row_dimensions[row].height = header_sheet.row_dimensions[row].height
                
            for col in range(1, header_max_cols + 1):
                source_cell = header_sheet.cell(row=row, column=col)
                target_cell = merged_sheet.cell(row=row, column=col)
                
                # Copy value
                target_cell.value = source_cell.value
                
                # Copy formatting if it exists
                if source_cell.has_style:
                    try:
                        # Explicitly copy font with all properties including color
                        if source_cell.font:
                            target_cell.font = Font(
                                name=source_cell.font.name,
                                size=source_cell.font.size,
                                bold=source_cell.font.bold,
                                italic=source_cell.font.italic,
                                underline=source_cell.font.underline,
                                strike=source_cell.font.strike,
                                color=source_cell.font.color
                            )
                        target_cell.fill = source_cell.fill
                        target_cell.border = source_cell.border
                        
                        #target_cell.alignment = source_cell.alignment
                        target_cell.number_format = source_cell.number_format
                        src_align = source_cell.alignment

                        target_cell.alignment = Alignment(
                            horizontal=src_align.horizontal,
                            vertical=src_align.vertical,
                            text_rotation=src_align.text_rotation,
                            wrap_text=True,  # Set wrap text to True
                            shrink_to_fit=src_align.shrink_to_fit,
                            indent=src_align.indent
                        )
                        
                    except Exception as e:
                        logging.debug(f"Could not copy formatting for header row {row}, col {col}: {str(e)}")

        # Copy images from header sheet
        if Image is not None:
            try:
                if hasattr(header_sheet, '_images') and header_sheet._images:
                    logging.info(f"Found {len(header_sheet._images)} images in header sheet")
                    for image in header_sheet._images:
                        try:
                            # Create a new image object with the same properties
                            new_image = Image(image.ref)
                            new_image.anchor = image.anchor
                            if hasattr(image, 'width'):
                                new_image.width = image.width
                            if hasattr(image, 'height'):
                                new_image.height = image.height
                            
                            # Add the image to the merged sheet
                            merged_sheet.add_image(new_image)
                            logging.info(f"Successfully copied image at anchor: {image.anchor}")
                        except Exception as img_error:
                            logging.warning(f"Could not copy individual image: {str(img_error)}")
                else:
                    logging.info("No images found in header sheet")
            except Exception as e:
                logging.warning(f"Could not access images from header sheet: {str(e)}")
        else:
            logging.warning("Image support not available - images will not be copied")

        # Then, copy ALL temp data (including all rows from temp file)
        temp_data_start_row = 1  # Copy all rows including the first row
        temp_data_end_row = temp_sheet.max_row
        temp_max_cols = temp_sheet.max_column
        
        total_data_rows = temp_data_end_row - temp_data_start_row + 1
        logging.info(f"Copying ALL temp data: {total_data_rows} rows (from row {temp_data_start_row} to {temp_data_end_row}), columns 1 to {temp_max_cols}")

        if total_data_rows > 0:
            # Start copying temp data after header_row_idx
            merged_start_row = header_row_idx + 1
            
            for temp_row in range(temp_data_start_row, temp_data_end_row + 1):
                # Calculate target row in merged sheet
                target_row = merged_start_row + (temp_row - temp_data_start_row)
                
                # Set default row height for data rows
                #merged_sheet.row_dimensions[target_row].height = 15
                merged_sheet.row_dimensions[target_row].height = 40
                
                # Copy all columns from this row with header formatting applied to data
                for col in range(1, temp_max_cols + 1):
                    source_cell = temp_sheet.cell(row=temp_row, column=col)
                    target_cell = merged_sheet.cell(row=target_row, column=col)
                    
                    # Copy value from temp data
                    target_cell.value = source_cell.value
                    
                    # ALWAYS apply center alignment to ALL data cells with values
                    if target_cell.value is not None:
                        target_cell.alignment = Alignment(horizontal="center", vertical="center",wrap_text=True)


                    
                    # Apply formatting from corresponding header column (use header_row_idx as template)
                    if col <= header_max_cols:
                        header_template_cell = header_sheet.cell(row=header_row_idx, column=col)
                        if header_template_cell.has_style:
                            try:
                                # Explicitly copy font with all properties including color
                                if header_template_cell.font:
                                    target_cell.font = Font(
                                        name=header_template_cell.font.name,
                                        size=header_template_cell.font.size,
                                        bold=header_template_cell.font.bold,
                                        italic=header_template_cell.font.italic,
                                        underline=header_template_cell.font.underline,
                                        strike=header_template_cell.font.strike,
                                        color=header_template_cell.font.color
                                    )
                                target_cell.fill = header_template_cell.fill
                                target_cell.border = header_template_cell.border
                                target_cell.number_format = header_template_cell.number_format
                                
                            except Exception as e:
                                logging.debug(f"Could not copy header formatting to data row {target_row}, col {col}: {str(e)}")

        # Add borders to all cells for print formatting
        logging.info("Adding borders to all cells for print formatting...")
        try:
            # Create a thin border style
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # Apply borders to all cells that have content
            max_row_with_data = merged_sheet.max_row
            max_col_with_data = merged_sheet.max_column
            
            for row in range(1, max_row_with_data + 1):
                for col in range(1, max_col_with_data + 1):
                    cell = merged_sheet.cell(row=row, column=col)
                    #if cell.value is not None or row <= header_row_idx:  # Apply to header rows and data cells
                    # Preserve existing formatting while adding borders
                    current_font = cell.font
                    current_fill = cell.fill
                    current_alignment = cell.alignment
                    current_number_format = cell.number_format
                    
                    # Apply border while keeping other formatting
                    cell.border = thin_border
                        
            logging.info(f"Successfully added borders to {max_row_with_data} rows x {max_col_with_data} columns")
            
        except Exception as e:
            logging.warning(f"Could not add borders: {str(e)}")

        # Save the new merged file
        merged_wb.save(output_file_path)
        logging.info(f"Successfully created merged Excel file with borders: {output_file_path}")
        logging.info(f"Total rows in merged file: {merged_sheet.max_row}")

    except Exception as e:
        logging.error(f"Error during Excel merge: {str(e)}", exc_info=True)
        raise
    finally:
        # Close workbooks to free memory
        if 'header_wb' in locals():
            header_wb.close()
        if 'temp_wb' in locals():
            temp_wb.close()
        if 'merged_wb' in locals():
            merged_wb.close()




# import pandas as pd
# import numpy as np
# import sys
# import re
# import logging
# import os
# import openpyxl
# from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
# from openpyxl.utils import get_column_letter
# try:
#     from openpyxl.drawing.image import Image
# except ImportError:
#     try:
#         from openpyxl.drawing import Image
#     except ImportError:
#         Image = None
# sys.path.append("../")  # Add parent directory to sys.path for relative imports
# from api.utils.excel_extraction import extract_excel_data, copy_cell_format
# from api.utils.extract_measurements import extract_measurements

# # Configure logging
# logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
# COLUMN_TO_RMV = ['OUT OF TOLERANCE', 'DEVIATION', 'OUT_OF_TOLERANCE','IDENTIFICATION NO']  # replace with your list

# def merge_excel_with_header(output_file_path, header_file_path, final_output_path, header_row_idx):
#     """
#     Append data workbook into header workbook starting after header_row_idx,
#     preserving header formatting from header_file_path.
    
#     Args:
#         output_file_path: Path to the data Excel file
#         header_file_path: Path to the header template Excel file
#         final_output_path: Path where the final merged file will be saved
#         header_row_idx: Row index where the header ends (1-based)
#     """
#     logging.info("Starting Excel merge - appending data after header with format preservation")
#     logging.info(f"Data file: {output_file_path}, Header file: {header_file_path}, header_row_idx: {header_row_idx}")

#     try:
#         # Load workbooks
#         header_wb = openpyxl.load_workbook(header_file_path)
#         data_wb = openpyxl.load_workbook(output_file_path)

#         # Choose sheets (first sheet for header, 'combined' or first for data)
#         header_sheet = header_wb[header_wb.sheetnames[0]]
#         data_sheet_name = 'combined' if 'combined' in data_wb.sheetnames else data_wb.sheetnames[0]
#         data_sheet = data_wb[data_sheet_name]

#         logging.info(f"Using data from sheet: {data_sheet_name}")
#         logging.info(f"Using header format from sheet: {header_wb.sheetnames[0]}")

#         # Determine column range to handle
#         max_cols = max(header_sheet.max_column, data_sheet.max_column)

#         # Set column widths (prefer header widths, fallback to data widths)
#         for col in range(1, max_cols + 1):
#             col_letter = get_column_letter(col)
#             if col_letter in header_sheet.column_dimensions and header_sheet.column_dimensions[col_letter].width:
#                 width = header_sheet.column_dimensions[col_letter].width
#             elif col_letter in data_sheet.column_dimensions and data_sheet.column_dimensions[col_letter].width:
#                 width = data_sheet.column_dimensions[col_letter].width
#             else:
#                 width = 10  # Default width
#             header_sheet.column_dimensions[col_letter].width = width

#         # Insert empty rows after header_row_idx to accommodate all data rows (skip header row in data)
#         data_rows = data_sheet.max_row - 1  # Exclude header row
#         if data_rows > 0:
#             insert_at = header_row_idx + 1
#             header_sheet.insert_rows(insert_at, amount=data_rows)
#             logging.info(f"Inserted {data_rows} rows starting at row {insert_at}")

#             # Copy data rows directly by position
#             for data_row in range(2, data_sheet.max_row + 1):  # Start from row 2 (skip header)
#                 target_row = insert_at + (data_row - 2)
                
#                 # Copy all columns by position
#                 for col in range(1, max_cols + 1):
#                     source_cell = data_sheet.cell(row=data_row, column=col) if col <= data_sheet.max_column else None
#                     target_cell = header_sheet.cell(row=target_row, column=col)
                    
#                     # Extract ONLY the raw value (not the cell object)
#                     cell_value = source_cell.value if source_cell is not None else None
                    
#                     # Set the value first
#                     target_cell.value = cell_value
                    
#                     # Always set black font FIRST (before any other formatting)
#                     target_cell.font = Font(
#                         name='Calibri',
#                         size=11,
#                         bold=False,
#                         italic=False,
#                         underline=None,
#                         strike=False,
#                         color="FF000000"  # BLACK text - explicit with alpha
#                     )
                    
#                     # Then apply other formatting from template (if exists)
#                     template_cell = header_sheet.cell(row=header_row_idx, column=col) if col <= header_sheet.max_column else None
#                     if template_cell is not None and template_cell.has_style:
#                         try:
#                             # Copy non-font formatting from template
#                             if template_cell.border:
#                                 target_cell.border = template_cell.border
#                             if template_cell.fill and template_cell.fill.fill_type:
#                                 target_cell.fill = template_cell.fill
#                             if template_cell.alignment:
#                                 #target_cell.alignment = template_cell.alignment
#                                   target_cell.alignment = Alignment(
#                                     horizontal=source_cell.alignment.horizontal,
#                                     vertical=source_cell.alignment.vertical,
#                                     text_rotation=source_cell.alignment.text_rotation,
#                                     wrap_text=True,
#                                     shrink_to_fit=source_cell.alignment.shrink_to_fit,
#                                     indent=source_cell.alignment.indent
#                                 )
                            
#                             if template_cell.number_format:
#                                 target_cell.number_format = template_cell.number_format
#                         except Exception as e:
#                             logging.debug(f"Failed to copy template style for row {target_row} col {col}: {str(e)}")

#                 # Set row height
#                 if target_row not in header_sheet.row_dimensions:
#                     header_sheet.row_dimensions[target_row].height = 15  # Default height

#             # Second pass: Force all data cells to have BLACK font color
#             logging.info("Second pass: Ensuring all data cells have black font color")
#             for target_row in range(insert_at, insert_at + data_rows):
#                 for col in range(1, max_cols + 1):
#                     cell = header_sheet.cell(row=target_row, column=col)
#                     if cell.value is not None:  # Only process cells with values
#                         try:
#                             # Get current font properties and override color to black
#                             current_font = cell.font
#                             cell.font = Font(
#                                 name=current_font.name if current_font.name else 'Calibri',
#                                 size=current_font.size if current_font.size else 11,
#                                 bold=False,
#                                 italic=False,
#                                 underline=None,
#                                 strike=False,
#                                 color="FF000000"  # Explicit black with alpha channel
#                             )
#                         except Exception as e:
#                             logging.debug(f"Failed to set black font for row {target_row} col {col}: {str(e)}")

#         # Save the final merged workbook
#         header_wb.save(final_output_path)
#         logging.info(f"Successfully saved merged file to {final_output_path}")

#     except Exception as e:
#         logging.error(f"Error during Excel merge: {str(e)}", exc_info=True)
#         raise


# def move_measured_columns_to_end(df):
#     """
#     Reorder DataFrame columns so all columns starting with 'MEASURED' (case-insensitive)
#     appear at the end, preserving the order of other columns.
#     Returns a new DataFrame with reordered columns.
#     """
#     cols = df.columns.tolist()
#     measured_cols = [col for col in cols if str(col).strip().upper().startswith("MEASURED")]
#     other_cols = [col for col in cols if not str(col).strip().upper().startswith("MEASURED")]
#     return df[other_cols + measured_cols]

# def _get_merged_cell_value(sheet, row, col):
#     """Get cell value, handling merged cells properly."""
#     cell = sheet.cell(row=row, column=col)
#     for range_ in sheet.merged_cells.ranges:
#         if cell.coordinate in range_:
#             # Return value from the top-left cell of the merged range
#             return sheet.cell(row=range_.min_row, column=range_.min_col).value
#     return cell.value

# def get_data_sheet_columns(sheet, header_row=1):
#     """Get column headers from sheet, handling merged cells and stripping whitespace."""
#     max_col = sheet.max_column
#     headers = []
#     for col in range(1, max_col + 1):
#         value = _get_merged_cell_value(sheet, header_row, col)
#         if value is not None:
#             # Strip whitespace and trailing dots
#             value = str(value).strip().rstrip('.')
#         headers.append(value)
#     return headers

# def _try_float(v):
#     """Safely convert v to float; return np.nan on failure."""
#     try:
#         if v is None:
#             return np.nan
#         return float(v)
#     except Exception:
#         return np.nan

# def final_data(excel_file_path, txt_file_paths, output_file_path):
#     """Merge Excel templates with one or more TXT measurement files.

#     txt_file_paths may be a single path (str) or a list of paths. When multiple TXT files
#     are provided, measured values are written into columns named MEASURED
#     """
#     logging.info("Starting data merging process.")

#     # Extract data from Excel file
#     excel_data, header_file_path,header_row_idx = extract_excel_data(excel_file_path)
#     logging.debug(f"excel_Data keys: {list(excel_data)}")

#     # Normalize keys inside excel_data templates to uppercase so they match pre_header columns
#     for k, v in list(excel_data.items()):
#         excel_data[k] = { (col.upper() if isinstance(col, str) else col): val for col, val in v.items() }

#     #logging.debug(f"Normalized excel_Data keys: {excel_data.values()}")
#     # # Normalize pre_header column names to uppercase for matching
#     # pre_header.columns = [col.upper() for col in pre_header.columns]
#     # pre_header = pre_header.reset_index(drop=True)

#     # Accept either a single path or a list of paths
#     if isinstance(txt_file_paths, (str, bytes)):
#         txt_file_paths = [txt_file_paths]

#     # For each TXT file, extract measurements and build a mapping dim->first_measurement
#     per_file_maps = []  # list of dicts: [{dim: measurement, ...}, ...]
#     for txt_path in txt_file_paths:
#         file_meas = extract_measurements(txt_path)
#         mmap = {}
#         for mes in file_meas:
#             if '#' in mes.get('dimension', ''):
#                 try:
#                     dp = mes.get('dimension', '').split('=')[0]
#                     d = re.search(r'#(\d+)', dp)
#                     if d:
#                         dn = int(d.group(1))
#                         # keep first measurement for this dimension in this file
#                         if dn not in mmap:
#                             mmap[dn] = mes
#                 except Exception:
#                     continue
#         per_file_maps.append(mmap)

#     logging.debug(f"Per-file measurement maps count: {len(per_file_maps)}")

#     # Build merged_data by iterating excel_data templates; create MEASURED-N columns for each file
#     merged_data = []
#     unmatched_data = []

#     multi_files = len(per_file_maps) > 1


#     for key, template in excel_data.items():
#         base = template.copy()

#         nominal = _try_float(base.get('NOMINAL'))
#         best_value = np.nan
#         smallest_deviation = float('inf')

#         for mmap in per_file_maps:
#             mes = mmap.get(key)
#             if mes is not None:
#                 measured = _try_float(mes.get('measured'))

#                 if not np.isnan(measured):
#                     # If nominal exists → compare deviation
#                     if not np.isnan(nominal):
#                         deviation = abs(measured - nominal)

#                         if deviation < smallest_deviation:
#                             smallest_deviation = deviation
#                             best_value = measured
#                     else:
#                         # If no nominal → just take first measured
#                         best_value = measured
#                         break

#         base['MEASURED'] = best_value
#         merged_data.append(base)
#     # for key, template in excel_data.items():
#     #     base = template.copy()

#     #     measured_value = np.nan

#     #     # Loop through all uploaded TXT files
#     #     for mmap in per_file_maps:
#     #         mes = mmap.get(key)
#     #         if mes is not None:
#     #             # Take the first found measured value
#     #             measured_value = _try_float(mes.get('measured'))
#     #             break   # stop after first match

#     #     base['MEASURED'] = measured_value
#     #     merged_data.append(base)

#     # Any measurement keys not present in excel_data are unmatched
#     logging.debug(f"MEger_data count before unmatched: {merged_data}")
#     all_keys_in_files = set().union(*[set(m.keys()) for m in per_file_maps]) if per_file_maps else set()
#     logging.info(f"All keys in measurement files: {all_keys_in_files}")
#     unmatched_keys = all_keys_in_files - set(excel_data.keys())
#     logging.info(f"Unmatched keys count: {unmatched_keys}")
#     for uk in unmatched_keys:
#         for mmap in per_file_maps:
#             mes = mmap.get(uk)
#             if mes:
#                 unmatched_record = {
#                     'DIMENSION_NUMBER': uk,
#                     'DIMENSION': mes.get('dimension'),
#                     'TOLERANCE_MAX': mes.get('+tol'),
#                     'TOLERANCE_MIN': mes.get('-tol'),
#                     'DEVIATION': mes.get('deviation'),
#                     'OUT_OF_TOLERANCE': mes.get('outtol'),
#                     'MEASURED': mes.get('measured')
#                 }
#                 unmatched_data.append(unmatched_record)

#     logging.debug(f"unmatxhed_data count: {unmatched_data}")
#     logging.info(f"Built merged_data rows: {len(merged_data)}; unmatched: {len(unmatched_data)}")

#     # Convert merged_data to DataFrame and drop columns that are all NaN
#     merged_df = pd.DataFrame(merged_data)
#     merged_df = merged_df.dropna(axis=1, how='all')
#     logging.debug(f"Merged DataFrame columns before dropping specified columns: {merged_df.columns.tolist()}")
#     cols_to_rmv = [c for c in COLUMN_TO_RMV if c in merged_df.columns]
#     logging.debug(f"Columns to be removed: {cols_to_rmv}")
#     if cols_to_rmv:
#        merged_df.drop(columns=cols_to_rmv, inplace=True)
#     merged_df = move_measured_columns_to_end(merged_df)
#     logging.debug(f"Merged DataFrame columns after dropping all-NaN and reordering: {merged_df.columns.tolist()}")
#     # Save the data to a temporary Excel file first
#     # temp_output = output_file_path
#     temp_output = output_file_path
#     logging.info(f"Writing to temporary file: {temp_output}")
#     logging.debug(f"Merged DataFrame preview:\n{merged_df.head()}")
    
#     temp_data_file = "data.xlsx"
    
#     # Write DataFrame to Excel using ExcelWriter to control formatting
#     with pd.ExcelWriter(temp_data_file, engine='openpyxl') as writer:
#         merged_df.to_excel(writer, sheet_name='Sheet1', index=False)
    
#     logging.info(f"Data written to temporary file: {temp_data_file}")
#     logging.info(f"header_file_path: {header_file_path}, header_row_idx: {header_row_idx}")
#     logging.info("Data merging process completed successfully.")

#     # # Merge the temporary file with the header file while preserving formatting
#     # try:
#     #     logging.info("Merging temporary file with header file to preserve formatting.")
#     #     merge_excel_with_header(temp_data_file, header_file_path, temp_output, header_row_idx)
#     #     logging.info(f"Final formatted data saved to {temp_output}")
    
#     # except Exception as e:
#     #     logging.error(f"Failed to create excel file: {str(e)}")
#     #     raise
    
#     # finally:
#     #     # Clean up temporary files
#     #     try:
#     #         if os.path.exists(temp_data_file):
#     #             os.remove(temp_data_file)
#     #             logging.info(f"Cleaned up temporary file: {temp_data_file}")
#     #         if os.path.exists("temp_modified.xlsx"):
#     #             os.remove("temp_modified.xlsx")
#     #             logging.info("Cleaned up temporary header file: temp_modified.xlsx")
#     #     except Exception as e:
#     #         logging.warning(f"Failed to clean up temporary files: {str(e)}")
#     # Align all data cells in the center
#     logging.info("Aligning all data cells in the center...")
#     try:
#         data_workbook = openpyxl.load_workbook(temp_data_file)
#         data_worksheet = data_workbook.active

#         # Iterate through ALL cells and set alignment to center
#         for row in data_worksheet.iter_rows(min_row=1, max_row=data_worksheet.max_row, 
#                                            min_col=1, max_col=data_worksheet.max_column):
#             for cell in row:
#                 if cell.value is not None:
#                     # Force center alignment for every cell with a value
#                     cell.alignment = Alignment(horizontal="center", vertical="center")

#         data_workbook.save(temp_data_file)
#         logging.info("Successfully aligned all data cells to center")
#     except Exception as e:
#         logging.warning(f"Failed to align data cells to center: {str(e)}")

#     append_excel_data(temp_data_file, header_file_path, header_row_idx, output_file_path)

# def append_excel_data(temp_data_file, header_file_path, header_row_idx, output_file_path=None):
#     """
#     Create a new Excel file that merges header data and temp data.
#     Preserves formatting from header file while copying data values.
    
#     Args:
#         temp_data_file: Path to the Excel file containing data to merge
#         header_file_path: Path to the Excel file with header data
#         header_row_idx: Row index where header data ends (1-based)
#         output_file_path: Path for the new merged Excel file (optional, defaults to temp_data_file)
#     """
#     if output_file_path is None:
#         output_file_path = temp_data_file
    
#     logging.info(f"Creating new merged Excel file: {output_file_path}")
#     logging.info(f"Merging header from {header_file_path} (up to row {header_row_idx}) with data from {temp_data_file}")

#     try:
#         # Load both workbooks
#         header_wb = openpyxl.load_workbook(header_file_path)
#         temp_wb = openpyxl.load_workbook(temp_data_file)

#         # Get the active sheets
#         header_sheet = header_wb.active
#         temp_sheet = temp_wb.active

#         logging.info(f"Header sheet: {header_sheet.title}, Temp sheet: {temp_sheet.title}")
        
#         # Create new workbook for merged result
#         merged_wb = openpyxl.Workbook()
#         merged_sheet = merged_wb.active
#         merged_sheet.title = "Merged_Data"

#         # First, copy header data with formatting (from row 1 to header_row_idx)
#         header_max_cols = header_sheet.max_column
#         logging.info(f"Copying header data with formatting: rows 1 to {header_row_idx}, columns 1 to {header_max_cols}")
        
#         # Copy column dimensions from header sheet
#         for col_letter, dimension in header_sheet.column_dimensions.items():
#             merged_sheet.column_dimensions[col_letter].width = dimension.width
        
#         # Copy row dimensions and cell data with formatting
#         for row in range(1, header_row_idx + 1):
#             # Copy row height if it exists
#             if row in header_sheet.row_dimensions:
#                 merged_sheet.row_dimensions[row].height = header_sheet.row_dimensions[row].height
                
#             for col in range(1, header_max_cols + 1):
#                 source_cell = header_sheet.cell(row=row, column=col)
#                 target_cell = merged_sheet.cell(row=row, column=col)
                
#                 # Copy value
#                 target_cell.value = source_cell.value
                
#                 # Copy formatting if it exists
#                 if source_cell.has_style:
#                     try:
#                         # Explicitly copy font with all properties including color
#                         if source_cell.font:
#                             target_cell.font = Font(
#                                 name=source_cell.font.name,
#                                 size=source_cell.font.size,
#                                 bold=source_cell.font.bold,
#                                 italic=source_cell.font.italic,
#                                 underline=source_cell.font.underline,
#                                 strike=source_cell.font.strike,
#                                 color=source_cell.font.color
#                             )
#                         target_cell.fill = source_cell.fill
#                         target_cell.border = source_cell.border
                        
#                         #target_cell.alignment = source_cell.alignment
#                         target_cell.number_format = source_cell.number_format
#                         src_align = source_cell.alignment

#                         target_cell.alignment = Alignment(
#                             horizontal=src_align.horizontal,
#                             vertical=src_align.vertical,
#                             text_rotation=src_align.text_rotation,
#                             wrap_text=True,  # Set wrap text to True
#                             shrink_to_fit=src_align.shrink_to_fit,
#                             indent=src_align.indent
#                         )
                        
#                     except Exception as e:
#                         logging.debug(f"Could not copy formatting for header row {row}, col {col}: {str(e)}")

#         # Copy images from header sheet
#         if Image is not None:
#             try:
#                 if hasattr(header_sheet, '_images') and header_sheet._images:
#                     logging.info(f"Found {len(header_sheet._images)} images in header sheet")
#                     for image in header_sheet._images:
#                         try:
#                             # Create a new image object with the same properties
#                             new_image = Image(image.ref)
#                             new_image.anchor = image.anchor
#                             if hasattr(image, 'width'):
#                                 new_image.width = image.width
#                             if hasattr(image, 'height'):
#                                 new_image.height = image.height
                            
#                             # Add the image to the merged sheet
#                             merged_sheet.add_image(new_image)
#                             logging.info(f"Successfully copied image at anchor: {image.anchor}")
#                         except Exception as img_error:
#                             logging.warning(f"Could not copy individual image: {str(img_error)}")
#                 else:
#                     logging.info("No images found in header sheet")
#             except Exception as e:
#                 logging.warning(f"Could not access images from header sheet: {str(e)}")
#         else:
#             logging.warning("Image support not available - images will not be copied")

#         # Then, copy ALL temp data (including all rows from temp file)
#         temp_data_start_row = 1  # Copy all rows including the first row
#         temp_data_end_row = temp_sheet.max_row
#         temp_max_cols = temp_sheet.max_column
        
#         total_data_rows = temp_data_end_row - temp_data_start_row + 1
#         logging.info(f"Copying ALL temp data: {total_data_rows} rows (from row {temp_data_start_row} to {temp_data_end_row}), columns 1 to {temp_max_cols}")

#         if total_data_rows > 0:
#             # Start copying temp data after header_row_idx
#             merged_start_row = header_row_idx + 1
            
#             for temp_row in range(temp_data_start_row, temp_data_end_row + 1):
#                 # Calculate target row in merged sheet
#                 target_row = merged_start_row + (temp_row - temp_data_start_row)
                
#                 # Set default row height for data rows
#                 #merged_sheet.row_dimensions[target_row].height = 15
#                 merged_sheet.row_dimensions[target_row].height = 40
                
#                 # Copy all columns from this row with header formatting applied to data
#                 for col in range(1, temp_max_cols + 1):
#                     source_cell = temp_sheet.cell(row=temp_row, column=col)
#                     target_cell = merged_sheet.cell(row=target_row, column=col)
                    
#                     # Copy value from temp data
#                     target_cell.value = source_cell.value
                    
#                     # ALWAYS apply center alignment to ALL data cells with values
#                     if target_cell.value is not None:
#                         target_cell.alignment = Alignment(horizontal="center", vertical="center",wrap_text=True)


                    
#                     # Apply formatting from corresponding header column (use header_row_idx as template)
#                     if col <= header_max_cols:
#                         header_template_cell = header_sheet.cell(row=header_row_idx, column=col)
#                         if header_template_cell.has_style:
#                             try:
#                                 # Explicitly copy font with all properties including color
#                                 if header_template_cell.font:
#                                     target_cell.font = Font(
#                                         name=header_template_cell.font.name,
#                                         size=header_template_cell.font.size,
#                                         bold=header_template_cell.font.bold,
#                                         italic=header_template_cell.font.italic,
#                                         underline=header_template_cell.font.underline,
#                                         strike=header_template_cell.font.strike,
#                                         color=header_template_cell.font.color
#                                     )
#                                 target_cell.fill = header_template_cell.fill
#                                 target_cell.border = header_template_cell.border
#                                 target_cell.number_format = header_template_cell.number_format
                                
#                             except Exception as e:
#                                 logging.debug(f"Could not copy header formatting to data row {target_row}, col {col}: {str(e)}")

#         # Add borders to all cells for print formatting
#         logging.info("Adding borders to all cells for print formatting...")
#         try:
#             # Create a thin border style
#             thin_border = Border(
#                 left=Side(style='thin'),
#                 right=Side(style='thin'),
#                 top=Side(style='thin'),
#                 bottom=Side(style='thin')
#             )
            
#             # Apply borders to all cells that have content
#             max_row_with_data = merged_sheet.max_row
#             max_col_with_data = merged_sheet.max_column
            
#             for row in range(1, max_row_with_data + 1):
#                 for col in range(1, max_col_with_data + 1):
#                     cell = merged_sheet.cell(row=row, column=col)
#                     #if cell.value is not None or row <= header_row_idx:  # Apply to header rows and data cells
#                     # Preserve existing formatting while adding borders
#                     current_font = cell.font
#                     current_fill = cell.fill
#                     current_alignment = cell.alignment
#                     current_number_format = cell.number_format
                    
#                     # Apply border while keeping other formatting
#                     cell.border = thin_border
                        
#             logging.info(f"Successfully added borders to {max_row_with_data} rows x {max_col_with_data} columns")
            
#         except Exception as e:
#             logging.warning(f"Could not add borders: {str(e)}")

#         # Save the new merged file
#         merged_wb.save(output_file_path)
#         logging.info(f"Successfully created merged Excel file with borders: {output_file_path}")
#         logging.info(f"Total rows in merged file: {merged_sheet.max_row}")

#     except Exception as e:
#         logging.error(f"Error during Excel merge: {str(e)}", exc_info=True)
#         raise
#     finally:
#         # Close workbooks to free memory
#         if 'header_wb' in locals():
#             header_wb.close()
#         if 'temp_wb' in locals():
#             temp_wb.close()
#         if 'merged_wb' in locals():
#             merged_wb.close()




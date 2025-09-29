import pandas as pd
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

    # Read the Excel file
    df = pd.read_excel(file_path, header=None)
    logging.debug(f"Original DataFrame shape: {df.shape}")

    # Remove rows after the specified row number
    df = df.iloc[:row_number]
    logging.info(f"Rows after row {row_number} have been removed.")
    logging.debug(f"Modified DataFrame shape: {df.shape}")

    # Save the modified DataFrame to an Excel file
    # df.to_excel(output_file, index=False, header=False)
    # logging.info(f"Modified DataFrame saved to {output_file}")
    return df

if __name__ == "__main__":
    output_file = "example.xlsx"  # Replace with your Excel file path
    row_number = 10  # Replace with the row number after which rows should be removed
    file_path = "/mnt/c/Users/admin/Desktop/conversion/conversion/REPORT/302/302.xls"  # Replace with your output file path

    remove_rows_after(file_path, row_number, output_file)
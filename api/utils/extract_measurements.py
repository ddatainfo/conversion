import re
import os
import chardet
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def extract_measurements(file_path):
    logging.info(f"Starting extraction of measurements from file: {file_path}")

    measurements = []
    current_dim = None
    units = None

    # Detect file encoding
    with open(file_path, 'rb') as f:
        raw_data = f.read()
    detected = chardet.detect(raw_data)
    encoding = detected['encoding']

    # Read the file with the detected encoding
    with open(file_path, 'r', encoding=encoding) as f:
        lines = f.readlines()

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        # Detect start of a measurement block
        dim_match = re.match(r'(DIM\s+#?.*?=\s*.+?)(UNITS=MM)', line)
        if dim_match:
            current_dim = dim_match.group(1).strip()
            units = 'MM'
            i += 2  # Skip header line (AX ...)
            # Collect measurement rows
            while i < len(lines):
                row = lines[i].strip()
                if not row or row.startswith('DIM') or row.startswith('PART NUMBER'):
                    break
                # Parse measurement row
                parts = re.split(r'\s+', row)
                if len(parts) >= 7:
                    ax = parts[0]
                    nominal = parts[1]
                    ptol = parts[2]
                    mtol = parts[3]
                    meas = parts[4]
                    dev = parts[5]
                    outtol = parts[6]
                    symbols = ' '.join(parts[7:]) if len(parts) > 7 else ''
                    measurement = {
                        'dimension': current_dim,
                        'units': units,
                        'axis': ax,
                        'nominal': nominal,
                        '+tol': ptol,
                        '-tol': mtol,
                        'measured': meas,
                        'deviation': dev,
                        'outtol': outtol,
                        'symbols': symbols
                    }
                    logging.debug(f"Extracted measurement: {measurement}")
                    measurements.append(measurement)
                i += 1
        else:
            i += 1
    logging.info(f"Extraction completed. Total measurements extracted: {len(measurements)}")
    return measurements

def process_and_write_measurements(folder_path):
    # Iterate through all TXT files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.TXT'):
            file_path = os.path.join(folder_path, file_name)
            # Extract measurements from the file
            logging.info(f"Processing file: {file_name}")
            measurements = extract_measurements(file_path)
            # Prepare output file path
            output_file_path = os.path.join(folder_path, f"output_{file_name}")
            with open(output_file_path, 'w', encoding='utf-8') as output_file:
                output_file.write(f"INPUT FILE: {file_name}\n")
                output_file.write("OUTPUT: EXTRACTED DATA:\n")
                output_file.write("-------------------\n")
                for entry in measurements:
                    output_file.write(f"{entry}\n")
                    output_file.write("-------------------\n")
            logging.info(f"Processed and wrote measurements to: {output_file_path}")

# Example usage
if __name__ == "__main__":
    folder_path = "/mnt/c/Users/admin/Desktop/conversion/TXT"  # WSL path
    process_and_write_measurements(folder_path)
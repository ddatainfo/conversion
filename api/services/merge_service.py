import sys
sys.path.append("../")  # Add parent directory to sys.path for relative imports
from utils.merge_data import final_data
import os
import uuid

def process_files(excel_path: str, txt_path: str) -> str:
    unique_id = uuid.uuid4().hex  # Generate a unique identifier
    output_filename = f"merged_output_{unique_id}.xlsx"
    output_path = os.path.join("uploads", output_filename)
    final_data(excel_path, txt_path, output_path)
    return output_path

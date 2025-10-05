import sys
sys.path.append("../")  # Add parent directory to sys.path for relative imports
from api.utils.merge_data import final_data
import os
import uuid
from typing import List


def process_files(excel_path: str, txt_paths: List[str]) -> str:
    """Process an Excel file and a list of TXT file paths. Returns output path."""
    unique_id = uuid.uuid4().hex  # Generate a unique identifier
    output_filename = f"merged_output_{unique_id}.xlsx"
    output_path = os.path.join("uploads", output_filename)
    final_data(excel_path, txt_paths, output_path)
    return output_path

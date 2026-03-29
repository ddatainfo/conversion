import sys
sys.path.append("../")  # Add parent directory to sys.path for relative imports
from api.utils.merge_data import final_data
import os
import uuid
from typing import List, Dict, Union


def process_files(excel_path: str, txt_paths: Union[List[str], Dict[str, List[str]]]) -> str:
    """
    Process an Excel file and TXT file paths (organized by folders or as a flat list).
    
    Args:
        excel_path: Path to the Excel file
        txt_paths: Either a list of file paths or a dict organized by folder
                   Example: {"1": ["path/1.txt", "path/2.txt"], "2": ["path/1.txt"]}
                   Folder context is preserved throughout processing
    
    Returns:
        Path to the output merged Excel file
    """
    unique_id = uuid.uuid4().hex  # Generate a unique identifier
    output_filename = f"merged_output_{unique_id}.xlsx"
    # Get the absolute path of the uploads directory
    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    output_dir = os.path.join(base_dir, "uploads")
    os.makedirs(output_dir, exist_ok=True)  # Ensure the directory exists
    output_path = os.path.join(output_dir, output_filename)
    
    # Pass txt_paths as-is (dict or list) to final_data for processing
    # Folder structure is preserved: {"folder_name": [list_of_file_paths]}
    final_data(excel_path, txt_paths, output_path)
    return output_path

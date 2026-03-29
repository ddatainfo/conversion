from fastapi import APIRouter, UploadFile, File, HTTPException
from typing import List, Dict
from api.services.merge_service import process_files
from fastapi.responses import FileResponse
import os
import zipfile
import tempfile
import shutil

router = APIRouter()

# Get absolute path for uploads directory
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

@router.post("/upload/")
async def upload_files(excel_file: UploadFile = File(...), folder_files: UploadFile = File(...)):
    """
    Upload an Excel file and a ZIP file containing multiple folders with TXT files.
    
    The folder_files should be a ZIP file with structure:
    - 1/file-name-1.txt
    - 1/file-name-2.txt
    - 2/file1.txt
    - 3/file1.txt
    - 4/file1.txt
    
    Returns organized dictionary: {"1": ["path/file1.txt", "path/file2.txt"], "2": [...]}
    
    Args:
        excel_file: The Excel template file (.xlsx or .xls)
        folder_files: ZIP file containing folders with TXT files
    
    Returns:
        Processed Excel file with merged data organized by folders
    """
    
    # Validate excel
    if not (excel_file.filename.endswith(".xlsx") or excel_file.filename.endswith(".xls")):
        raise HTTPException(status_code=400, detail="Invalid Excel file format. Only .xlsx and .xls files are allowed.")
    
    # Validate zip file
    if not folder_files.filename.endswith(".zip"):
        raise HTTPException(status_code=400, detail="Invalid file format. Only .zip files are allowed.")
    
    # Save excel file
    excel_path = os.path.join(UPLOAD_DIR, excel_file.filename)
    with open(excel_path, "wb") as f:
        f.write(await excel_file.read())
    
    # Create a temporary directory for extraction
    temp_extract_dir = tempfile.mkdtemp()
    
    try:
        # Save zip file temporarily
        zip_temp_path = os.path.join(temp_extract_dir, "temp.zip")
        zip_content = await folder_files.read()
        with open(zip_temp_path, "wb") as f:
            f.write(zip_content)
        
        # Extract zip file
        with zipfile.ZipFile(zip_temp_path, 'r') as zip_ref:
            zip_ref.extractall(temp_extract_dir)
        
        # Build dictionary of folder -> txt files
        txt_dict: Dict[str, List[str]] = {}
        
        # Walk through extracted folders
        for folder_name in os.listdir(temp_extract_dir):
            folder_path = os.path.join(temp_extract_dir, folder_name)
            
            # Skip if not a directory or if it's the zip file
            if not os.path.isdir(folder_path) or folder_name == "temp.zip":
                continue
            
            # Create folder in uploads directory
            upload_folder_path = os.path.join(UPLOAD_DIR, folder_name)
            os.makedirs(upload_folder_path, exist_ok=True)
            
            # Find all txt files in this folder
            txt_files = []
            for file_name in os.listdir(folder_path):
                file_path = os.path.join(folder_path, file_name)
                
                # Check if it's a txt file
                if os.path.isfile(file_path) and (file_name.endswith(".txt") or file_name.endswith(".TXT")):
                    # Copy to uploads folder
                    dest_file_path = os.path.join(upload_folder_path, file_name)
                    shutil.copy(file_path, dest_file_path)
                    txt_files.append(dest_file_path)
            
            # Add to dictionary if folder has txt files
            if txt_files:
                txt_dict[folder_name] = txt_files
        
        if not txt_dict:
            raise HTTPException(status_code=400, detail="No txt files found in the uploaded ZIP file")
        
        # Pass excel_path and txt_dict to process_files
        output_path = process_files(excel_path, txt_dict)
        
        # Verify file exists before sending response
        if not os.path.exists(output_path):
            raise HTTPException(
                status_code=500,
                detail="Output file was not generated successfully"
            )
        
        return FileResponse(
            output_path, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="merged_output.xlsx"
        )
    
    except zipfile.BadZipFile:
        raise HTTPException(status_code=400, detail="Invalid ZIP file. Please upload a valid ZIP file.")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error processing files: {str(e)}"
        )
    finally:
        # Clean up temporary directory
        shutil.rmtree(temp_extract_dir, ignore_errors=True)


# from fastapi import APIRouter, UploadFile, File, HTTPException, Request
# from typing import List, Dict
# from api.services.merge_service import process_files
# from fastapi.responses import FileResponse
# import os

# router = APIRouter()

# # Get absolute path for uploads directory
# BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
# os.makedirs(UPLOAD_DIR, exist_ok=True)

# @router.post("/upload/")
# async def upload_files(request: Request):
#     """
#     Upload an Excel file and multiple folders containing TXT files.
    
#     Form fields:
#     - excel_file: The Excel template file (.xlsx or .xls)
#     - folder_1: Files for folder 1
#     - folder_2: Files for folder 2
#     - folder_3: Files for folder 3
#     - folder_4: Files for folder 4
    
#     Example curl:
#     curl -X POST "http://localhost:8000/upload/" \\
#       -F "excel_file=@report.xlsx" \\
#       -F "folder_1=@/path/to/folder1/file1.txt" \\
#       -F "folder_1=@/path/to/folder1/file2.txt" \\
#       -F "folder_2=@/path/to/folder2/file1.txt" \\
#       -F "folder_3=@/path/to/folder3/file1.txt" \\
#       -F "folder_4=@/path/to/folder4/file1.txt"
    
#     Returns:
#         Processed Excel file with merged data organized by folders
#     """
    
#     try:
#         form_data = await request.form()
        
#         # Validate and get Excel file
#         if "excel_file" not in form_data:
#             raise HTTPException(status_code=400, detail="Excel file is required")
        
#         excel_file = form_data["excel_file"]
#         if not (excel_file.filename.endswith(".xlsx") or excel_file.filename.endswith(".xls")):
#             raise HTTPException(status_code=400, detail="Invalid Excel file format. Only .xlsx and .xls files are allowed.")
        
#         # Save excel file
#         excel_path = os.path.join(UPLOAD_DIR, excel_file.filename)
#         with open(excel_path, "wb") as f:
#             f.write(await excel_file.read())
        
#         # Process folders and txt files
#         txt_dict: Dict[str, List[str]] = {}
        
#         # Iterate through all form fields
#         for field_name, field_value in form_data.items():
#             if field_name == "excel_file":
#                 continue
            
#             # Extract folder name from field name (e.g., "folder_1" -> "1")
#             if field_name.startswith("folder_"):
#                 folder_name = field_name.replace("folder_", "")
#             else:
#                 folder_name = field_name
            
#             # Handle multiple files for the same folder
#             files = field_value if isinstance(field_value, list) else [field_value]
            
#             for file_obj in files:
#                 # Validate file extension
#                 if not (file_obj.filename.endswith(".txt") or file_obj.filename.endswith(".TXT")):
#                     raise HTTPException(status_code=400, detail=f"Invalid file format: {file_obj.filename}. Only .txt files are allowed.")
                
#                 # Create folder directory
#                 folder_path = os.path.join(UPLOAD_DIR, folder_name)
#                 os.makedirs(folder_path, exist_ok=True)
                
#                 # Save the file
#                 full_file_path = os.path.join(folder_path, file_obj.filename)
#                 with open(full_file_path, "wb") as f:
#                     f.write(await file_obj.read())
                
#                 # Add to dictionary: {folder_name: [list of file paths]}
#                 if folder_name not in txt_dict:
#                     txt_dict[folder_name] = []
#                 txt_dict[folder_name].append(full_file_path)
        
#         if not txt_dict:
#             raise HTTPException(status_code=400, detail="No txt files uploaded")
        
#         # Pass excel_path and txt_dict (organized by folders) to process_files
#         output_path = process_files(excel_path, txt_dict)
        
#         # Verify file exists before sending response
#         if not os.path.exists(output_path):
#             raise HTTPException(
#                 status_code=500,
#                 detail="Output file was not generated successfully"
#             )
        
#         return FileResponse(
#             output_path, 
#             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#             filename="merged_output.xlsx"
#         )
#     except HTTPException:
#         raise
#     except Exception as e:
#         raise HTTPException(
#             status_code=500,
#             detail=f"Error processing files: {str(e)}"
#         )

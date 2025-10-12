from fastapi import APIRouter, UploadFile, File, HTTPException
from typing import List
from api.services.merge_service import process_files
from fastapi.responses import FileResponse
import os

router = APIRouter()

# Get absolute path for uploads directory
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

@router.post("/upload/")
async def upload_files(excel_file: UploadFile = File(...), txt_files: List[UploadFile] = File(...)):
    # Validate excel
    if not (excel_file.filename.endswith(".xlsx") or excel_file.filename.endswith(".xls")):
        raise HTTPException(status_code=400, detail="Invalid Excel file format. Only .xlsx and .xls files are allowed.")

    # Validate txt files
    for t in txt_files:
        if not (t.filename.endswith(".txt") or t.filename.endswith(".TXT")):
            raise HTTPException(status_code=400, detail="Invalid TXT file format. Only .txt files are allowed.")

    # Save excel
    excel_path = os.path.join(UPLOAD_DIR, excel_file.filename)
    with open(excel_path, "wb") as f:
        f.write(await excel_file.read())

    # Save all txt files and collect their paths
    txt_paths = []
    for t in txt_files:
        path = os.path.join(UPLOAD_DIR, t.filename)
        with open(path, "wb") as f:
            f.write(await t.read())
        txt_paths.append(path)

    try:
        output_path = process_files(excel_path, txt_paths)
        
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
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error processing files: {str(e)}"
        )

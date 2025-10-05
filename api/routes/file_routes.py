from fastapi import APIRouter, UploadFile, File, HTTPException
from typing import List
from api.services.merge_service import process_files
from fastapi.responses import FileResponse
import os

router = APIRouter()

UPLOAD_DIR = "uploads"
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

    output_path = process_files(excel_path, txt_paths)

    return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="merged_output.xlsx")

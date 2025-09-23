from fastapi import APIRouter, UploadFile, File, HTTPException
from services.merge_service import process_files
from fastapi.responses import FileResponse
import os

router = APIRouter()

UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

@router.post("/upload/")
async def upload_files(excel_file: UploadFile = File(...), txt_file: UploadFile = File(...)):
    if not (excel_file.filename.endswith(".xlsx") or excel_file.filename.endswith(".xls")):
        raise HTTPException(status_code=400, detail="Invalid Excel file format. Only .xlsx and .xls files are allowed.")
    if not (txt_file.filename.endswith(".txt") or txt_file.filename.endswith(".TXT")):
        raise HTTPException(status_code=400, detail="Invalid TXT file format. Only .txt files are allowed.")

    excel_path = os.path.join(UPLOAD_DIR, excel_file.filename)
    txt_path = os.path.join(UPLOAD_DIR, txt_file.filename)

    with open(excel_path, "wb") as f:
        f.write(await excel_file.read())
    with open(txt_path, "wb") as f:
        f.write(await txt_file.read())

    output_path = process_files(excel_path, txt_path)

    return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="merged_output.xlsx")

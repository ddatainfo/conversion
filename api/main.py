from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from api.routes.file_routes import router as file_router
import os

app = FastAPI()

app.include_router(file_router, prefix="/files")

# Mount static files directory if it exists
if os.path.exists("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
def read_root():
    """Serve the main HTML page"""
    return FileResponse("templates/index.html")

@app.get("/api")
def api_root():
    """API root endpoint"""
    return {"message": "Welcome to the File Merge API!"}

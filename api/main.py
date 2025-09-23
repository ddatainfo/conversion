from fastapi import FastAPI
from routes.file_routes import router as file_router

app = FastAPI()

app.include_router(file_router, prefix="/files")

@app.get("/")
def read_root():
    return {"message": "Welcome to the File Merge API!"}

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import os
import shutil
import mimetypes
from pathlib import Path

UPLOAD_DIR = Path(__file__).resolve().parent / "uploaded_files"
EXPORT_DIR = Path(__file__).resolve().parent / "exported_files"
UPLOAD_DIR.mkdir(exist_ok=True)
EXPORT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="File Send & Receive API")


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Nhận file từ client và lưu vào thư mục uploaded_files."""
    safe_name = Path(file.filename).name
    if not safe_name:
        raise HTTPException(status_code=400, detail="Invalid filename")

    dest = UPLOAD_DIR / safe_name
    with open(dest, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return {"status": "success", "filename": safe_name, "size": os.path.getsize(dest)}


@app.get("/download/{filename}")
async def download_file(filename: str):
    """Gửi file từ server (exported_files) về client."""
    safe_name = Path(filename).name
    file_path = EXPORT_DIR / safe_name

    if not file_path.exists() or not file_path.is_file():
        raise HTTPException(status_code=404, detail="File not found")

    media_type, _ = mimetypes.guess_type(safe_name)
    return FileResponse(
        path=str(file_path),
        filename=safe_name,
        media_type=media_type or "application/octet-stream",
    )


@app.get("/files")
async def list_files():
    """Liệt kê các file có sẵn để download."""
    files = [f.name for f in EXPORT_DIR.iterdir() if f.is_file()]
    return {"files": files}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "API_send_receive_file:app",
        host="0.0.0.0",
        port=8001,
        workers=1,
        reload=False,
    )

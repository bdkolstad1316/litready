"""
LitReady — Unified Server
==========================
Serves the frontend AND the cleaning API from one app.
One repo. One deploy. One URL.

Deploy to Railway:
    1. Create a GitHub repo with these files
    2. Go to railway.app → New Project → Deploy from GitHub repo
    3. That's it. Railway handles the rest.
"""

import os
import shutil
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

from litready_engine import process_docx

app = FastAPI(title="LitReady", version="1.0.0")


@app.get("/health")
def health_check():
    return {"status": "ok", "service": "litready"}


@app.post("/clean")
async def clean_document(
    file: UploadFile = File(...),
    genre: str = Form(default="prose"),
):
    """Upload a .docx, get a cleaned .docx back."""
    
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Only .docx files are accepted")
    
    if genre not in ('prose', 'poetry', 'nonfiction', 'hybrid'):
        raise HTTPException(status_code=400, detail="Genre must be prose, poetry, nonfiction, or hybrid")
    
    output_tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    output_tmp.close()
    
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / file.filename
            with open(input_path, 'wb') as f:
                content = await file.read()
                f.write(content)
            
            stem = input_path.stem
            processing_output = Path(tmpdir) / f"{stem}_CLEAN.docx"
            
            try:
                process_docx(str(input_path), str(processing_output), genre=genre)
            except Exception as e:
                os.unlink(output_tmp.name)
                raise HTTPException(status_code=500, detail=f"Processing failed: {str(e)}")
            
            if not processing_output.exists():
                os.unlink(output_tmp.name)
                raise HTTPException(status_code=500, detail="Output file was not created")
            
            shutil.copy2(str(processing_output), output_tmp.name)
        
        def cleanup():
            try:
                os.unlink(output_tmp.name)
            except OSError:
                pass
        
        return FileResponse(
            path=output_tmp.name,
            filename=f"{stem}_CLEAN.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            background=BackgroundTask(cleanup),
        )
    except HTTPException:
        raise
    except Exception as e:
        try:
            os.unlink(output_tmp.name)
        except OSError:
            pass
        raise HTTPException(status_code=500, detail=str(e))


# Serve the frontend — this MUST be last (catch-all)
app.mount("/", StaticFiles(directory="static", html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)

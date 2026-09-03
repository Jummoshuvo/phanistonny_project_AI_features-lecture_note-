"""
Backend API for Lecture Notes Study App
FastAPI / Python Web Server
"""
import os
import threading
import dotenv
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Dict, Any, Optional

dotenv.load_dotenv()

from lecture_note_engine import (
    extract_raw_text,
    build_visual_notes_from_text,
    apply_ai_edit,
    get_realistic_diagram_metadata,
    get_ocr_reader
)

app = FastAPI(title="Lecture Notes Study App API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_origin_regex=".*",
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

assets_dir = os.path.join(os.path.dirname(__file__), "assets")
if os.path.exists(assets_dir):
    from fastapi.staticfiles import StaticFiles
    app.mount("/assets", StaticFiles(directory=assets_dir), name="assets")

# Pre-warm EasyOCR model in a background thread on startup
def _warmup_ocr():
    try:
        get_ocr_reader()
    except Exception as e:
        print(f"[OCR Warmup Notice] {e}")

threading.Thread(target=_warmup_ocr, daemon=True).start()

class EditRequest(BaseModel):
    plan: Dict[str, Any]
    instruction: str

class TextNoteRequest(BaseModel):
    title: Optional[str] = "Manual Lecture Notes"
    text: str

class DiagramRequest(BaseModel):
    topic: str
    text: Optional[str] = ""

def get_lecture_note_html_content():
    file_path = os.path.join(os.path.dirname(__file__), "lecture_note.html")
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="lecture_note.html file not found")
    with open(file_path, "r", encoding="utf-8") as f:
        return f.read()

@app.get("/", response_class=HTMLResponse)
async def get_root():
    return get_lecture_note_html_content()

@app.get("/lecture_note", response_class=HTMLResponse)
async def get_lecture_note():
    return get_lecture_note_html_content()

@app.post("/api/v1/lecture-notes")
async def upload_lecture_notes(file: UploadFile = File(...)):
    filename = file.filename or "Uploaded_Notes.pdf"
    try:
        content_bytes = await file.read()
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {e}")
    
    try:
        extracted_text = extract_raw_text(filename, content_bytes)
    except Exception as e:
        print(f"[Text Extraction Error] {e}")
        extracted_text = f"--- {filename} Content ---\nExtracted text ready for visual notes generation."

    try:
        plan = build_visual_notes_from_text(filename, extracted_text)
    except Exception as e:
        print(f"[Visual Notes Build Error] {e}")
        from lecture_note_engine import create_default_plan
        plan = create_default_plan(filename=filename, text_content=extracted_text)
    
    plan["filename"] = filename
    if "original_text" not in plan or not plan["original_text"]:
        plan["original_text"] = extracted_text
    return JSONResponse(content=plan)

@app.post("/api/v1/lecture-notes/from-text")
async def generate_from_text(payload: TextNoteRequest):
    title = payload.title or "Manual Lecture Notes"
    text = payload.text or ""
    try:
        plan = build_visual_notes_from_text(title, text)
    except Exception as e:
        print(f"[From Text Build Error] {e}")
        from lecture_note_engine import create_default_plan
        plan = create_default_plan(filename=title, text_content=text)
    
    plan["filename"] = title
    if "original_text" not in plan or not plan["original_text"]:
        plan["original_text"] = text
    return JSONResponse(content=plan)

@app.post("/api/v1/lecture-notes/{plan_id}/edit")
async def edit_lecture_notes(plan_id: str, payload: EditRequest):
    plan = payload.plan
    instruction = payload.instruction
    if not plan:
        raise HTTPException(status_code=400, detail="Plan object is required")
    
    try:
        updated_plan = apply_ai_edit(plan, instruction)
    except Exception as e:
        print(f"[AI Edit Error] {e}")
        updated_plan = plan
    return JSONResponse(content=updated_plan)

@app.post("/api/v1/generate-diagram-image")
async def generate_diagram_image(payload: DiagramRequest):
    metadata = get_realistic_diagram_metadata(payload.topic, payload.text or "")
    return JSONResponse(content=metadata)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)

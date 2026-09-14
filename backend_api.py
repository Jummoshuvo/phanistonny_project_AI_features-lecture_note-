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
    analyze_document_for_original_notes,
    build_deterministic_original_note,
    build_visual_notes_from_original_note,
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
    filename: Optional[str] = ""

class VisualNoteGenRequest(BaseModel):
    original_note: Optional[Dict[str, Any]] = None
    filename: Optional[str] = "Uploaded_Notes.pdf"
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
@app.post("/api/v1/analyze-document")
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
        extracted_text = f"--- {filename} Content ---\nExtracted text ready for document analysis."

    try:
        original_note = analyze_document_for_original_notes(filename, extracted_text)
    except Exception as e:
        print(f"[Original Notes Build Error] {e}")
        from lecture_note_engine import build_deterministic_original_note
        original_note = build_deterministic_original_note(filename, extracted_text)
    
    original_note["filename"] = filename
    original_note["original_text"] = extracted_text

    return JSONResponse(content=original_note)

@app.post("/api/v1/lecture-notes/from-text")
async def generate_from_text(payload: TextNoteRequest):
    title = payload.title or "Manual Lecture Notes"
    text = payload.text or ""
    try:
        original_note = analyze_document_for_original_notes(title, text)
    except Exception as e:
        print(f"[From Text Build Error] {e}")
        from lecture_note_engine import build_deterministic_original_note
        original_note = build_deterministic_original_note(title, text)
    
    original_note["filename"] = title
    original_note["original_text"] = text

    try:
        visual_plan = build_visual_notes_from_original_note(original_note)
        return JSONResponse(content=visual_plan)
    except Exception as e:
        print(f"[Visual Plan Build Error] {e}")
        return JSONResponse(content=original_note)

@app.post("/api/v1/generate-visual-notes")
async def generate_visual_notes_endpoint(payload: VisualNoteGenRequest):
    orig = payload.original_note
    filename = payload.filename or (orig.get("filename") if orig else None) or "Uploaded_Notes.pdf"
    text = payload.text or (orig.get("original_text") if orig else "")
    
    if not orig:
        orig = analyze_document_for_original_notes(filename, text)
    else:
        orig["filename"] = filename
        
    try:
        plan = build_visual_notes_from_original_note(orig)
        plan["filename"] = filename
    except Exception as e:
        print(f"[Visual Notes Gen Error] {e}")
        from lecture_note_engine import create_default_plan
        plan = create_default_plan(filename=filename, text_content=text)
        plan["filename"] = filename
        
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
    try:
        from lecture_note_engine import detect_primary_subject, resolve_or_generate_visual
        topic = payload.topic or "Medical Concept"
        text = payload.text or ""
        doc_subject, doc_name = detect_primary_subject(topic, text)
        
        topic_lower = topic.lower()
        if "pathway" in topic_lower or "flow" in topic_lower or "mechanism" in topic_lower or "patho" in topic_lower:
            return JSONResponse(content={
                "success": False,
                "generated": False,
                "reused": False,
                "asset_url": None,
                "type": "pathway",
                "subject": doc_name,
                "message": "No image required for Pathophysiology/Pathway; presented as an icon-based step flowchart."
            })
        elif "anatomy" in topic_lower or "structure" in topic_lower:
            vis_type = "anatomy"
        elif "medication" in topic_lower or "drug" in topic_lower or "treatment" in topic_lower:
            vis_type = "medication"
        else:
            vis_type = "organ"
            
        req = {
            "type": vis_type,
            "subject": doc_name,
            "section_title": topic or vis_type,
            "purpose": f"Show {vis_type} for {doc_name} as described in lecture notes",
            "required": True
        }
        
        doc_filename = payload.filename or payload.topic or ""
        asset_url, generated, reused = resolve_or_generate_visual(req, text, filename=doc_filename)
        
        if asset_url:
            return JSONResponse(content={
                "success": True,
                "generated": generated,
                "reused": reused,
                "asset_url": asset_url,
                "type": vis_type,
                "subject": doc_name
            })
        else:
            return JSONResponse(content={
                "success": False,
                "generated": False,
                "reused": False,
                "asset_url": None,
                "type": vis_type,
                "subject": doc_name,
                "error": "Image generation failed or API key unavailable"
            })
    except Exception as e:
        return JSONResponse(content={
            "success": False,
            "generated": False,
            "reused": False,
            "asset_url": None,
            "error": str(e)
        })

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)

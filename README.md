# AI Lecture Notes Study App

An AI-powered lecture notes study application that converts document uploads into color-coded visual study notes, system views, diagrams, and interactive AI-promptable edits. Built with a FastAPI backend and a vanilla JS single-page frontend.

## Features
- **Two-Column Layout**: Fixed left sidebar + 3-tab main content card matching modern SaaS dashboard aesthetics.
- **Original Notes Tab**: Displays raw extracted document text with a one-click download option.
- **Visual Notes Tab**: Grid of color-coded cards (Blue, Red, Purple, Green accents) with an inline SVG 4-chamber heart diagram, flow chips, and full-width "Clinical Tip" (amber) & "Remember" (purple) callout boxes.
- **Edit Tab**: Split view showing a dimmed preview of visual notes alongside an interactive AI Prompt panel with suggestion chips, chat log, and prompt submission input.
- **FastAPI Backend + Standalone Frontend**: Works as a single static HTML file (`lecture_note.html`) or served via the FastAPI web backend (`backend_api.py`).

## File Structure
- `lecture_note.html`: Main single static HTML application (with inline CSS & JS).
- `backend_api.py`: FastAPI server handling app routes (`/` and `/lecture_note`) and API endpoints (`/api/v1/lecture-notes`).
- `lecture_note_engine.py`: Engine handling document parsing, EasyOCR, and AI visual transformations.
- `requirements.txt`: Python dependencies (`fastapi`, `uvicorn`, `easyocr`, etc.).

## Quick Start

### Option 1: Static HTML (No Server Required)
Simply double-click or open `lecture_note.html` directly in any modern web browser.

### Option 2: Python FastAPI Backend

1. **Navigate to the Project Directory**:
   ```bash
   cd "lecture_note Py Script"
   ```

2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run FastAPI Backend**:
   ```bash
   python backend_api.py
   ```
   ```bash
   python -m uvicorn backend_api:app --port 8000
   ```

4. **Open App**:
   Navigate to any of the following URLs in your browser:
   - `http://127.0.0.1:8000/`
   - `http://127.0.0.1:8000/lecture_note`


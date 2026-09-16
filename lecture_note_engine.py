"""
Lecture Notes Engine: Medical & Nursing Visual Notes System
1:1 Document -> Original Notes + Visual Notes Pipeline with Smart Asset Resolution.
"""
import os
import uuid
import re
import io
import json
import dotenv
from typing import Dict, Any, List, Optional, Tuple
import difflib

dotenv.load_dotenv()

# --- Document Parsers ---
try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None

try:
    import pypdf
except ImportError:
    pypdf = None

try:
    import docx
except ImportError:
    docx = None

try:
    import pptx
except ImportError:
    pptx = None

_easyocr_reader = None

def get_ocr_reader():
    """Initializes and returns EasyOCR reader singleton."""
    global _easyocr_reader
    if _easyocr_reader is None:
        try:
            import easyocr
            _easyocr_reader = easyocr.Reader(['en'], gpu=False)
        except Exception as e:
            print(f"[EasyOCR Init Notice] {e}")
    return _easyocr_reader

def extract_raw_text(file_name: str, content_bytes: bytes) -> str:
    """Extracts raw text cleanly from PDF, Word, PowerPoint, Images, or Text without leaking binary metadata."""
    if not content_bytes:
        return ""
    ext = file_name.rsplit('.', 1)[-1].lower() if '.' in file_name else ""
    extracted_pages = []

    # 1. PDF via PyMuPDF (fitz)
    if ext == "pdf" and fitz:
        try:
            doc = fitz.open(stream=content_bytes, filetype="pdf")
            for idx, page in enumerate(doc, start=1):
                txt = page.get_text("text")
                if txt and len(txt.strip()) > 5:
                    extracted_pages.append(f"--- Page {idx} ---\n" + txt.strip())
            doc.close()
            if extracted_pages:
                return "\n\n".join(extracted_pages).strip()
        except Exception as e:
            print(f"[PyMuPDF Error] {e}")

    # 2. PDF Fallback via PyPDF
    if ext == "pdf" and pypdf and not extracted_pages:
        try:
            reader = pypdf.PdfReader(io.BytesIO(content_bytes))
            for idx, page in enumerate(reader.pages, start=1):
                txt = page.extract_text() or ""
                if txt and len(txt.strip()) > 5:
                    extracted_pages.append(f"--- Page {idx} ---\n" + txt.strip())
            if extracted_pages:
                return "\n\n".join(extracted_pages).strip()
        except Exception as e:
            print(f"[PyPDF Error] {e}")

    # 3. Word (.docx)
    if ext in ["docx", "doc"] and docx:
        try:
            doc_file = docx.Document(io.BytesIO(content_bytes))
            lines = [p.text.strip() for p in doc_file.paragraphs if p.text and p.text.strip()]
            if lines:
                return "\n".join(lines).strip()
        except Exception as e:
            print(f"[docx Error] {e}")

    # 4. PowerPoint (.pptx)
    if ext in ["pptx", "ppt"] and pptx:
        try:
            prs = pptx.Presentation(io.BytesIO(content_bytes))
            slides_text = []
            for s_idx, slide in enumerate(prs.slides, start=1):
                s_lines = [p.text.strip() for shape in slide.shapes if shape.has_text_frame for p in shape.text_frame.paragraphs if p.text.strip()]
                if s_lines:
                    slides_text.append(f"--- Slide {s_idx} ---\n" + "\n".join(s_lines))
            if slides_text:
                return "\n\n".join(slides_text).strip()
        except Exception as e:
            print(f"[pptx Error] {e}")

    # 5. Images via EasyOCR
    if ext in ["png", "jpg", "jpeg", "webp", "bmp"]:
        try:
            reader = get_ocr_reader()
            if reader:
                ocr_results = reader.readtext(content_bytes, detail=0)
                if ocr_results:
                    return "\n".join(ocr_results).strip()
        except Exception as e:
            print(f"[EasyOCR Error] {e}")

    # 6. Plain Text files ONLY
    if ext in ["txt", "md", "markdown", "csv", "json", "rtf", "log"] or not ext:
        try:
            decoded = content_bytes.decode("utf-8", errors="ignore").strip()
            if decoded:
                return decoded
        except Exception:
            pass

    return f"--- {file_name} Document Content ---\nExtracted text ready for visual study notes transformation."


# --- Smart Asset Resolution System ---
def detect_primary_subject(title: str, text: str, sections: Optional[List[dict]] = None) -> Tuple[str, str]:
    """Detects primary organ or topic subject dynamically from text content, title, or section headings."""
    combined_content = f"{title or ''} {text or ''} "
    if sections:
        for s in sections:
            if isinstance(s, dict):
                combined_content += f" {s.get('title','')} {s.get('heading','')} {s.get('content','')} "

    content_lower = combined_content.lower()

    organ_map = [
        (["heart", "cardiac", "coronary", "myocard", "cardiovascular", "atrium", "ventricle", "valve", "aorta", "angina"], "heart", "Heart"),
        (["lung", "respiratory", "ards", "pulmonary", "alveol", "pleura", "bronch"], "lungs", "Lungs"),
        (["brain", "neuro", "cerebral", "stroke", "neuron", "cortex", "head"], "brain", "Brain"),
        (["kidney", "renal", "nephro", "dialysis", "glomerul"], "kidney", "Kidney"),
        (["liver", "hepatic", "cirrhosis", "gallbladder", "bile"], "liver", "Liver"),
        (["stomach", "gastric", "gastro", "bowel", "colon", "gut"], "stomach", "Stomach"),
        (["vessel", "vascular", "artery", "vein"], "blood_vessels", "Blood Vessels")
    ]

    for keywords, slug, organ_name in organ_map:
        if any(re.search(r'\b' + re.escape(k) + r'\b', content_lower) for k in keywords):
            return slug, organ_name

    clean_title = (title or "").strip()
    clean_title = re.sub(r'^\d+[\.\)]\s*', '', clean_title)
    clean_title = re.sub(r'[\._\-]+', ' ', clean_title).strip()

    if clean_title and not any(g in clean_title.lower() for g in ["page", "uploaded", "lecture", "document", "notes", "file", "slide"]):
        slug = re.sub(r'[^a-z0-9]', '_', clean_title.lower()).strip('_')
        return slug or "medical_concept", clean_title.title()

    if sections:
        for s in sections:
            if isinstance(s, dict):
                h = (s.get("title") or s.get("heading") or "").strip()
                clean_h = re.sub(r'^\d+[\.\)]\s*', '', h).strip()
                if clean_h and len(clean_h) < 40 and not any(g in clean_h.lower() for g in ["overview", "section", "page"]):
                    slug = re.sub(r'[^a-z0-9]', '_', clean_h.lower()).strip('_')
                    return slug or "medical_concept", clean_h.title()

    return "medical_concept", "Clinical Concept"

def check_generated_asset_cache(subject: str, visual_type: str = "") -> Optional[str]:
    """Checks assets/generated folder for an existing matching generated image file."""
    gen_dir = os.path.join(os.path.dirname(__file__), "assets", "generated")
    if not os.path.exists(gen_dir):
        return None
    clean_subj = re.sub(r'[^a-z0-9]', '', (subject or "").lower())
    clean_type = re.sub(r'[^a-z0-9]', '', (visual_type or "").lower())
    
    # Ignore generic page/uploaded dummy tokens
    if any(g in clean_subj for g in ["page1", "uploadednotes", "manuallecturenotes"]):
        clean_subj = ""

    if not clean_subj:
        return None
    try:
        files = os.listdir(gen_dir)
        if clean_type:
            for f in files:
                if not f.lower().endswith(('.png', '.jpg', '.jpeg', '.webp')):
                    continue
                f_clean = re.sub(r'[^a-z0-9]', '', f.lower())
                if (clean_subj in f_clean or f_clean.startswith(clean_subj)) and clean_type in f_clean:
                    return f"/assets/generated/{f}"
            return None
            
        for f in files:
            if not f.lower().endswith(('.png', '.jpg', '.jpeg', '.webp')):
                continue
            f_clean = re.sub(r'[^a-z0-9]', '', f.lower())
            if clean_subj in f_clean or f_clean.startswith(clean_subj):
                return f"/assets/generated/{f}"
    except Exception as e:
        print(f"[Cache Search Error] {e}")
    return None

def resolve_visual_asset(subject: str, visual_type: str = "primary", is_medication: bool = False, text_context: str = "") -> Optional[str]:
    """Returns None as pre-setup static images have been removed to prioritize dynamic document-specific image generation."""
    return None

def generate_medical_image(prompt: str, filename: str) -> Optional[str]:
    """Calls OpenAI Image API to generate an educational medical illustration and saves to assets/generated/."""
    key = os.getenv("OPENAI_API_KEY", "").strip().strip('"').strip("'")
    if not key or key.startswith("sk-proj-placeholder"):
        print("[Image Gen Notice] No valid OPENAI_API_KEY found.")
        return None
        
    safe_filename = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', filename)
    if not safe_filename.endswith(('.png', '.jpg')):
        safe_filename += ".png"
        
    gen_dir = os.path.join(os.path.dirname(__file__), "assets", "generated")
    os.makedirs(gen_dir, exist_ok=True)
    file_path = os.path.join(gen_dir, safe_filename)
    
    if os.path.exists(file_path):
        return f"/assets/generated/{safe_filename}"
        
    try:
        import openai
        import base64
        import urllib.request
        client = openai.OpenAI(api_key=key, timeout=120.0)
        

        response = None
        for model in ["gpt-image-2"]:
            try:
                response = client.images.generate(
                    model=model,
                    prompt=prompt,
                    n=1,
                    size="1024x1024"
                )
                if response and response.data:
                    print(f"[Image Gen Success via {model}] Generated image for: {filename}")
                    break
            except Exception as e:
                print(f"[OpenAI Image Gen Model Notice ({model})] {e}")
                
        if response and response.data:
            item = response.data[0]
            if getattr(item, 'b64_json', None):
                img_bytes = base64.b64decode(item.b64_json)
                with open(file_path, "wb") as f:
                    f.write(img_bytes)
                print(f"[Image Gen Success] Saved generated image to {file_path}")
                return f"/assets/generated/{safe_filename}"
            elif getattr(item, 'url', None):
                req = urllib.request.Request(item.url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req) as resp, open(file_path, "wb") as f:
                    f.write(resp.read())
                print(f"[Image Gen Success from URL] Saved generated image to {file_path}")
                return f"/assets/generated/{safe_filename}"
    except Exception as e:
        print(f"[Image Gen Exception] {e}")
        
    return None

def extract_visual_requirements(doc_title: str, raw_text: str, sections: Optional[List[dict]] = None) -> List[dict]:
    """Generates structured visual requirements grounded strictly in document content (Anatomy cards only)."""
    doc_subject, doc_organ_name = detect_primary_subject(doc_title, raw_text, sections)
    text_lower = (raw_text or "").lower()
    
    requirements = []
    
    # Anatomy / Structure is the ONLY section that requires an image
    has_anatomy = any("anatomy" in (s.get("heading","") or s.get("title","")).lower() for s in (sections or [])) or "anatomy" in text_lower
    requirements.append({
        "type": "anatomy",
        "subject": f"{doc_organ_name} Anatomy",
        "purpose": f"Show detailed internal anatomy and cross-section structure of {doc_organ_name}",
        "required": True if has_anatomy or doc_organ_name else False
    })
    
    return requirements

def build_image_prompt(requirement: dict, document_context: str = "") -> str:
    """Builds a focused, document-grounded image generation prompt strictly for Anatomy cards."""
    v_type = requirement.get("type", "")
    is_anatomy = requirement.get("is_anatomy") or v_type == "anatomy"
    if not is_anatomy:
        return ""

    subject = requirement.get("subject", "Human Anatomy")
    purpose = requirement.get("purpose", "Show anatomical structure")
    sec_text = requirement.get("section_text", "")
    doc_topic = requirement.get("doc_topic", "")
    
    context_snippet = f" Section Details: {sec_text[:280]}." if sec_text else ""
    if doc_topic and doc_topic.lower() not in (subject or "").lower():
        context_snippet = f" Clinical Context: {doc_topic}.{context_snippet}"

    prompt = (
        f"Medically accurate anatomical textbook diagram of {subject}. "
        f"Purpose: {purpose}.{context_snippet} "
        f"Show the relevant organ and its visible internal anatomical structures, clearly highlighting disease-specific changes and key anatomical features. "
        f"Must include clean callout pointer lines and arrows pointing directly to key visible anatomical points, "
        f"such as affected structures, abnormal tissue, fluid buildup, swelling, narrowing, obstruction, "
        f"plaque, inflammation, or other relevant changes. "
        f"Keep the annotations minimal, sharp, and directly connected to the corresponding anatomy. "
        f"Clean medical textbook diagram illustration, crisp detail, white background."
    )

    return prompt

def resolve_or_generate_visual(requirement: dict, raw_text: str = "", filename: str = "") -> Tuple[Optional[str], bool, bool]:
    """
    Pipeline for Anatomy visuals only:
    1. Check if document-specific generated asset exists in assets/generated/
    2. Call OpenAI API if API key exists to generate NEW document-specific visual and save to assets/generated/{unique_filename}
    3. Fallback to matching cached asset if available
    Returns: (asset_url, generated, reused)
    """
    subject = requirement.get("subject", "")
    v_type = requirement.get("type", "primary")
    is_req = requirement.get("required", True)
    sec_title_low = (requirement.get("section_title") or requirement.get("title") or "").lower()
    
    # Image generation is strictly restricted to Anatomy cards only
    is_anatomy = requirement.get("is_anatomy") or v_type == "anatomy" or "anatomy" in sec_title_low or ("structure" in sec_title_low and not any(k in sec_title_low for k in ["cell", "protect", "neuron"]))
    if not is_req or not is_anatomy:
        return (None, False, False)
        
    clean_fn = os.path.splitext(os.path.basename(filename))[0] if filename else ""
    doc_slug = re.sub(r'[^a-z0-9]', '_', clean_fn.lower()).strip('_') if clean_fn else ""
    doc_slug = re.sub(r'_+', '_', doc_slug)
    
    # Filter out journal metadata stamps or generic dummy tags
    if any(k in doc_slug for k in ["ajphi", "volume", "page_1", "page1", "uploaded_notes", "manual_lecture"]):
        doc_slug = ""
    
    # Clean card / section name
    raw_sec_name = requirement.get("section_title") or requirement.get("title") or ""
    raw_sec_name = re.sub(r'^\d+[\.\)]\s*', '', raw_sec_name).strip()
    sec_slug = re.sub(r'[^a-z0-9]', '_', raw_sec_name.lower()).strip('_') if raw_sec_name else ""
    sec_slug = re.sub(r'_+', '_', sec_slug)
    
    card_name = sec_slug or "anatomy"
    if len(card_name) > 30:
        card_name = card_name[:30].rstrip('_')

    subj_slug = re.sub(r'[^a-z0-9]', '_', (subject or "").lower()).strip('_')
    subj_slug = re.sub(r'_+', '_', subj_slug)
    
    if doc_slug:
        safe_name = f"{doc_slug}_{card_name}.png"
    else:
        prefix = subj_slug[:30] if subj_slug else "anatomy_visual"
        safe_name = f"{prefix}_{card_name}.png"
        
    gen_dir = os.path.join(os.path.dirname(__file__), "assets", "generated")
    file_path = os.path.join(gen_dir, safe_name)
    
    # 1. Check if document-specific generated file exists
    if os.path.exists(file_path):
        return (f"/assets/generated/{safe_name}", False, True)
        
    # 1b. Check if asset exists by visual type or subject in cache
    cached_asset = check_generated_asset_cache(doc_slug or subj_slug or subject, "anatomy")
    if cached_asset:
        return (cached_asset, False, True)
        
    # 2. Generate new image with OpenAI API
    prompt = build_image_prompt(requirement, raw_text)
    if not prompt:
        return (None, False, False)
    gen_url = generate_medical_image(prompt, safe_name)
    if gen_url:
        return (gen_url, True, False)
        
    return (None, False, False)


# --- AI System Prompts ---
DOCUMENT_ANALYSIS_ORIGINAL_NOTE_PROMPT = """You are building the document-analysis engine for a medical/healthcare-focused study notes application.

The user may upload ANY supported document format, including:
* PDF
* DOC/DOCX
* TXT
* Markdown
* Image
* Scanned document
* Screenshot
* Other supported document formats

Your job is to analyze the uploaded material FIRST, understand what it is actually about, extract the important information, summarize it accurately, and then organize the original document content into MUST BE between 6 and 8 intelligent sections.

Do NOT blindly use a fixed template. The uploaded document is the ONLY source of truth.

---

# CORE OBJECTIVE
1. Detect file/content type & extract text / OCR.
2. Understand document: identify primary topic, document type (disease/condition, medication, procedure, lab test, anatomy, nursing topic, physiology, pathology, treatment, clinical concept, etc.).
3. Extract important information matching CONTENT PRIORITIES (Diseases, Medications, Procedures, Labs).
4. Apply PRIORITIZATION LOGIC (topic emphasis, depth of explanation, clinical importance, exam/revision value).
5. Generate MUST HAVE between 6 and 8 dynamic note sections.
6. Return structured JSON with original note content organized into 6 to 8 sections.

---

# STEP 1 — UNDERSTAND THE DOCUMENT FIRST
Internally analyze:
{
  "document_topic": "",
  "document_type": "",
  "primary_subject": "",
  "related_organs": [],
  "related_medications": [],
  "related_procedures": [],
  "related_laboratory_tests": [],
  "major_concepts": [],
  "has_pathophysiology": false,
  "has_treatment": false,
  "has_medication": false,
  "has_diagnostic_information": false
}

---

# STEP 2 — CONTENT PRIORITIES
Look for content matching:
A. DISEASES & CONDITIONS: Definition, Pathophysiology, Causes/Risk factors, Signs & symptoms, Assessment, Diagnostic tests, Labs, Complications, Medical management, Nursing interventions, Patient education, Red flags, Exam tips, Memory aid.
B. MEDICATIONS: Generic name, Brand name (if in doc), Class, Indications, Mechanism of action, Route, Side effects, Adverse effects, Contraindications, Precautions, Interactions, Monitoring, Lab values, Nursing considerations, Patient teaching.
C. PROCEDURES: Purpose, Indications, Preparation, Equipment, Steps, Nursing responsibilities, Monitoring, Complications, Post-care, Teaching, Safety alerts.
D. LABORATORY VALUES: Test name, What it measures, Reference range, Meaning of high/low, Nursing considerations, Urgent findings.

---

# STEP 3 — PRIORITIZATION LOGIC
Do NOT treat all information equally. Focus on topics with the most explanation, detail, repetition, clinical importance, and study value.

---

# STEP 4 — REQUIRED SECTION 1: Overview
Always Section 1: Overview. Answer "What is this document actually about?" in 3 to 5 concise bullet points based on content richness (4-5 points for detailed documents, 3 points for concise ones; MAXIMUM 5 points). Each bullet point MUST contain between 20 and 30 words.

---

# STEP 5 — REQUIRED SECTION 2: Anatomy of the {Organ/Body Part}
Include Section 2: "Anatomy of the {Organ}" ONLY when document relates to an organ/body part (e.g. Heart, Kidney, Liver, Brain, Lungs). Provide 3 to 5 points directly relevant to document depending on detail available (maximum 5 points). Omit if no organ relationship exists.

---

# STEP 6 — PATHOPHYSIOLOGY (CONDITIONAL)
Include Section 3: "Pathophysiology" when document contains biological disease mechanisms, physiological changes, or cause->effect processes. Represent as step-by-step flow (Risk factor -> Physiological change -> Pathological change -> Dysfunction -> Symptoms -> Complications). Do NOT generate or require an image for Pathophysiology as it is presented directly as an icon-based step flowchart in the UI. Omit if missing.

---

# DYNAMIC CARD POINT COUNT (3 TO 5 POINTS PER CARD - HIGHEST 5)
Each section card MUST dynamically contain between 3 and 5 items/bullet points based on the amount of information in the document for that topic. NEVER fix every card to 3 points. If a section has rich detailed information in the uploaded document, provide 4 or 5 points (maximum 5 points). If concise, provide 3 points (minimum 3 points).

---

# STRICT CONSTRAINT: MUST CONTAIN 6 TO 8 SECTIONS
The 'sections' array MUST contain between 6 and 8 sections. Never less than 6 (if info permits), never more than 8. Do NOT create empty sections. Do NOT invent fake information. Preserve numbers, lab values, drug names, and clinical meaning.

---

# OUTPUT JSON SCHEMA
Return valid JSON ONLY matching:
{
  "document_analysis": {
    "title": "Document Title",
    "topic": "Primary Topic",
    "document_type": "disease | medication | procedure | lab | anatomy | nursing_topic | general",
    "primary_subject": "Primary Subject",
    "organ": "Organ Name or empty",
    "has_pathophysiology": true,
    "key_terms_focused": ["Term1", "Term2", "Term3"]
  },
  "sections": [
    {
      "section_number": 1,
      "title": "Overview",
      "type": "overview",
      "priority_source": "General Overview",
      "content": [
        "Point 1: Main topic definition and primary biological mechanisms.",
        "Point 2: Key clinical risk factors and population prevalence indicators.",
        "Point 3: Primary diagnostic features and hallmark symptom presentation.",
        "Point 4: High-priority therapeutic goals and patient safety considerations.",
        "Point 5: Critical nursing interventions and long-term care management."
      ],
      "visual": {
        "required": false,
        "type": "organ",
        "subject": "Organ Name",
        "purpose": "Overview anatomical visualization"
      }
    },
    {
      "section_number": 2,
      "title": "Anatomy / Core Structure",
      "type": "anatomy",
      "priority_source": "Anatomy",
      "content": [
        "Point 1: Primary organ anatomy and vascular supply.",
        "Point 2: Microscopic cellular structure and tissue arrangement.",
        "Point 3: Functional anatomical zones and physiological roles.",
        "Point 4: Surrounding structural landmarks and innervation."
      ],
      "visual": {
        "required": true,
        "type": "anatomy",
        "subject": "Anatomy",
        "purpose": "Anatomical cross section"
      }
    },
    {
      "section_number": 3,
      "title": "Pathophysiology",
      "type": "pathophysiology",
      "priority_source": "Pathophysiology",
      "flow": [
        { "step": 1, "title": "Insult / Trigger", "description": "Details" },
        { "step": 2, "title": "Pathological Response", "description": "Details" },
        { "step": 3, "title": "Tissue Dysfunction", "description": "Details" },
        { "step": 4, "title": "Clinical Complications", "description": "Details" }
      ],
      "visual": {
        "required": false,
        "type": "none",
        "subject": "",
        "purpose": "Presented via icon-based step flowchart; no image required."
      }
    },
    {
      "section_number": 4,
      "title": "Signs & Symptoms",
      "type": "dynamic",
      "priority_source": "Diseases and Conditions",
      "content": [
        "Point 1: Primary cardinal symptom and early clinical signs.",
        "Point 2: Secondary systemic manifestations and lab alerts.",
        "Point 3: Late-stage progression indicators and severe red flags.",
        "Point 4: Differential diagnostic symptoms and physical exam findings.",
        "Point 5: Patient-reported subjective symptoms and functional impact."
      ],
      "visual": {
        "required": false,
        "type": "diagram",
        "subject": "Symptoms",
        "purpose": "Visual aid"
      }
    }
  ]
}
"""

MEDICAL_STUDY_NOTES_ARCHITECT_PROMPT = """You are Nursing Study Sheet Architect, an educational content and visual-layout assistant for nursing students.

Your job is to transform raw study material into a clear, accurate, visually structured nursing study sheet suitable for rendering as a mobile study card, printable infographic, PNG, or PDF.

Your responsibilities are to:
1. Understand and organize the uploaded material.
2. Preserve the meaning of the source notes.
3. Correct obvious spelling, grammar, formatting, and OCR errors.
4. Identify the most educationally important concepts.
5. Convert the content into concise nursing-student language.
6. Create a logical visual hierarchy.
7. Recommend suitable icons, diagrams, flowcharts, tables, and illustrations.
8. Produce structured output that a rendering engine can reliably convert into an infographic.
9. Clearly distinguish source-supported content from supplemental educational context.
10. Avoid inventing unsupported medical facts.

PRIMARY GOAL:
Create a high-yield nursing study sheet that is accurate, easy to scan, visually organized, useful for exams/clinical review, appropriate for nursing students, faithful to uploaded notes, and safe for educational use.

SUMMARY & CARD CONSTRAINTS:
1. SUMMARIZE THE DOCUMENT FIRST: Populate the "summary" field with a 1-sentence executive overview ("one_sentence_overview") and 3 to 6 high-yield core points ("high_yield_points").
2. MAXIMUM 8 PRIMARY SECTION CARDS: Return a MAXIMUM of 8 primary section cards in the "sections" array (highest 8 cards).
3. DYNAMIC POINTS PER CARD (3 TO 5 POINTS): Each card MUST dynamically contain between 3 and 5 items/points based on document detail (never fix to 3 points for all cards). Minimum 3 points, maximum 5 points per card.
4. CONTENT PRIORITIES: Organize sections using high-yield categories such as:
   - Disease/Condition: Definition/Overview, Pathophysiology, Causes/Risk Factors, Signs & Symptoms, Diagnostic Workup/Labs, Complications, Medical Management, Nursing Interventions, Patient Education, Red Flags.
   - Medications: Generic/Brand Name, Class, Indications, Mechanism, Route, Side Effects/Adverse Effects, Contraindications, Nursing Considerations, Teaching.
   - Procedures: Purpose, Indications, Preparation, Main Steps, Nursing Responsibilities, Complications.
   - Lab Values: Test Name, Measurement, Reference Range, High/Low Meanings, Nursing Considerations.

OUTPUT REQUIREMENTS:
Return valid JSON ONLY matching this exact schema:

{
  "document": {
    "title": "string",
    "subtitle": "string",
    "topic_type": "medication | disease | procedure | lab | comparison | mixed",
    "student_level": "beginner | intermediate | advanced | unspecified",
    "layout_type": "medication_card | disease_overview | comparison_chart | procedure_guide | lab_reference | concept_map | multi_page_study_pack",
    "page_size": "mobile_portrait | a4_portrait | letter_portrait | landscape",
    "estimated_pages": 1,
    "educational_disclaimer": "Educational study support material. Follow facility policy and provider orders."
  },
  "summary": {
    "one_sentence_overview": "string (no more than 35 words)",
    "high_yield_points": [
      "string (3 to 6 high-yield bullet points)"
    ]
  },
  "sections": [
    {
      "id": "section_1",
      "heading": "SECTION TITLE IN UPPERCASE",
      "title": "SECTION TITLE IN UPPERCASE",
      "section_type": "tagged_items | flowchart",
      "type": "tagged_items | flowchart",
      "priority": "essential | important | supplemental",
      "origin": "source | supplemental | inferred",
      "items": [
        {
          "label": "Badge / Subhead Title",
          "text": "Concise key point text",
          "emphasis": "normal | key | warning | red_flag",
          "origin": "source | supplemental | inferred"
        }
      ],
      "flows": [
        {
          "label": "PATHWAY NAME",
          "steps": [
            { "name": "STEP TITLE", "subtext": "Step details" }
          ]
        }
      ]
    }
  ],
  "exam_support": {
    "memory_aid": {
      "title": "string",
      "content": "string",
      "is_source_supported": true
    },
    "exam_tips": [
      {
        "tip": "string",
        "origin": "source | supplemental | inferred"
      }
    ]
  },
  "safety_panel": {
    "red_flags": [
      "string"
    ],
    "medication_safety_notes": [
      "string"
    ]
  }
}
"""

SYSTEM_PROMPT = MEDICAL_STUDY_NOTES_ARCHITECT_PROMPT



# --- Deterministic Parser (1:1 Source Fidelity) ---
def clean_raw_text_metadata(raw_text: str) -> str:
    """Removes raw PDF journal headers, volume stamps, author names, repetitive title prefixes, and reference lists."""
    if not raw_text:
        return ""
    
    # Remove journal volume stamps and repetitive metadata lines from text body
    text = re.sub(r'AJPHI\s*[|I]\s*VOLUME\s*\d+\s*[|I]\s*\d{4}\s*(?:ORIGINAL ARTICLE)?', '', raw_text, flags=re.I)
    text = re.sub(r'The\s+American\s+Journal\s+of\s+Patient\s+Health\s+Info\s*:\s*\d{4}', '', text, flags=re.I)
    text = re.sub(r'AJPHI\s+I\s+VOLUME\s+\d+\s+I\s+\d{4}', '', text, flags=re.I)

    lines = text.split('\n')
    clean_lines = []
    in_references = False
    for line in lines:
        s = line.strip()
        if not s:
            continue
        if re.search(r'\b(AJPHI|VOLUME\s+\d+|ORIGINAL ARTICLE|Journal of Patient Health|ISSN|DOI|Published by|Available from:)\b', s, re.I):
            continue
        if re.search(r'\b(MD\s+[a-z]|University of|Hospital|Department of|Faculty of)\b', s, re.I) and len(s) < 140:
            continue
        if re.match(r'^(?:References|BIBLIOGRAPHY|Citations)\b', s, re.I):
            in_references = True
            continue
        if in_references:
            continue
        # Remove trailing/leading bullet junk or repetitive title prefixes
        s = re.sub(r'^[A-Z0-9\s_\-]{15,80}\s*OVERVIEW\s*:\s*', '', s, flags=re.I)
        clean_lines.append(s)
    return "\n".join(clean_lines)

def clean_heading_title(line: str) -> str:
    return re.sub(r'^(?:[0-9]+[\.\)]|#{1,6}\s+|[IVXLCDM]+[\.\)])\s*', '', line).strip(':').strip()

def is_major_heading(line: str) -> bool:
    """Accurately detects true top-level section headers including domain title matches."""
    stripped = line.strip()
    if not stripped:
        return False

    if re.match(r'^(?:[0-9]+[\.\)]|#{1,6}\s+|[IVXLCDM]+[\.\)]|[A-Z][\.\)])\s+[A-Za-z]', stripped):
        return True

    clean = re.sub(r'^[•\-\*\→●▪■◆➢►○✔✓]\s*', '', stripped)
    clean = re.sub(r'^(?:[0-9]+[\.\)]|#{1,6}\s+|[IVXLCDM]+[\.\)])\s*', '', clean).strip(':').strip()
    clean_lower = clean.lower()

    if not clean or '→' in clean or '->' in clean or len(clean) > 75:
        return False

    if clean.endswith(('.', '?', ';')) and len(clean.split()) > 4:
        return False

    heading_keywords = [
        "overview", "definition", "introduction", "background", "demystified",
        "anatomy", "structure", "physiology",
        "pathophysiology", "pathway", "mechanism", "pathogenesis",
        "risk factor", "cause", "etiology", "risk factors", "modifiable risk", "non-modifiable risk",
        "clinical presentation", "manifestation", "signs & symptoms", "signs and symptoms", "symptom", "discomfort",
        "diagnostic", "diagnosis", "test", "lab", "evaluation", "assessment", "angiogram", "catheterization", "echocardiogram",
        "treatment", "management", "pharmacology", "medication", "drug", "antiplatelet", "statin", "surgical", "procedure",
        "nursing", "intervention", "patient education", "complication", "prevention", "red flag"
    ]

    for th in heading_keywords:
        if th in clean_lower and len(clean_lower) < 65:
            return True

    if re.match(r'^(?:SECTION|PART|CHAPTER|UNIT)\s+\d+', clean, re.I):
        return True

    return False


def parse_dynamic_sections_from_text(raw_text: str, topic_name: str = "") -> List[dict]:
    """Parses text into dynamic 1:1 section blocks preserving exact document headings and sub-bullets."""
    if not raw_text or not raw_text.strip():
        return []

    lines = [l.strip() for l in raw_text.split('\n') if l.strip()]
    if not lines:
        return []

    def clean_and_join_lines(raw_lines: List[str]) -> List[str]:
        isolated_glyphs = {'i', 'g', 'n', 'l', 'o', 'q', 'v', '•', '●', '■', '▪', '◆', '➢', '✓', '✔', '►', '○', '—', '–', '-', '*', ''}
        normalized = []
        pending_bullet = False

        for l in raw_lines:
            if not l:
                continue
            s = l.strip()
            if s.startswith('--- Page') or s.startswith('--- Slide'):
                continue
            s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\uf000-\uf8ff]', '', s).strip()
            if not s:
                continue

            if s.lower() in isolated_glyphs or (len(s) == 1 and not s.isalnum()):
                pending_bullet = True
                continue

            has_bullet = bool(re.match(r'^[•\-\*\→●▪■◆➢►○✔✓]', s)) or bool(re.match(r'^[IGnl]\s+(?=[A-Za-z0-9])', s))
            s = re.sub(r'^[•\-\*\→●▪■◆➢►○✔✓]\s*', '', s)
            s = re.sub(r'^[IGnl]\s+(?=[A-Za-z0-9])', '', s).strip()
            if not s:
                continue

            is_bullet = pending_bullet or has_bullet
            pending_bullet = False
            normalized.append((s, is_bullet))

        joined = []
        for text, is_bullet in normalized:
            is_numbered = bool(re.match(r'^(?:[0-9]+[\.\)]|#{1,6}\s+|[IVXLCDM]+[\.\)])\s+[A-Za-z]', text))
            is_arrow = '→' in text or '->' in text
            is_subhead = bool(re.match(r'^[A-Za-z0-9\s&/()\-–]+:\s*$', text))

            if joined and not is_bullet and not is_numbered and not is_arrow and not is_subhead:
                prev = joined[-1]
                prev_is_numbered = bool(re.match(r'^(?:[0-9]+[\.\)]|#{1,6}\s+|[IVXLCDM]+[\.\)])\s+[A-Za-z]', prev))
                prev_is_subhead = prev.endswith(':')

                if not prev_is_numbered and not prev_is_subhead and not is_major_heading(text):
                    # Only join if text starts with lowercase (sentence continuation)
                    if text[0].islower():
                        joined[-1] = f"{prev} {text}"
                        continue
            joined.append(text)
        return joined

    joined_lines = clean_and_join_lines(lines)
    raw_sections = []
    current_sec = None

    for idx, line in enumerate(joined_lines):
        # Skip leading title if redundant
        if idx == 0 and not re.match(r'^\d+[\.\)]', line) and ('study notes' in line.lower() or 'lecture notes' in line.lower()):
            continue

        if is_major_heading(line):
            if current_sec and (current_sec['heading'] or current_sec['lines'] or current_sec['items'] or current_sec['flows']):
                raw_sections.append(current_sec)
            clean_title = clean_heading_title(line)
            current_sec = {
                "heading": clean_title,
                "lines": [],
                "current_sublabel": "",
                "items": [],
                "bullets": [],
                "flows": []
            }
        else:
            if not current_sec:
                current_sec = {
                    "heading": f"{topic_name or 'Document'} Overview",
                    "lines": [],
                    "current_sublabel": "",
                    "items": [],
                    "bullets": [],
                    "flows": []
                }

            current_sec["lines"].append(line)

            # Sequential flowchart detection
            if '→' in line or '->' in line:
                pathway_label = current_sec["current_sublabel"] or f"{current_sec['heading']} Pathway"
                steps_raw = [s.strip() for s in re.split(r'→|->', line) if s.strip()]
                if len(steps_raw) >= 2:
                    flow_steps = []
                    for st in steps_raw:
                        st_clean = re.sub(r'^[•\-\*\d\.\)\→●▪■◆➢✓✔]\s*', '', st).strip()
                        parts = re.split(r':|\s+[–—\-]\s+|[–—]', st_clean, maxsplit=1)
                        if len(parts) == 2 and len(parts[0].strip()) < 30 and len(parts[1].strip()) > 1:
                            flow_steps.append({"name": parts[0].strip().upper(), "subtext": parts[1].strip()})
                        else:
                            flow_steps.append({"name": st_clean.upper(), "subtext": st_clean})

                    clean_flow_label = re.sub(r'\bpathway\b', '', pathway_label, flags=re.I).strip() + " PATHWAY"
                    current_sec["flows"].append({
                        "label": clean_flow_label.upper(),
                        "steps": flow_steps
                    })
                    current_sec["bullets"].append(line)
                    continue

            # Subheader detection (e.g. "Has 4 chambers:", "Major blood vessels:")
            if line.endswith(':') and len(line) < 50:
                sub_title = line.strip(':').strip()
                current_sec["current_sublabel"] = sub_title
                continue

            lbl = current_sec["current_sublabel"]
            # Split only on colon, en/em dash, or hyphen surrounded by spaces (preserving compound words like heart-healthy)
            parts = re.split(r':|\s+[–—\-]\s+|[–—]', line, maxsplit=1)

            if len(parts) == 2 and len(parts[0].strip()) < 40 and len(parts[1].strip()) >= 1 and not parts[1].strip().startswith(('http', '//')):
                item_label = parts[0].strip()
                item_text = parts[1].strip()
                current_sec["items"].append({"label": item_label, "text": item_text, "tag": item_label})
                current_sec["bullets"].append(f"{item_label}: {item_text}")
            else:
                badge_lbl = lbl if lbl else current_sec["heading"]
                current_sec["items"].append({"label": badge_lbl, "text": line, "tag": badge_lbl})
                current_sec["bullets"].append(line)

    if current_sec and (current_sec['heading'] or current_sec['lines'] or current_sec['items'] or current_sec['flows']):
        raw_sections.append(current_sec)

    valid_sections = [s for s in raw_sections if s["heading"] or s["lines"] or s["items"] or s["flows"]]
    accents = ["blue", "red", "purple", "green", "amber"]
    sections = []

    for idx, rsec in enumerate(valid_sections, start=1):
        heading = rsec["heading"] or f"Section {idx}"
        sec_lines = rsec["lines"]
        bullets = rsec["bullets"] if rsec["bullets"] else sec_lines
        items = rsec["items"] if rsec["items"] else [{"label": heading, "text": b} for b in bullets]
        flows = rsec["flows"]

        heading_lower = heading.lower()
        is_non_pathway = any(k in heading_lower for k in ["cause", "diagnostic", "nursing", "intervention", "medication", "drug", "complication", "overview"])
        is_stage_or_phase = any(k in heading_lower for k in ["stage", "phase"])
        is_patho = (not is_non_pathway and any(k in heading_lower for k in ["patho", "pathway", "flow", "circulation", "process", "mechanism"])) or bool(flows)
        
        if is_patho or flows:
            sec_type = "flowchart"
            if not flows:
                flow_steps = []
                raw_source = bullets if bullets else (sec_lines if sec_lines else [it.get("text", "") for it in items if isinstance(it, dict)])
                for st in raw_source:
                    st_clean = re.sub(r'^[•\-\*\d\.\)\→●▪■◆➢✓✔]\s*', '', str(st)).strip()
                    if not st_clean:
                        continue
                    parts = re.split(r':|\s+[–—\-]\s+|[–—]', st_clean, maxsplit=1)
                    if len(parts) == 2 and len(parts[0].strip()) < 30 and len(parts[1].strip()) > 1:
                        flow_steps.append({"name": parts[0].strip().upper(), "subtext": parts[1].strip()})
                    else:
                        flow_steps.append({"name": st_clean.upper(), "subtext": st_clean})
                if flow_steps:
                    flows = [{
                        "label": f"{heading.upper()} PATHWAY",
                        "steps": flow_steps
                    }]
        elif is_non_pathway or is_stage_or_phase:
            sec_type = "tagged_items"
        else:
            sec_type = "flowchart" if (flows and len(flows) > 0) else "tagged_items"

        sections.append({
            "id": f"section_{idx}",
            "title": f"{idx}. {heading.upper()}" if not re.match(r'^\d+\.', heading) else heading.upper(),
            "heading": heading.upper(),
            "section_type": sec_type,
            "type": sec_type,
            "content": "\n".join(sec_lines),
            "bullets": bullets,
            "items": items,
            "flows": flows,
            "accent": accents[(idx - 1) % len(accents)]
        })

    return sections


def sanitize_and_group_sections(sections: List[dict]) -> List[dict]:
    """Consolidates duplicate topic cards into single cards preserving exact document titles."""
    if not sections or not isinstance(sections, list):
        return []

    accents_cycle = ["blue", "red", "purple", "green", "amber"]
    grouped: Dict[str, dict] = {}
    ordered_keys = []

    for sec in [s for s in sections if isinstance(s, dict)]:
        raw_title = (sec.get("title") or sec.get("heading") or "").strip()
        if not raw_title:
            continue
        clean_title = re.sub(r'^\d+[\.\)]\s*', '', raw_title).strip()
        canonical_title = clean_title.upper()
        if canonical_title.strip() in ["OVERVIEW", "DOCUMENT OVERVIEW", "TOPIC OVERVIEW"]:
            canonical_title = "OVERVIEW"

        if canonical_title not in grouped:
            sec_type = sec.get("type") or sec.get("section_type") or "tagged_items"
            sec_img = sec.get("image") or sec.get("image_url") or (sec.get("visual", {}).get("image_url") if isinstance(sec.get("visual"), dict) else "") or ""
            grouped[canonical_title] = {
                "id": sec.get("id") or f"section_{uuid.uuid4().hex[:6]}",
                "title": canonical_title,
                "heading": canonical_title,
                "type": sec_type,
                "section_type": sec_type,
                "content": sec.get("content") or "",
                "bullets": [],
                "items": [],
                "flows": list(sec.get("flows") or []),
                "accent": sec.get("accent") or accents_cycle[len(grouped) % len(accents_cycle)],
                "image": sec_img,
                "image_url": sec_img,
                "visual": sec.get("visual") or None
            }
            if sec.get("isAnatomyCard"):
                grouped[canonical_title]["isAnatomyCard"] = True
            if sec.get("organ_type"):
                grouped[canonical_title]["organ_type"] = sec.get("organ_type")
            ordered_keys.append(canonical_title)
        else:
            sec_img = sec.get("image") or sec.get("image_url") or (sec.get("visual", {}).get("image_url") if isinstance(sec.get("visual"), dict) else "") or ""
            if not grouped[canonical_title].get("image") and sec_img:
                grouped[canonical_title]["image"] = sec_img
                grouped[canonical_title]["image_url"] = sec_img
                if sec.get("visual"):
                    grouped[canonical_title]["visual"] = sec.get("visual")

        card = grouped[canonical_title]
        title_lower = canonical_title.lower()
        is_non_pathway = any(k in title_lower for k in ["cause", "diagnostic", "nursing", "intervention", "medication", "drug", "complication", "overview"])
        is_stage_title = any(k in title_lower for k in ["stage", "phase"])
        is_patho_title = not is_stage_title and any(k in title_lower for k in ["patho", "pathway", "flow", "circulation", "process", "mechanism"])
        if is_patho_title or (not is_non_pathway and not is_stage_title and (sec.get("type") == "flowchart" or sec.get("flows"))):
            card["type"] = "flowchart"
            card["section_type"] = "flowchart"
            for fl in (sec.get("flows") or []):
                if fl not in card["flows"]:
                    card["flows"].append(fl)
        elif is_non_pathway or is_stage_title:
            card["type"] = "tagged_items"
            card["section_type"] = "tagged_items"
            if sec.get("flows"):
                card["flows"] = list(sec.get("flows"))

        existing_item_tuples = {((it.get("label") or "").strip().lower(), (it.get("text") or "").strip().lower()) for it in card["items"]}
        existing_bullets = {b.strip().lower() for b in card["bullets"] if isinstance(b, str)}

        for it in (sec.get("items") or []):
            it_copy = dict(it)
            lbl = (it_copy.get("label") or "").strip() or clean_title.upper()
            txt = (it_copy.get("text") or "").strip()
            if lbl.lower() in ["overview", "subheader", "key point"]:
                lbl = clean_title.upper()
                it_copy["label"] = lbl
            key_tuple = (lbl.lower(), txt.lower())
            if key_tuple not in existing_item_tuples:
                card["items"].append(it_copy)
                existing_item_tuples.add(key_tuple)
            if txt and txt.lower() not in existing_bullets:
                card["bullets"].append(txt)
                existing_bullets.add(txt.lower())

        for b in (sec.get("bullets") or []):
            if isinstance(b, str) and b.strip() and b.strip().lower() not in existing_bullets:
                b_clean = b.strip()
                card["bullets"].append(b_clean)
                existing_bullets.add(b_clean.lower())
                # Auto-sync bullets into items only if not already represented in card items
                if not any(it.get("text", "").strip().lower() == b_clean.lower() or b_clean.lower() in it.get("text", "").strip().lower() or it.get("text", "").strip().lower() in b_clean.lower() for it in card["items"]):
                    parts = re.split(r':|\s+[–—\-]\s+', b_clean, maxsplit=1)
                    if len(parts) == 2 and 2 < len(parts[0].strip()) < 35 and len(parts[1].strip()) > 1:
                        lbl_cand = parts[0].strip().title()
                        txt_cand = parts[1].strip()
                    else:
                        lbl_cand = clean_title.title()
                        txt_cand = b_clean
                    card["items"].append({
                        "label": lbl_cand,
                        "text": txt_cand
                    })
                    existing_item_tuples.add((lbl_cand.lower(), txt_cand.lower()))

    # Filter out empty cards (cards with no items, no bullets, and no flows)
    valid_cards = []
    for k in ordered_keys:
        card = grouped[k]
        has_items = bool(card.get("items") and len(card["items"]) > 0)
        has_bullets = bool(card.get("bullets") and len(card["bullets"]) > 0)
        has_flows = bool(card.get("flows") and len(card["flows"]) > 0)
        if has_items or has_bullets or has_flows or card.get("isAnatomyCard"):
            valid_cards.append(card)
    return valid_cards


def build_modular_visual_dashboard_schema(filename: str, raw_text: str, parsed_ai: Optional[dict] = None) -> dict:
    """Constructs 100% source-faithful Original Notes & Visual Notes schema with synchronized layout and assets."""
    topic_name = filename.rsplit('.', 1)[0] if '.' in filename else filename
    topic_name = topic_name.replace("_", " ").replace("-", " ").title()

    doc_subject, doc_organ_name = detect_primary_subject(topic_name, raw_text)
    dynamic_sections = parse_dynamic_sections_from_text(raw_text, topic_name)

    # Use AI-parsed / Original Note sections directly when present to ensure 1:1 card mapping
    if parsed_ai and isinstance(parsed_ai, dict) and "sections" in parsed_ai:
        sections_to_use = list(parsed_ai.get("sections") or [])
    else:
        sections_to_use = dynamic_sections

    # Check Anatomy rule
    for sec in sections_to_use:
        t_low = (sec.get("title", "") or sec.get("heading", "")).lower()
        if "anatomy" in t_low or ("structure" in t_low and not any(k in t_low for k in ["cell", "protect", "neuron"])):
            sec["isAnatomyCard"] = True
            break

    # Organize document-derived sections into visual dashboard layout order
    overview_card, anatomy_card = None, None
    pathway_cards, other_cards = [], []

    for idx, sec in enumerate(sections_to_use):
        if not isinstance(sec, dict):
            continue
        t_low = (sec.get("title", "") or sec.get("heading", "")).lower()
        is_non_path = any(k in t_low for k in ["cause", "diagnostic", "nursing", "intervention", "medication", "drug", "complication"])
        is_stage = any(k in t_low for k in ["stage", "phase"])
        is_path = not is_non_path and not is_stage and (sec.get("type") == "flowchart" or sec.get("flows") or any(k in t_low for k in ["pathway", "flow", "circulation", "patho", "process", "mechanism"]))
        is_anat = sec.get("isAnatomyCard") or "anatomy" in t_low
        is_over = ("overview" in t_low or "definition" in t_low or (idx == 0 and not is_non_path and not is_stage and not is_path and not is_anat)) and not is_anat and not is_path and not is_stage

        if is_over and not overview_card:
            overview_card = sec
        elif is_anat and not anatomy_card:
            anatomy_card = sec
        elif is_path:
            pathway_cards.append(sec)
        else:
            other_cards.append(sec)

    # 1. Sanitize Overview card items to keep Overview card clean & concise (up to 5 items max)
    if overview_card:
        leakage_labels = {"pathophysiology", "patho", "diagnostic", "nursing", "intervention", "treatment", "medication", "cause", "causes", "stage", "stages"}
        
        clean_overview_items = []
        for it in overview_card.get("items", []):
            if isinstance(it, dict):
                lbl = (it.get("label") or "").strip().lower()
                is_leaked = any(kl in lbl for kl in leakage_labels)
                if not is_leaked:
                    clean_overview_items.append(it)
        if clean_overview_items:
            overview_card["items"] = clean_overview_items[:5]
        else:
            dyn_overview = next((ds for ds in dynamic_sections if "overview" in (ds.get("heading") or "").lower()), None)
            if dyn_overview and dyn_overview.get("items"):
                overview_card["items"] = [it for it in dyn_overview["items"] if isinstance(it, dict) and not any(kl in (it.get("label") or "").lower() for kl in leakage_labels)][:5]

    # 2. Extract step-by-step flows for Pathway cards directly from the document's real text
    for psec in pathway_cards:
        if not psec.get("flows") or not psec["flows"] or not psec["flows"][0].get("steps"):
            dyn_patho = next((ds for ds in dynamic_sections if any(k in (ds.get("heading") or "").lower() for k in ["patho", "pathway", "flow", "process", "mechanism"])), None)
            raw_source = psec.get("items") or psec.get("bullets") or psec.get("lines") or (dyn_patho.get("items") if dyn_patho else []) or (dyn_patho.get("bullets") if dyn_patho else []) or []
            flow_steps = []
            for st in raw_source:
                if isinstance(st, dict):
                    lbl = (st.get("label") or st.get("tag") or "").strip()
                    txt = (st.get("text") or st.get("desc") or "").strip()
                    if lbl and txt and lbl.lower() not in ["pathophysiology", "patho", "overview"]:
                        flow_steps.append({"name": lbl.upper(), "subtext": txt})
                    elif txt:
                        parts = re.split(r':|\s+[–—\-]\s+|[–—]', txt, maxsplit=1)
                        if len(parts) == 2 and len(parts[0].strip()) < 30 and len(parts[1].strip()) > 1:
                            flow_steps.append({"name": parts[0].strip().upper(), "subtext": parts[1].strip()})
                        else:
                            flow_steps.append({"name": txt.upper()[:30], "subtext": txt})
                elif isinstance(st, str) and st.strip():
                    st_clean = st.strip()
                    parts = re.split(r':|\s+[–—\-]\s+|[–—]', st_clean, maxsplit=1)
                    if len(parts) == 2 and len(parts[0].strip()) < 30 and len(parts[1].strip()) > 1:
                        flow_steps.append({"name": parts[0].strip().upper(), "subtext": parts[1].strip()})
                    else:
                        flow_steps.append({"name": st_clean.upper()[:30], "subtext": st_clean})
            if flow_steps:
                sec_title_clean = re.sub(r'^\d+[\.\)]\s*', '', (psec.get('title') or psec.get('heading') or 'PATHOPHYSIOLOGY')).strip().upper()
                psec["type"] = "flowchart"
                psec["section_type"] = "flowchart"
                psec["flows"] = [{
                    "label": f"{sec_title_clean} PATHWAY",
                    "steps": flow_steps
                }]

    ordered = [s for s in [overview_card, anatomy_card] if s and isinstance(s, dict)]
    ordered.extend([s for s in pathway_cards if s and isinstance(s, dict)])
    ordered.extend([s for s in other_cards if s and isinstance(s, dict)])
    # Cap ordered sections to maximum 8 cards BEFORE generating images
    ordered = ordered[:8]

    # Structured Visual Requirements & Pipeline Asset Resolution for Cards
    for idx, sec in enumerate(ordered, start=1):
        if not sec or not isinstance(sec, dict):
            continue
        t_low = (sec.get("title", "") or sec.get("heading", "")).lower()
        sec_title = sec.get("title") or sec.get("heading") or f"Section {idx}"
        
        # Extract snippet of section content to ground the image prompt
        sec_text_lines = []
        for it in (sec.get("items") or []):
            if isinstance(it, dict) and it.get("text"):
                sec_text_lines.append(it["text"])
            elif isinstance(it, str):
                sec_text_lines.append(it)
        for b in (sec.get("bullets") or []):
            if isinstance(b, str):
                sec_text_lines.append(b)
        sec_text_snippet = " ".join(sec_text_lines[:4])

        # Strict rule: Only Anatomy card gets image generation; all other cards require no images
        is_anatomy_card = bool(sec.get("isAnatomyCard") or "anatomy" in t_low or ("structure" in t_low and not any(k in t_low for k in ["cell", "protect", "neuron"])))

        if is_anatomy_card:
            subj = f"{doc_organ_name} Anatomy"
            purpose = f"Show detailed internal anatomy cross-section of {doc_organ_name}"
            req = {
                "type": "anatomy",
                "is_anatomy": True,
                "subject": subj,
                "purpose": purpose,
                "doc_topic": topic_name,
                "section_title": sec_title,
                "section_text": sec_text_snippet,
                "required": True
            }

            existing_img = sec.get("image") or sec.get("image_url") or (sec.get("visual") or {}).get("image_url")
            if existing_img:
                sec_img = existing_img
            else:
                sec_img, _, _ = resolve_or_generate_visual(req, raw_text=raw_text, filename=filename)

            sec["visual"] = {
                "required": True,
                "type": "anatomy",
                "subject": subj,
                "purpose": purpose,
                "image_url": sec_img or ""
            }
            sec["image"] = sec_img or ""
            sec["image_url"] = sec_img or ""
        else:
            sec["visual"] = {
                "required": False,
                "type": "none",
                "subject": "",
                "purpose": "No image required",
                "image_url": ""
            }
            sec["image"] = ""
            sec["image_url"] = ""

    clean_sections = sanitize_and_group_sections(ordered)
    # Cap sections array to maximum 8 cards as required by prompt specification
    clean_sections = clean_sections[:8]

    bottom_panels = {
        "clinical_tip": {"title": "Clinical Tip", "text": f"Monitor hemodynamic status and organ perfusion parameters when assessing {doc_organ_name.lower()} function."},
        "remember_mnemonic": {"title": "Remember", "text": f"Recall the sequential blood flow and anatomical landmarks of the {doc_organ_name.lower()} during patient evaluations."}
    }
    if parsed_ai and isinstance(parsed_ai, dict) and "bottom_panels" in parsed_ai:
        if isinstance(parsed_ai["bottom_panels"], dict):
            bottom_panels.update(parsed_ai["bottom_panels"])

    summary_obj = None
    if parsed_ai and isinstance(parsed_ai, dict) and "summary" in parsed_ai:
        summary_obj = parsed_ai["summary"]

    if not summary_obj or not isinstance(summary_obj, dict):
        overview_sentence = f"High-yield study sheet summarizing {topic_name} core pathophysiological mechanisms, diagnostic indicators, and essential nursing interventions."
        high_yield = []
        if dynamic_sections:
            for ds in dynamic_sections[:5]:
                heading_txt = (ds.get("heading") or ds.get("title") or "").title()
                if ds.get("items"):
                    it_sample = ds["items"][0]
                    if isinstance(it_sample, dict) and it_sample.get("text"):
                        high_yield.append(f"{heading_txt}: {it_sample['text'][:120]}")
                elif ds.get("bullets"):
                    b_sample = ds["bullets"][0]
                    if isinstance(b_sample, str) and len(b_sample) > 5:
                        high_yield.append(f"{heading_txt}: {b_sample[:120]}")
        if not high_yield:
            high_yield = [
                f"Key clinical concepts and anatomical structures of the {doc_organ_name}.",
                "Diagnostic workup and high-priority nursing assessments.",
                "Pharmacological management, interventions, and safety red flags."
            ]
        summary_obj = {
            "one_sentence_overview": overview_sentence,
            "high_yield_points": high_yield[:5]
        }

    exam_support = (parsed_ai.get("exam_support") if parsed_ai and isinstance(parsed_ai, dict) else None) or {
        "memory_aid": {"title": f"{doc_organ_name.upper()} Assessment", "content": f"Recall key anatomical landmarks and monitoring priorities for {topic_name}.", "is_source_supported": True},
        "exam_tips": [{"tip": "Focus on diagnostic findings and high-priority nursing interventions.", "origin": "source"}]
    }

    safety_panel = (parsed_ai.get("safety_panel") if parsed_ai and isinstance(parsed_ai, dict) else None) or {
        "red_flags": [f"Monitor for acute distress or sudden changes in {doc_organ_name.lower()} function parameters."],
        "medication_safety_notes": ["Verify current orders and institutional guidelines prior to medication administration."]
    }

    doc_meta = (parsed_ai.get("document") if parsed_ai and isinstance(parsed_ai, dict) else None) or {
        "title": f"{topic_name} Visual Notes",
        "subtitle": "Nursing Study Sheet • High-Yield Summary",
        "topic_type": "disease",
        "student_level": "intermediate",
        "layout_type": "disease_overview",
        "page_size": "mobile_portrait",
        "estimated_pages": 1,
        "educational_disclaimer": "Educational study support material. Follow institutional policy and clinical judgment."
    }

    return {
        "id": f"doc_{uuid.uuid4().hex[:8]}",
        "filename": filename,
        "topic": topic_name,
        "organ_subject": doc_subject,
        "organ_name": doc_organ_name,
        "original_text": raw_text.strip(),
        "visual_requirements": extract_visual_requirements(filename, raw_text, clean_sections),
        "summary": summary_obj,
        "exam_support": exam_support,
        "safety_panel": safety_panel,
        "document": doc_meta,
        "sections": clean_sections,
        "cards": clean_sections,
        "bottom_panels": bottom_panels,
        "header": {
            "title": f"{topic_name} Visual Notes",
            "subtitle": "AI Visual Notes System • High-Yield Study Sheet (Max 8 Cards)",
            "topic": topic_name
        }
    }



# --- OpenAI API with Graceful Deterministic Fallback ---
def call_openai_api(system_prompt: str, user_content: str) -> Optional[dict]:
    """Calls OpenAI API with JSON mode, returning None on quota or network limits."""
    key = os.getenv("OPENAI_API_KEY", "").strip().strip('"').strip("'")
    if not key or key.startswith("sk-proj-placeholder"):
        return None

    for model_name in ["gpt-4o-mini", "gpt-4o"]:
        try:
            import openai
            client = openai.OpenAI(api_key=key)
            resp = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": user_content}],
                temperature=0.1, max_tokens=4000, response_format={"type": "json_object"}
            )
            raw = resp.choices[0].message.content.strip()
            if raw.startswith("```"):
                raw = re.sub(r"^```(?:json)?\n|\n```$", "", raw, flags=re.MULTILINE).strip()
            parsed = json.loads(raw)
            if isinstance(parsed, dict):
                return parsed
        except Exception as e:
            print(f"[{model_name} API Notice] {e}")
    return None

def ensure_6_to_8_sections(original_note: dict, filename: str, raw_text: str) -> dict:
    """Guarantees that the sections array contains between 6 and 8 dynamic sections."""
    sections = original_note.get("sections")
    if not isinstance(sections, list):
        sections = []

    doc_analysis = original_note.get("document_analysis") or {}
    topic_name = doc_analysis.get("topic") or filename.rsplit('.', 1)[0].replace("_", " ").replace("-", " ").title()

    key_terms = extract_key_terms_from_document({"raw_text": raw_text, "sections": sections}, "important terms")
    if key_terms:
        doc_analysis["key_terms_focused"] = key_terms[:10]
    else:
        doc_analysis["key_terms_focused"] = [topic_name, "Clinical Management", "Nursing Interventions"]

    original_note["document_analysis"] = doc_analysis

    existing_titles = { (s.get("title") or "").strip().lower() for s in sections if isinstance(s, dict) }

    priority_fillers = [
        ("Causes & Risk Factors", "dynamic", "Diseases and Conditions", [f"Etiology, primary risk factors, and precipitating triggers for {topic_name}."]),
        ("Signs & Symptoms", "dynamic", "Diseases and Conditions", [f"Clinical manifestations, patient presentation, and physical assessment findings."]),
        ("Diagnostic Workup & Labs", "dynamic", "Laboratory Values", [f"Diagnostic imaging, laboratory indicators, and reference parameters for {topic_name}."]),
        ("Medical Management", "dynamic", "Diseases and Conditions", [f"Pharmacological treatment, therapeutic interventions, and clinical management goals."]),
        ("Nursing Interventions", "dynamic", "Diseases and Conditions", [f"High-priority nursing care, monitoring protocols, and clinical safety measures."]),
        ("Patient Education", "dynamic", "Diseases and Conditions", [f"Patient teaching points, lifestyle modifications, and self-management instructions."]),
        ("Red Flags & Critical Alerts", "dynamic", "Diseases and Conditions", [f"Urgent clinical indicators, red flag symptoms, and emergency complications."]),
        ("Exam Tips & Memory Aid", "dynamic", "Diseases and Conditions", [f"High-yield exam review tips and memory mnemonics for {topic_name}."])
    ]

    idx = len(sections) + 1
    if len(sections) < 6:
        for title, sec_type, source, content in priority_fillers:
            if len(sections) >= 6:
                break
            t_low = title.lower()
            if not any(t in t_low or t_low in t for t in existing_titles):
                sections.append({
                    "section_number": idx,
                    "title": title,
                    "type": sec_type,
                    "priority_source": source,
                    "content": content,
                    "visual": {
                        "required": False,
                        "type": "diagram",
                        "subject": title,
                        "purpose": f"Educational visual for {title}"
                    }
                })
                existing_titles.add(t_low)
                idx += 1

    sections = sections[:8]

    for i, sec in enumerate(sections, start=1):
        if isinstance(sec, dict):
            sec["section_number"] = i

    original_note["sections"] = sections
    return original_note


def clean_bullet_sentence(sentence: str, max_words: int = 40) -> Optional[str]:
    """Cleans a single sentence string, removing metadata, journal stamps, and capping length at max_words."""
    if not sentence or not isinstance(sentence, str):
        return None
    s = sentence.strip()
    s = re.sub(r'AJPHI\s*[|I]\s*VOLUME\s*\d+\s*[|I]\s*\d{4}\s*(?:ORIGINAL ARTICLE)?', '', s, flags=re.I)
    s = re.sub(r'The\s+American\s+Journal\s+of\s+Patient\s+Health\s+Info\s*:\s*\d{4}', '', s, flags=re.I)
    s = re.sub(r'CORONARY\s+ARTERY\s+DISEASE\s+CAD\s+DEMYSTIFIED\s+CAUSES\s+[A-Z0-9\s_\-]*OVERVIEW\s*:\s*', '', s, flags=re.I)
    s = re.sub(r'^[A-Z0-9\s_\-]{10,80}\s*OVERVIEW\s*:\s*', '', s, flags=re.I).strip()
    
    if len(s) < 10 or any(k in s.lower() for k in ["journal", "volume", "ajphi", "author", "available from", "doi:", "issn:"]):
        return None
        
    words = s.split()
    if len(words) > max_words:
        s = " ".join(words[:max_words]).rstrip(',;:-')
        if not s.endswith('.'):
            s += '.'
    elif not s.endswith(('.', '?', '!', ':')):
        s += '.'
        
    return s


def build_section_bullets(raw_texts: list, section_title: str = "", filename: str = "", min_bullets: int = 3, max_bullets: int = 5) -> list:
    """
    Dynamically generates AT LEAST 3 bullet points and AT MOST 5 bullet points for EVERY section based on content richness.
    Each bullet point contains 15 to 35 words of high-yield clinical content.
    """
    candidate_sentences = []
    for rt in raw_texts:
        sentences = re.split(r'(?<=[.!?])\s+|\n+|[;•\-\*●▪■◆➢►○✔✓]\s*', str(rt))
        for st in sentences:
            c_st = clean_bullet_sentence(st, max_words=35)
            if c_st and c_st not in candidate_sentences:
                candidate_sentences.append(c_st)

    bullets = []
    for st in candidate_sentences:
        if st not in bullets:
            bullets.append(st)
            if len(bullets) >= max_bullets:
                break

    # If long sentences can be cleanly split into 4 or 5 distinct high-yield points when content permits
    if len(bullets) < max_bullets and candidate_sentences:
        expanded_bullets = []
        for b in bullets:
            words = b.split()
            if len(words) >= 24 and len(expanded_bullets) < max_bullets:
                sub_parts = re.split(r'\s+;\s+|\s+,\s+and\s+|\s+,\s+which\s+', b)
                if len(sub_parts) > 1 and (len(bullets) - 1 + len(sub_parts)) <= max_bullets:
                    for sp in sub_parts:
                        clean_sp = clean_bullet_sentence(sp, max_words=35)
                        if clean_sp and clean_sp not in expanded_bullets:
                            expanded_bullets.append(clean_sp)
                            if len(expanded_bullets) >= max_bullets:
                                break
                    continue
            if b not in expanded_bullets:
                expanded_bullets.append(b)
        if len(expanded_bullets) >= min_bullets:
            bullets = expanded_bullets[:max_bullets]

    clean_sec_title = re.sub(r'^\d+[\.\)]\s*', '', section_title).strip()

    fillers = [
        f"Primary physiological mechanisms, structural characteristics, and assessment parameters associated with {clean_sec_title}.",
        f"Key clinical diagnostic findings, laboratory indicators, and monitoring protocols relevant to {clean_sec_title}.",
        f"High-priority nursing interventions, therapeutic management goals, and patient safety precautions for {clean_sec_title}.",
        f"Essential patient education points, risk factor modifications, and clinical monitoring recommendations for {clean_sec_title}."
    ]

    for f in fillers:
        if len(bullets) >= min_bullets:
            break
        if f not in bullets:
            bullets.append(f)

    return bullets[:max_bullets]


def sanitize_original_note_sections(original_note: dict, filename: str, raw_text: str) -> dict:
    """
    Sanitizes all sections in Original Note (from AI or deterministic engine):
    1. Guarantees 6 to 8 structured sections.
    2. EVERY section (Sections 1-8) contains AT LEAST 3 bullet points and MAXIMUM 5 bullet points.
    3. Strips all journal headers, author stamps, volume tags, and repetitive title prefixes.
    4. Capping length of each bullet point to 20-35 words.
    """
    if not original_note or not isinstance(original_note, dict):
        original_note = {}

    sections = original_note.get("sections")
    if not isinstance(sections, list):
        sections = []

    clean_sections = []
    for s_idx, sec in enumerate(sections, start=1):
        if not isinstance(sec, dict):
            continue
            
        raw_title = sec.get("title") or sec.get("heading") or f"Section {s_idx}"
        clean_title = re.sub(r'^\d+[\.\)]\s*', '', raw_title).strip()
        clean_title = clean_raw_text_metadata(clean_title)
        if not clean_title:
            clean_title = f"Clinical Section {s_idx}"

        raw_texts = []
        if sec.get("content"):
            if isinstance(sec["content"], list):
                raw_texts.extend([str(c) for c in sec["content"] if c])
            elif isinstance(sec["content"], str):
                raw_texts.append(sec["content"])

        if sec.get("items"):
            for it in sec["items"]:
                if isinstance(it, dict) and it.get("text"):
                    raw_texts.append(it["text"])
                elif isinstance(it, str):
                    raw_texts.append(it)

        if sec.get("bullets"):
            for b in sec["bullets"]:
                if isinstance(b, str):
                    raw_texts.append(b)

        clean_bullets = build_section_bullets(raw_texts, section_title=clean_title, filename=filename, min_bullets=3, max_bullets=5)

        sec_items = []
        for cb in clean_bullets:
            parts = re.split(r':|\s+[–—\-]\s+', cb, maxsplit=1)
            if len(parts) == 2 and len(parts[0].strip()) < 25 and len(parts[1].strip()) > 5:
                lbl = parts[0].strip().upper()
                txt = parts[1].strip()
                if not txt.endswith('.'): txt += '.'
                sec_items.append({"label": lbl, "text": txt, "tag": lbl})
            else:
                sec_items.append({"label": clean_title.upper(), "text": cb, "tag": clean_title.upper()})

        is_anatomy_sec = "anatomy" in clean_title.lower() or bool(sec.get("isAnatomyCard"))
        is_pathway_sec = not is_anatomy_sec and ("patho" in clean_title.lower() or "pathway" in clean_title.lower() or sec.get("type") == "flowchart")
        
        if is_anatomy_sec:
            req = {
                "type": "anatomy",
                "is_anatomy": True,
                "subject": clean_title,
                "section_title": clean_title,
                "purpose": f"Anatomical structure visual for {clean_title}",
                "required": True
            }
            existing_img = sec.get("image") or sec.get("image_url") or (sec.get("visual") or {}).get("image_url")
            if existing_img:
                sec_img = existing_img
            else:
                sec_img, _, _ = resolve_or_generate_visual(req, raw_text=raw_text, filename=filename)
            sec_visual = {
                "required": True,
                "type": "anatomy",
                "subject": clean_title,
                "purpose": f"Anatomical structure visual for {clean_title}",
                "image_url": sec_img or ""
            }
        else:
            sec_img = ""
            sec_visual = {
                "required": False,
                "type": "none",
                "subject": clean_title,
                "purpose": "No image required",
                "image_url": ""
            }

        clean_sections.append({
            "section_number": s_idx,
            "title": clean_title.title(),
            "heading": clean_title.upper(),
            "type": sec.get("type") or sec.get("section_type") or ("flowchart" if is_pathway_sec else "dynamic"),
            "priority_source": sec.get("priority_source") or "Document Emphasis",
            "content": clean_bullets,
            "items": sec_items,
            "bullets": clean_bullets,
            "flows": sec.get("flows") or sec.get("flow") or [],
            "image": sec_img or "",
            "image_url": sec_img or "",
            "visual": sec_visual
        })

    original_note["sections"] = clean_sections
    return ensure_6_to_8_sections(original_note, filename, raw_text)


def build_deterministic_original_note(file_name: str, raw_text: str) -> Dict[str, Any]:
    """Fallback parser that structures raw document text into dynamic section cards (6-8 sections)."""
    cleaned_text = clean_raw_text_metadata(raw_text)
    topic_slug, topic_title = detect_primary_subject(file_name, cleaned_text)
    parsed_sections = parse_dynamic_sections_from_text(cleaned_text, topic_name=topic_title)
    
    sections_list = []
    for idx, sec in enumerate(parsed_sections, start=1):
        heading = sec.get("heading") or f"Section {idx}"
        lines = sec.get("lines") or []
        bullets = build_section_bullets(lines, section_title=heading, filename=file_name, min_bullets=3, max_bullets=5)
        sections_list.append({
            "section_number": idx,
            "title": heading,
            "heading": heading,
            "type": sec.get("type", "dynamic"),
            "priority_source": "Document Emphasis",
            "content": bullets,
            "bullets": bullets
        })
        
    clean_doc_title = os.path.splitext(file_name)[0].replace("_", " ").replace("-", " ").title() if file_name else "Uploaded Notes"
    original_note = {
        "document_analysis": {
            "title": file_name,
            "topic": clean_doc_title,
            "primary_subject": topic_title,
            "document_type": "general"
        },
        "filename": file_name,
        "original_text": raw_text,
        "sections": sections_list
    }
    return ensure_6_to_8_sections(original_note, file_name, raw_text)


def analyze_document_for_original_notes(file_name: str, raw_text: str) -> Dict[str, Any]:
    """
    Executes Stage 1 Document Analysis Engine.
    Analyzes content, extracts priorities, and organizes document into 6-8 structured Original Note sections.
    """
    truncated = raw_text[:15000] if len(raw_text) > 15000 else raw_text
    user_msg = f"Document Filename: {file_name}\n\nRAW UPLOADED DOCUMENT CONTENT:\n{truncated}"

    parsed_json = call_openai_api(DOCUMENT_ANALYSIS_ORIGINAL_NOTE_PROMPT, user_msg)
    if parsed_json and isinstance(parsed_json, dict) and "sections" in parsed_json:
        result = sanitize_original_note_sections(parsed_json, file_name, raw_text)
        result["filename"] = file_name
        result["original_text"] = raw_text
        return result

    result = build_deterministic_original_note(file_name, raw_text)
    result = sanitize_original_note_sections(result, file_name, raw_text)
    result["filename"] = file_name
    result["original_text"] = raw_text
    return result


def build_visual_notes_from_original_note(original_note: dict) -> dict:
    """
    Executes Stage 2 Visual Notes Generation.
    Converts structured Original Note (6-8 sections) into 3D Visual Cards Dashboard Schema with diagram resolution.
    """
    file_name = original_note.get("filename") or "Uploaded_Notes.pdf"
    raw_text = original_note.get("original_text") or ""
    
    sections_in = original_note.get("sections") or []
    
    card_sections = []
    for s in sections_in:
        if not isinstance(s, dict):
            continue
        stitle = s.get("title") or "Section"
        content_items = s.get("content") or []
        if isinstance(content_items, str):
            content_items = [content_items]
            
        items_objs = []
        bullets_list = []
        for c in content_items:
            if isinstance(c, str):
                parts = c.split(":", 1)
                if len(parts) == 2 and len(parts[0].strip()) < 35:
                    items_objs.append({"label": parts[0].strip().upper(), "text": parts[1].strip()})
                else:
                    items_objs.append({"label": stitle.upper(), "text": c.strip()})
                bullets_list.append(c.strip())
            elif isinstance(c, dict):
                items_objs.append(c)

        flow_steps = s.get("flow") or s.get("flows") or []
        sec_type = "flowchart" if (flow_steps or "pathophysiology" in stitle.lower()) else "tagged_items"

        card_sections.append({
            "id": f"section_{s.get('section_number', len(card_sections)+1)}",
            "title": stitle.upper(),
            "heading": stitle.upper(),
            "type": sec_type,
            "section_type": sec_type,
            "items": items_objs,
            "bullets": bullets_list,
            "flows": flow_steps,
            "content": "\n".join(bullets_list)
        })

    clean_doc_name = os.path.splitext(file_name)[0].replace("_", " ").replace("-", " ").title() if file_name else "Uploaded Notes"
    parsed_ai = {
        "document": {
            "title": f"{clean_doc_name} Visual Notes",
            "subtitle": "AI Visual Notes • 3D Cards Dashboard",
            "topic_type": "disease"
        },
        "sections": card_sections
    }

    plan = build_modular_visual_dashboard_schema(file_name, raw_text, parsed_ai)
    plan["filename"] = file_name
    plan["topic"] = clean_doc_name
    plan["original_note"] = original_note
    return plan


def build_visual_notes_from_text(file_name: str, raw_text: str) -> Dict[str, Any]:
    """Builds visual study notes schema using OpenAI API or deterministic engine."""
    orig = analyze_document_for_original_notes(file_name, raw_text)
    return build_visual_notes_from_original_note(orig)


# --- AI Edit Pipeline ---
AI_EDIT_SYSTEM_PROMPT = """# VISUAL NOTES AI MODIFICATION RULES

You are modifying an already-generated Visual Notes page.

Your MOST IMPORTANT responsibility is to modify ONLY the content or visual styling requested by the user while preserving the existing Visual Notes structure exactly.

==================================================
1. IMMUTABLE STRUCTURE RULE — ABSOLUTE PRIORITY
==================================================

The existing Visual Notes structure is LOCKED.

NEVER change, remove, reorder, resize, merge, split, or recreate:

- Cards
- Card positions
- Card hierarchy
- Card dimensions
- Grid/column structure
- Section structure
- Section order
- Visual hierarchy
- Header/title position
- Organ illustration position
- Diagram position
- Note blocks
- Spacing system
- Padding
- Margins
- Border radius
- Card borders
- Existing icons
- Existing images
- Existing diagrams
- Existing responsive behavior
- Overall page layout

The existing design must remain visually recognizable as the SAME Visual Notes page after every AI modification.

Think of the existing UI as a fixed template.

The AI is allowed to modify CONTENT INSIDE the template, but it is NOT allowed to redesign the template.

==================================================
2. NEVER MODIFY THE CODE STRUCTURE
==================================================

If this request is being applied inside an existing application:

DO NOT modify:

- React component structure
- HTML structure
- DOM hierarchy
- CSS layout rules
- Tailwind layout classes
- Grid/flex configuration
- Card components
- reusable UI components
- routing
- state architecture
- rendering logic
- existing component names
- existing IDs/classes
- database structure

Do not generate replacement UI code.

Modify only the data/content/state that controls the existing Visual Notes.

If a requested change cannot be achieved without changing the layout, DO NOT change the layout. Apply the closest possible content-only modification.

==================================================
3. "MAKE IT MORE CONCISE"
==================================================

When the user says:

"Make it more concise"

OR:

"Make it shorter"

OR similar instructions:

DO NOT redesign the cards.

DO NOT remove cards.

DO NOT remove important sections.

Instead:

- Every single card MUST contain EXACTLY 3 points of information (the 3 most critical, high-yield points).
- If a card has more than 3 items or bullet points, identify and keep ONLY the top 3 most important points (e.g. definition, primary mechanism/cause, hallmark sign, or key treatment), removing lower priority details.
- Reduce long paragraphs into short, concise single sentences for each of the 3 points.
- Shorten unnecessarily long sentences.
- Keep important information and remove repetition.
- Keep the same headings.
- Keep the same card positions.
- Keep the same visual structure.

==================================================
4. "HIGHLIGHT KEY TERMS IN RED"
==================================================

When the user says:

"Highlight key terms in red"

OR similar instructions:

DO NOT change the card design.

DO NOT create new cards.

DO NOT change the layout.

Identify important terms inside the existing content and render ONLY those terms using the existing red highlight/text style.

Only the important terms should become red.
Do NOT make entire paragraphs red.
Do NOT randomly color text.
Do NOT change unrelated colors.

==================================================
5. "SIMPLIFY FOR FIRST YEAR STUDENTS"
==================================================

When the user says:

"Simplify for first year students"

OR similar instructions (e.g. "simplify", "1st year", "easier", "simple", "beginner"):

Keep:

- Same cards
- Same sections
- Same card positions
- Same order
- Same layout
- Same illustrations
- Same diagrams
- CRITICAL: KEEP EXACTLY THE SAME NUMBER OF POINTS / ITEMS / BULLETS IN EACH CARD as were originally present. DO NOT remove, omit, or reduce any points or items.

Only simplify the language:

- Rewrite complex medical, anatomical, or scientific jargon into clear, easy-to-read, plain English suitable for 1st-year beginner students.
- Whenever a complex medical term is used (e.g. Dyspnea, Pathophysiology, Atelectasis, Infarction, Etiology), keep the term but ALWAYS follow it with a clear parenthetical explanation (e.g. "Dyspnea (shortness of breath)", "Etiology (Causes)", "Hypoxemia (low blood oxygen)").
- Break down dense multi-clause sentences into short, simple sentences.
- Simplify section titles and item labels to be intuitive (e.g. "Etiology & Pathogenesis" -> "Causes & How It Happens (Etiology)").

==================================================
6. "CHANGE THE COLOR THEME"
==================================================

When the user says:

"Change the color theme"

OR similar instructions:

Change ONLY the visual color palette.
The structure must remain EXACTLY the same.

Keep:

- Same cards
- Same card sizes
- Same positions
- Same spacing
- Same typography
- Same content
- Same illustrations
- Same diagrams
- Same layout

Use a subtle, light, professional medical color theme.

Do NOT use:

- very dark backgrounds
- excessive gradients
- neon colors
- extremely saturated colors
- colors that reduce readability

==================================================
7. USER CUSTOM AI PROMPTS
==================================================

The user may type any custom request into the AI Prompt field.

Always follow the user's requested content change.

BUT:

The existing layout remains LOCKED.

If the user asks to elaborate:

- Expand the relevant content only.
- Keep the same card.
- Keep the same section.
- Keep the same position.
- Keep the same design.
- Keep the same typography.

==================================================
8. "ADD MORE INFORMATION"
==================================================

If the user asks to elaborate or add information:
First determine whether the requested information is already supported by the original document/notes.
If it is source-supported: Add or expand the relevant content.
If it is not supported by the original source: Do not fabricate it.
Do not alter the existing card structure.

==================================================
9. "ADD MORE DIAGRAMS"
==================================================

If the user requests additional diagrams:
Use the existing visual system.
Do NOT randomly insert diagrams into the page.
Prefer placing the requested visual inside an existing appropriate visual/card area if the existing design supports it.
If there is no suitable existing area: Do not break the layout. Preserve the original design.

==================================================
10. ORIGINAL CONTENT MUST REMAIN SAFE
==================================================

Never accidentally delete content because of a small modification.
Identify exactly what the user wants changed.
Locate the relevant content.
Modify only that content.
Preserve everything else.

==================================================
11. NO UNREQUESTED CHANGES
==================================================

If the user asks for ONE change, make ONE change.
Do NOT automatically rewrite other sections, change colors, font, rearrange cards, or change section order.

==================================================
12. PRESERVE VISUAL CONSISTENCY
==================================================

After every modification, the Visual Notes should still look like the original design.

==================================================
13. PRIORITY ORDER
==================================================

1. Preserve existing layout and card structure.
2. Follow the user's requested modification.
3. Preserve original/source-supported information.
4. Maintain readability.
5. Maintain visual consistency.
6. Make the smallest necessary change.

If there is a conflict between the requested modification and the existing layout:
PRESERVE THE LAYOUT. Adapt the content instead.

==================================================
14. FINAL VALIDATION BEFORE RENDERING
==================================================

Verify same number of cards, same card positions, same section order, same layout, same visual hierarchy.
Modify ONLY requested content/style.

==================================================
15. CORE PRINCIPLE
==================================================

CONTENT CAN CHANGE.
STYLING CAN CHANGE ONLY WHEN REQUESTED.
LAYOUT DOES NOT CHANGE.
CARD STRUCTURE DOES NOT CHANGE.

Return valid JSON object matching the input schema structure."""

def clean_markdown_stars(val: str) -> str:
    if not val or not isinstance(val, str):
        return val
    return val.replace("**", "").replace("*", "").strip()


def fuzzy_find_target_section(instruction: str, sections: list) -> dict:
    """Intelligently identifies the target section card matching the user's prompt (handles typos, Banglish, synonyms)."""
    text_lower = instruction.lower()
    
    topic_keywords = {
        "overview": ["overview", "definition", "introduction", "overvew", "intro", "summary", "basics"],
        "anatomy": ["anatomy", "atonomy", "structure", "organ", "lung", "alveoli", "membrane", "vessel", "heart", "brain"],
        "pathophysiology": ["pathophysiology", "pathology", "pathway", "mechanism", "process", "patho", "pasthology", "blood flow", "flowchart", "causes", "etiology"],
        "diagnostic": ["diagnostic", "diagostic", "diagnostics", "diagnos", "test", "tests", "abg", "xray", "ct", "lab", "findings"],
        "nursing": ["nursing", "nursin", "intervention", "interventions", "care", "management", "monitoring", "vital"],
        "stages": ["stage", "stages", "stags", "phase", "phases", "exudative", "proliferative", "fibrotic"],
        "medication": ["medication", "medications", "meds", "drug", "drugs", "treatment", "treament", "rx", "pharmacology", "prescription"]
    }

    words = re.findall(r'\b\w+\b', text_lower)
    best_sec = None
    best_score = 0.0

    for sec in sections:
        if not isinstance(sec, dict):
            continue
        title = (sec.get("title") or sec.get("heading") or "").lower()
        clean_title = re.sub(r'^\d+[\.\)]\s*', '', title).strip()

        card_key = "other"
        for key, k_list in topic_keywords.items():
            if any(k in clean_title for k in k_list[:3]):
                card_key = key
                break

        target_k_list = topic_keywords.get(card_key, [clean_title])
        for w in words:
            if len(w) < 3:
                continue
            for kw in target_k_list:
                if w in kw or kw in w:
                    score = 0.95
                else:
                    score = difflib.SequenceMatcher(None, w, kw).ratio()

                if score > best_score and score >= 0.65:
                    best_score = score
                    best_sec = sec

    return best_sec



STUDENT_SIMPLIFICATIONS = [
    (r'\bdyspnea\b(?!\s*\()', 'shortness of breath (dyspnea)'),
    (r'\bhypoxemia\b(?!\s*\()', 'low blood oxygen levels (hypoxemia)'),
    (r'\bhypoxia\b(?!\s*\()', 'low tissue oxygen (hypoxia)'),
    (r'\betiology\b(?!\s*\()', 'causes & origin (etiology)'),
    (r'\bpathogenesis\b(?!\s*\()', 'how the disease develops (pathogenesis)'),
    (r'\bpathophysiology\b(?!\s*\()', 'body function changes (pathophysiology)'),
    (r'\batelectasis\b(?!\s*\()', 'collapsed lung air sacs (atelectasis)'),
    (r'\binfarction\b(?!\s*\()', 'tissue death due to blocked blood supply (infarction)'),
    (r'\bischemia\b(?!\s*\()', 'reduced blood flow (ischemia)'),
    (r'\bnecrosis\b(?!\s*\()', 'cell and tissue death (necrosis)'),
    (r'\bedema\b(?!\s*\()', 'fluid swelling (edema)'),
    (r'\bhemorrhage\b(?!\s*\()', 'severe bleeding (hemorrhage)'),
    (r'\btachycardia\b(?!\s*\()', 'rapid heart rate (tachycardia)'),
    (r'\bbradycardia\b(?!\s*\()', 'slow heart rate (bradycardia)'),
    (r'\bhypertension\b(?!\s*\()', 'high blood pressure (hypertension)'),
    (r'\bhypotension\b(?!\s*\()', 'low blood pressure (hypotension)'),
    (r'\btachypnea\b(?!\s*\()', 'rapid breathing rate (tachypnea)'),
    (r'\bcyanosis\b(?!\s*\()', 'bluish skin tint from low oxygen (cyanosis)'),
    (r'\bsepsis\b(?!\s*\()', 'severe body-wide infection response (sepsis)'),
    (r'\bidiopathic\b(?!\s*\()', 'unknown cause (idiopathic)'),
    (r'\bprophylaxis\b(?!\s*\()', 'prevention step (prophylaxis)'),
    (r'\bprophylactic\b(?!\s*\()', 'preventative (prophylactic)'),
    (r'\bdiagnosis\b(?!\s*\()', 'identifying the disease (diagnosis)'),
    (r'\basymptomatic\b(?!\s*\()', 'without noticeable symptoms (asymptomatic)'),
    (r'\bthrombosis\b(?!\s*\()', 'blood clot formation (thrombosis)'),
    (r'\bembolism\b(?!\s*\()', 'traveling blood clot blocking an artery (embolism)'),
    (r'\bauscultation\b(?!\s*\()', 'listening with a stethoscope (auscultation)'),
    (r'\bpalpation\b(?!\s*\()', 'feeling by hand (palpation)'),
    (r'\bfibrosis\b(?!\s*\()', 'tissue scarring and stiffness (fibrosis)'),
    (r'\bstenosis\b(?!\s*\()', 'abnormal narrowing (stenosis)'),
    (r'\beffusion\b(?!\s*\()', 'fluid collection (effusion)'),
    (r'\bsyncope\b(?!\s*\()', 'fainting (syncope)'),
    (r'\bpruritus\b(?!\s*\()', 'itching sensation (pruritus)'),
    (r'\berythema\b(?!\s*\()', 'skin redness (erythema)'),
    (r'\bleukocytosis\b(?!\s*\()', 'high white blood cell count (leukocytosis)'),
    (r'\bthrombocytopenia\b(?!\s*\()', 'low platelet count (thrombocytopenia)'),
    (r'\banemia\b(?!\s*\()', 'low red blood cell count (anemia)'),
    (r'\barrhythmia\b(?!\s*\()', 'irregular heart rhythm (arrhythmia)'),
    (r'\bjaundice\b(?!\s*\()', 'yellowish skin and eyes from liver issues (jaundice)'),
    (r'\bcongenital\b(?!\s*\()', 'present from birth (congenital)'),
    (r'\bcharacterized by\b', 'marked by'),
    (r'\bmanifests as\b', 'shows up as'),
    (r'\bsecondary to\b', 'caused by'),
    (r'\bconcomitant with\b', 'occurring alongside'),
    (r'\bfirst-line therapy\b', 'first choice treatment'),
    (r'\bpharmacotherapy\b', 'medication treatment'),
    (r'\bcontraindicated\b', 'unsafe / should not be used (contraindicated)'),
    (r'\badverse effects\b', 'side effects'),
]

def simplify_text_for_students(text: str) -> str:
    """Simplifies complex medical text into plain English with clear parenthetical explanations for 1st year students."""
    if not text or not isinstance(text, str):
        return ""
    clean = clean_markdown_stars(re.sub(r'<[^>]+>', '', text)).strip()
    for pattern, replacement in STUDENT_SIMPLIFICATIONS:
        clean = re.sub(pattern, replacement, clean, flags=re.IGNORECASE)
    return clean

def extract_elaboration_from_document_text(plan: dict, target_sec: dict, instruction: str = "") -> Tuple[str, str]:
    """Extracts real 1:1 clinical sentences from uploaded document text strictly based on source content."""
    raw_text = (plan.get("original_text") or plan.get("raw_text") or plan.get("document_text") or "").strip()
    sec_title = (target_sec.get("title") or target_sec.get("heading") or "Clinical Section").strip()
    clean_sec_title = re.sub(r'^\d+[\.\)]\s*', '', sec_title).strip()

    inst_clean = instruction.lower()
    for phrase in ["what is", "tell me about", "i want to know more in details", "i want to know more", "details", "explain", "describe", "more about", "in detail"]:
        inst_clean = inst_clean.replace(phrase, "")
    query_topic = inst_clean.strip(" .?!,")

    matched_kws = [w for w in re.findall(r'\b[a-z]{3,}\b', query_topic) if w not in ["what", "want", "know", "more", "detail", "details", "this", "that"]]
    if not matched_kws:
        matched_kws = [w for w in clean_sec_title.lower().split() if len(w) > 3]

    extracted_sentences = []
    if raw_text:
        sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+|\n+', raw_text) if len(s.strip()) > 10]
        for s in sentences:
            s_lower = s.lower()
            if any(kw in s_lower for kw in matched_kws):
                existing_texts = [it.get("text", "").lower() for it in target_sec.get("items", []) if isinstance(it, dict)]
                if not any(s_lower in ext for ext in existing_texts):
                    extracted_sentences.append(s)
                    if len(extracted_sentences) >= 2:
                        break

    label_topic = query_topic.upper() if query_topic else clean_sec_title.upper()
    elab_label = f"{label_topic} (DETAILS)"

    if extracted_sentences:
        elab_text = " ".join(extracted_sentences)
    else:
        elab_text = f"Key details regarding {query_topic or clean_sec_title} from document content."

    return elab_label, elab_text


def extract_key_terms_from_document(plan: dict, instruction: str = "") -> List[str]:
    """Dynamically extracts all important medical & clinical terms from ANY uploaded document without hardcoded lists."""
    raw_text = (plan.get("raw_text") or plan.get("document_text") or "").strip()
    custom_words = [w.strip() for w in re.findall(r'["\']([^"\']+)["\']', instruction)]

    sections = plan.get("sections") or plan.get("cards") or []
    all_texts = []
    for sec in sections:
        if isinstance(sec, dict):
            for it in sec.get("items", []):
                if isinstance(it, dict) and it.get("text"):
                    all_texts.append(it["text"])
            for b in sec.get("bullets", []):
                if isinstance(b, str):
                    all_texts.append(b)

    combined_text = raw_text + " " + " ".join(all_texts)

    # 1. Medical Acronyms & Abbreviations (2-8 uppercase letters like ABG, ECG, CT, MRI, BP, HR, HbA1c)
    acronyms = set(re.findall(r'\b[A-Z0-9]{2,8}\b', combined_text))

    # 2. Medical Suffix Words (-itis, -emia, -osis, -pathy, -trophy, -pnea, -uria, -gram, -stasis, -sclerosis, -tension, etc.)
    suffix_pattern = re.compile(r'\b\w*(?:itis|emia|osis|pathy|trophy|pnea|uria|gram|stasis|sclerosis|tension|plegia|lysis|oma|stomy|tomy|plasty|spasm|emia|itis)\b', re.IGNORECASE)
    suffix_words = set(suffix_pattern.findall(combined_text))

    # 3. Capitalized Clinical Nouns & Specialty Words (4-20 chars)
    cap_words = set(re.findall(r'\b[A-Z][a-z]{3,20}\b', combined_text))

    stop_words = {"This", "That", "There", "These", "Those", "What", "When", "Where", "Which", "While", "With", "From", "Into", "Over", "Under", "About", "After", "Before", "During", "Each", "Every", "Most", "Some", "Such", "Than", "Then", "They", "Them", "Their", "Stage", "Phase", "Step", "Card", "Section", "Note", "Notes", "Visual"}
    cap_words = {w for w in cap_words if w not in stop_words}

    all_terms = list(set(custom_words + list(acronyms) + list(suffix_words) + list(cap_words)))
    return [t for t in all_terms if len(t) >= 3]


def extract_color_from_instruction(instruction: str) -> str:
    """Extracts target color accent/theme name from user prompt."""
    txt = instruction.lower().strip()
    if "pastel green" in txt or "mint" in txt or "sage" in txt:
        return "emerald"
    elif "pastel blue" in txt or "sky" in txt:
        return "sky"
    elif "pastel purple" in txt or "lavender" in txt or "lilac" in txt:
        return "purple"
    elif "pastel pink" in txt or "rose" in txt:
        return "pink"
    elif "pastel" in txt or "soft color" in txt or "palette" in txt:
        return "emerald"
    
    colors = ["rose", "pink", "purple", "violet", "indigo", "blue", "sky", "teal", "emerald", "green", "amber", "orange", "yellow", "coral"]
    for c in colors:
        if c in txt:
            return "emerald" if c == "green" else c
    return "emerald"


def process_deterministic_ai_edit(plan: dict, instruction: str) -> dict:
    """Smart deterministic rule processor for AI modification instructions obeying all 15 locked rules."""
    updated = json.loads(json.dumps(plan))
    text_lower = instruction.strip().lower()
    sections = updated.get("sections") or updated.get("cards") or []

    # Clean any markdown asterisks and HTML tags in labels or headings
    for sec in sections:
        if isinstance(sec, dict):
            if sec.get("title"):
                sec["title"] = clean_markdown_stars(re.sub(r'<[^>]+>', '', str(sec["title"])))
            if sec.get("heading"):
                sec["heading"] = clean_markdown_stars(re.sub(r'<[^>]+>', '', str(sec["heading"])))
            for it in sec.get("items", []):
                if isinstance(it, dict) and it.get("label"):
                    it["label"] = clean_markdown_stars(re.sub(r'<[^>]+>', '', str(it["label"])))
            for fl in sec.get("flows", []):
                if isinstance(fl, dict) and fl.get("label"):
                    fl["label"] = clean_markdown_stars(re.sub(r'<[^>]+>', '', str(fl["label"])))

    # Rule 3: Concise / Shorter (Exactly 3 most important information points per card)
    if any(k in text_lower for k in ["concise", "shorter", "brief", "summarize", "short"]):
        for sec in sections:
            if isinstance(sec, dict):
                items = sec.get("items") or []
                if items and isinstance(items, list):
                    if len(items) > 3:
                        def item_importance_score(it_obj, idx):
                            score = 0.0
                            lbl = (it_obj.get("label") or "").upper()
                            txt = (it_obj.get("text") or "").lower()
                            # Priority clinical labels
                            if any(k in lbl for k in ["DEFINITION", "OVERVIEW", "KEY CONCEPT", "WHAT IT IS"]):
                                score += 20
                            elif any(k in lbl for k in ["HALLMARK", "PRIMARY", "FIRST-LINE", "GOLD STANDARD", "CRITICAL"]):
                                score += 18
                            elif any(k in lbl for k in ["ETIOLOGY", "CAUSE", "PATHOPHYSIOLOGY", "MECHANISM"]):
                                score += 15
                            elif any(k in lbl for k in ["DIAGNOSIS", "TRIAD", "SYMPTOMS", "PRESENTATION", "CLINICAL"]):
                                score += 14
                            elif any(k in lbl for k in ["TREATMENT", "MANAGEMENT", "INTERVENTION", "MEDICATION"]):
                                score += 12
                            elif any(k in lbl for k in ["COMPLICATION", "PROGNOSIS", "RISK"]):
                                score += 10

                            # Content keyword signals
                            if any(k in txt for k in ["hallmark", "most common", "triad", "first-line", "gold standard", "diagnostic", "caused by", "primary"]):
                                score += 5

                            # Positional bonus for foundational points
                            score += max(0, 5 - idx)
                            return score

                        indexed_items = list(enumerate(items))
                        indexed_items.sort(key=lambda x: item_importance_score(x[1], x[0]), reverse=True)
                        top_3_indexed = indexed_items[:3]
                        # Restore original relative ordering for logical narrative
                        top_3_indexed.sort(key=lambda x: x[0])
                        items = [x[1] for x in top_3_indexed]

                    # Shorten each of the 3 items to 1 concise sentence
                    for it in items:
                        if isinstance(it, dict) and it.get("text"):
                            sentences = [s.strip() for s in re.split(r'\.\s+', str(it["text"])) if s.strip()]
                            if sentences:
                                first_s = sentences[0]
                                if not first_s.endswith('.'):
                                    first_s += '.'
                                it["text"] = first_s

                    sec["items"] = items
                    sec["bullets"] = [it.get("text") for it in items if isinstance(it, dict) and it.get("text")]

                elif sec.get("bullets") and isinstance(sec["bullets"], list):
                    bullets = [b for b in sec["bullets"] if isinstance(b, str) and b.strip()]
                    if len(bullets) > 3:
                        bullets = bullets[:3]
                    short_bullets = []
                    for b in bullets:
                        sentences = [s.strip() for s in re.split(r'\.\s+', b) if s.strip()]
                        if sentences:
                            first_s = sentences[0]
                            if not first_s.endswith('.'):
                                first_s += '.'
                            short_bullets.append(first_s)
                    sec["bullets"] = short_bullets

                # Truncate flows to max 3 steps
                if sec.get("flows") and isinstance(sec["flows"], list) and len(sec["flows"]) > 3:
                    sec["flows"] = [sec["flows"][0], sec["flows"][len(sec["flows"]) // 2], sec["flows"][-1]]

                # Truncate content to max 1 sentence
                if sec.get("content") and isinstance(sec["content"], str):
                    sentences = [s.strip() for s in re.split(r'\.\s+', sec["content"]) if s.strip()]
                    if sentences:
                        sec["content"] = sentences[0] + ("." if not sentences[0].endswith('.') else "")

    # Rule 4: Highlight key terms in red (Universal Dynamic Term Extractor)
    elif any(k in text_lower for k in ["red", "highlight", "important terms", "color red"]):
        all_terms = extract_key_terms_from_document(plan, instruction)

        for sec in sections:
            if isinstance(sec, dict):
                if sec.get("items") and isinstance(sec["items"], list):
                    for it in sec["items"]:
                        if isinstance(it, dict):
                            # Clean label from any HTML tags
                            if it.get("label"):
                                it["label"] = re.sub(r'<[^>]+>', '', str(it["label"]))
                            val = it.get("text", "")
                            if isinstance(val, str) and val and "<span" not in val:
                                for term in all_terms:
                                    pattern = re.compile(r'\b' + re.escape(term) + r'\b', re.IGNORECASE)
                                    if pattern.search(val):
                                        val = pattern.sub(f'<span style="color:#dc2626; font-weight:bold;">{term}</span>', val)
                                        it["text"] = val
                if sec.get("bullets") and isinstance(sec["bullets"], list):
                    new_bullets = []
                    for b in sec["bullets"]:
                        if isinstance(b, str) and b and "<span" not in b:
                            b_val = b
                            for term in all_terms:
                                pattern = re.compile(r'\b' + re.escape(term) + r'\b', re.IGNORECASE)
                                if pattern.search(b_val):
                                    b_val = pattern.sub(f'<span style="color:#dc2626; font-weight:bold;">{term}</span>', b_val)
                            new_bullets.append(b_val)
                        else:
                            new_bullets.append(b)
                    sec["bullets"] = new_bullets

    # Rule 5: Simplify for first year students (Beginner Wording & Easy Explanations - Preserves ALL Points)
    elif any(k in text_lower for k in ["simplify", "first year", "1st year", "student", "easier", "simple", "beginner", "easy"]):
        for sec in sections:
            if isinstance(sec, dict):
                # Apply pastel colors suitable for students
                sec["accent"] = "sky"
                # PRESERVES ALL ITEMS without dropping any point
                if sec.get("items") and isinstance(sec["items"], list):
                    for it in sec["items"]:
                        if isinstance(it, dict):
                            if it.get("text"):
                                it["text"] = simplify_text_for_students(it["text"])
                            if it.get("label"):
                                it["label"] = simplify_text_for_students(it["label"])
                # PRESERVES ALL BULLETS without dropping any point
                if sec.get("bullets") and isinstance(sec["bullets"], list):
                    sec["bullets"] = [simplify_text_for_students(b) for b in sec["bullets"] if isinstance(b, str) and b.strip()]
                # PRESERVES ALL FLOWS without dropping any point
                if sec.get("flows") and isinstance(sec["flows"], list):
                    for fl in sec["flows"]:
                        if isinstance(fl, dict) and fl.get("label"):
                            fl["label"] = simplify_text_for_students(fl["label"])

    # Rule 6: Change Color Theme (Universal Dynamic Color Phrase Extractor)
    elif any(k in text_lower for k in ["theme", "color", "pastel", "palette", "soft color", "pink", "rose", "purple", "indigo", "emerald", "green", "teal", "amber", "yellow", "orange", "blue", "sky", "violet", "lavender", "peach", "mint", "coral"]):
        target_accent = extract_color_from_instruction(instruction)
        for sec in sections:
            if isinstance(sec, dict):
                sec["accent"] = target_accent
        updated["theme"] = target_accent

    # Universal Custom Request Engine: Dynamic Document Text & Medical Knowledge Base Extraction
    else:
        target_sec = fuzzy_find_target_section(instruction, sections)
        if not target_sec and sections:
            target_sec = sections[0]

        if target_sec and isinstance(target_sec, dict):
            elab_label, elab_text = extract_elaboration_from_document_text(plan, target_sec, instruction)

            if not isinstance(target_sec.get("items"), list):
                target_sec["items"] = []
            if not any(it.get("label") == elab_label for it in target_sec["items"] if isinstance(it, dict)):
                target_sec["items"].append({
                    "label": elab_label,
                    "text": elab_text
                })
            if isinstance(target_sec.get("bullets"), list):
                if not any(elab_text in str(b) for b in target_sec["bullets"]):
                    target_sec["bullets"].append(elab_text)

    updated["sections"] = sections
    updated["cards"] = sections
    updated["is_visual_generated"] = True
    return updated


def apply_ai_edit(plan: dict, instruction: str) -> dict:
    """Applies user's modification prompt targetedly to the relevant card, adding exactly 1 bullet point while keeping all layout and images 100% intact."""
    text_lower = instruction.strip().lower()
    
    # Route structured instructions directly to deterministic engine to guarantee exact rule application
    if any(k in text_lower for k in [
        "concise", "shorter", "brief", "summarize", "short",
        "red", "highlight", "important terms", "color red",
        "simplify", "first year", "1st year", "student", "easier", "simple", "beginner", "easy",
        "theme", "color", "pastel", "palette", "soft color", "pink", "rose", "purple", "indigo", "emerald", "green", "teal", "amber", "yellow", "orange", "blue", "sky", "violet", "lavender", "peach", "mint", "coral"
    ]):
        return process_deterministic_ai_edit(plan, instruction)

    sections = plan.get("sections") or plan.get("cards") or []
    if not sections:
        return process_deterministic_ai_edit(plan, instruction)

    # Provide available cards context so LLM can identify target card and write 1 specific bullet point
    card_summaries = []
    for s_idx, sec in enumerate(sections):
        if not isinstance(sec, dict):
            continue
        c_id = sec.get("id") or f"card_{s_idx}"
        c_title = sec.get("title") or sec.get("heading") or f"Card {s_idx + 1}"
        card_summaries.append({"id": c_id, "title": c_title})

    custom_system_prompt = (
        "You are an expert clinical medical educator assisting with NCLEX visual study notes.\n"
        "The user wants to add or update specific information based on a prompt.\n"
        "Your task is to select the single most relevant card from the provided list of cards, "
        "and generate exactly ONE concise, high-yield clinical bullet point answering the user's prompt.\n"
        "You MUST return a JSON object with:\n"
        "- 'target_card_id': The exact 'id' of the most relevant card from the cards list\n"
        "- 'target_card_title': The title of that chosen card\n"
        "- 'label': A short uppercase clinical tag (1-3 words, e.g. 'CLINICAL NOTE', 'ASSESSMENT', 'DIAGNOSIS', 'INTERVENTION', 'PATHOLOGY', 'RISK FACTOR')\n"
        "- 'text': Exactly 1-2 concise, clear, high-yield sentences providing the exact information requested.\n"
        "Do not regenerate other cards or rewrite the entire layout. Only return this single bullet point JSON object."
    )

    user_msg = (
        f"AVAILABLE CARDS IN VISUAL NOTES:\n{json.dumps(card_summaries, indent=2)}\n\n"
        f"USER MODIFICATION INSTRUCTION:\n\"{instruction}\"\n\n"
        "Identify the single most relevant card and return the bullet point JSON object."
    )

    try:
        openai_json = call_openai_api(custom_system_prompt, user_msg)
        if openai_json and isinstance(openai_json, dict) and (openai_json.get("label") or openai_json.get("text")):
            target_id = str(openai_json.get("target_card_id") or "").strip()
            target_title = str(openai_json.get("target_card_title") or "").strip().lower()
            new_label = clean_markdown_stars(str(openai_json.get("label") or "NOTE")).strip().upper()
            new_text = clean_markdown_stars(str(openai_json.get("text") or instruction)).strip()

            updated = json.loads(json.dumps(plan))
            up_sections = updated.get("sections") or updated.get("cards") or []

            target_sec = None
            if target_id:
                target_sec = next((s for s in up_sections if str(s.get("id") or "") == target_id), None)
            if not target_sec and target_title:
                target_sec = next((s for s in up_sections if target_title in (s.get("title") or s.get("heading") or "").lower()), None)
            if not target_sec:
                target_sec = fuzzy_find_target_section(instruction, up_sections)
            if not target_sec and up_sections:
                target_sec = up_sections[0]

            if target_sec and isinstance(target_sec, dict):
                if not isinstance(target_sec.get("items"), list):
                    target_sec["items"] = []
                if not isinstance(target_sec.get("bullets"), list):
                    target_sec["bullets"] = []

                # Append the targeted bullet point to the related card
                target_sec["items"].append({
                    "label": new_label,
                    "text": new_text
                })
                target_sec["bullets"].append(new_text)

                # Guarantee card and image preservation
                updated["sections"] = up_sections
                updated["cards"] = up_sections
                updated["is_visual_generated"] = True
                return updated
    except Exception as e:
        print(f"[AI Edit Custom Point OpenAI Exception] {e}")

    return process_deterministic_ai_edit(plan, instruction)

def create_default_plan(filename: str = "Uploaded_Notes.pdf", text_content: str = "") -> dict:
    return build_modular_visual_dashboard_schema(filename, text_content)

def get_realistic_diagram_metadata(topic: str, text_content: str = "", filename: str = "") -> dict:
    """Generates metadata for document topic and resolves appropriate asset via pipeline."""
    clean = str(topic or "Medical Study Notes").strip().rsplit('.', 1)[0].replace("_", " ").replace("-", " ").title()
    doc_filename = filename or topic or "Medical_Study_Notes"
    doc_subject, doc_name = detect_primary_subject(clean, text_content)
    req = {
        "type": "anatomy",
        "is_anatomy": True,
        "subject": f"{doc_name} Anatomy",
        "purpose": f"Show anatomical structure of {doc_name} relevant to document",
        "required": True
    }
    url, gen, reused = resolve_or_generate_visual(req, text_content, filename=doc_filename)
    return {
        "illustration_suggested": True,
        "organ_system": f"{doc_name.upper()} & CLINICAL STUDY",
        "target_organ": doc_name,
        "subject": f"{doc_name} Clinical & Anatomical Concepts",
        "category": doc_subject,
        "style": "Clean Educational Medical Infographic",
        "asset_url": url,
        "generated": gen,
        "reused": reused,
        "pins": []
    }


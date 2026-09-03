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
ASSET_CATALOG: Dict[str, Dict[str, str]] = {
    "heart": {
        "organ_3d": "/assets/generated/heart_organ_3d.png",
        "cross_section": "/assets/generated/heart_cross_section.png",
        "primary": "/assets/heart_3d_organ.png",
        "pathway": "/assets/vessels_3d_cutaway.png"
    },
    "lungs": {
        "organ_3d": "/assets/lungs_3d_organ.png",
        "cross_section": "/assets/generated/lungs_cross_section.png",
        "primary": "/assets/lungs_3d_organ.png",
        "pathway": "/assets/generated/respiratory_cross_section.png"
    },
    "kidney": {
        "organ_3d": "/assets/generated/kidney_organ_3d.png",
        "cross_section": "/assets/generated/kidney_cross_section.png",
        "primary": "/assets/kidney_3d_organ.png",
        "pathway": "/assets/generated/kidney_and_nephron_organ_3d.png"
    },
    "liver": {
        "organ_3d": "/assets/generated/liver_organ_3d.png",
        "cross_section": "/assets/generated/liver_cross_section.png",
        "primary": "/assets/liver_3d_organ.png",
        "pathway": "/assets/liver_3d_organ.png"
    },
    "brain": {
        "organ_3d": "/assets/generated/brain_organ_3d.png",
        "cross_section": "/assets/generated/brain_cross_section.png",
        "primary": "/assets/brain_3d_organ.png",
        "pathway": "/assets/generated/brain_cross_section.png"
    },
    "stomach": {
        "organ_3d": "/assets/stomach_3d_organ.png",
        "cross_section": "/assets/stomach_3d_organ.png",
        "primary": "/assets/stomach_3d_organ.png",
        "pathway": "/assets/stomach_3d_organ.png"
    },
    "vessels": {
        "organ_3d": "/assets/vessels_3d_cutaway.png",
        "cross_section": "/assets/vessels_3d_cutaway.png",
        "primary": "/assets/vessels_3d_cutaway.png",
        "pathway": "/assets/vessels_3d_cutaway.png"
    },
    "pharmacology": {
        "organ_3d": "/assets/generated/pharmacology_organ_3d.png",
        "cross_section": "/assets/generated/pharmacology_cross_section.png",
        "primary": "/assets/medication.jpg",
        "pathway": "/assets/generated/pharmacology_cross_section.png"
    },
    "pathology": {
        "organ_3d": "/assets/generated/pathology_organ_3d.png",
        "cross_section": "/assets/generated/pathology_organ_3d.png",
        "primary": "/assets/generated/pathology_organ_3d.png",
        "pathway": "/assets/generated/pathology_organ_3d.png"
    }
}

def detect_primary_subject(title: str, text: str, sections: Optional[List[dict]] = None) -> Tuple[str, str]:
    """Detects primary organ or domain: Title -> Headings -> Document Body."""
    title_lower = (title or "").lower()
    headings_lower = " ".join([s.get("title", "") or s.get("heading", "") for s in (sections or [])]).lower()
    body_lower = (text or "").lower().replace("heart rate", "").replace("respiratory rate", "")

    patterns = [
        ("heart", "Heart", r"\b(heart|cardiac|cardiovascular|myocard|atrium|ventricle|aorta|coronary|pericard)\b"),
        ("lungs", "Lungs", r"\b(lung|lungs|pulmon|respir|ards|asthma|copd|pneumonia|alveol|bronch)\b"),
        ("kidney", "Kidneys", r"\b(kidney|renal|nephron|glomerul|aki|ckd|dialysis|creatinine|gfr)\b"),
        ("liver", "Liver", r"\b(liver|hepatic|biliary|gallbladder|hepatitis|cirrhosis|bilirubin)\b"),
        ("brain", "Brain", r"\b(brain|neuro|cerebr|cns|stroke|seizure|encephal|spinal cord|neuron)\b"),
        ("stomach", "Stomach", r"\b(stomach|gastric|peptic|gerd|ulcer|digestive|gut|gi tract|duoden)\b"),
        ("vessels", "Blood Vessels", r"\b(vessel|vascular|artery|vein|capillary|blood flow|circulation|dvt)\b"),
        ("pharmacology", "Pharmacology", r"\b(pharmacolog|medication|drug|drugs|dosage|antibiotic|therapy)\b"),
    ]

    for key, name, regex in patterns:
        if re.search(regex, title_lower):
            return key, name
    for key, name, regex in patterns:
        if re.search(regex, headings_lower):
            return key, name
    for key, name, regex in patterns:
        if re.search(regex, body_lower):
            return key, name
    return "heart", "Heart"

def resolve_visual_asset(subject: str, visual_type: str = "primary", is_medication: bool = False) -> str:
    """Smart Asset Resolver with asset-first reuse logic."""
    if is_medication:
        return ASSET_CATALOG["pharmacology"].get(visual_type, "/assets/medication.jpg")
    clean = (subject or "heart").lower().strip()
    entry = ASSET_CATALOG.get(clean) or next((v for k, v in ASSET_CATALOG.items() if k in clean or clean in k), None)
    if entry:
        return entry.get(visual_type) or entry.get("primary") or "/assets/heart_3d_organ.png"
    return "/assets/heart_3d_organ.png"


# --- AI System Prompts ---
MEDICAL_STUDY_NOTES_ARCHITECT_PROMPT = """# MEDICAL DOCUMENT → ORIGINAL NOTES + VISUAL NOTES
You are an expert AI Medical Visual Notes Architect.
RULES:
1. SOURCE FIDELITY & STRICT VERBATIM EXTRACTION:
   - The user's document is 100% the primary source of truth. Zero hallucinations or fake facts.
   - For all cards except Anatomy, extract text EXCLUSIVELY and VERBATIM from the uploaded document text.
2. STRICT SECTION BOUNDARIES & NO DATA LEAKAGE:
   - CARD 1 (OVERVIEW): Must contain ONLY general definitions and high-level introduction points. DO NOT include Pathophysiology steps, Diagnostic findings, or Nursing details inside Overview.
   - CARD 2 (ANATOMY): If the document lacks an Anatomy section, provide standard anatomical facts for the organ so the card is not empty. If the document has an Anatomy section, use its exact text.
   - CARD 3 (PATHOPHYSIOLOGY / PATHWAY): Extract the step-by-step disease mechanism DIRECTLY from the Pathophysiology section of the document. Build 'flows' with step 'name' and 'subtext' from the real document text.
   - CARD 4+ (DIAGNOSTIC, NURSING INTERVENTIONS, STAGES, TREATMENT, etc.): Create a separate card for EVERY major top-level section heading present in the document. DO NOT OMIT OR MERGE STAGES OR PHASES!
3. CARD LAYOUT:
   - ROW 1: [ CARD 1 (Overview + Organ 3D Visual) ] [ CARD 2 (Anatomy + Cross-Section Visual) ]
   - ROW 2: [ CARD 3 — FULL WIDTH (Pathway/Flowchart with steps & arrows, extracted from document) ]
   - ROW 3+: [ CARD 4 ] [ CARD 5 ] ... (2 cards per row for remaining source document sections)

RETURN VALID JSON:
{
  "header": {"title": "Document Title", "subtitle": "AI-converted Visual Study Notes", "topic": "Main Topic"},
  "sections": [
    {
      "id": "section_1",
      "title": "SECTION TITLE IN UPPERCASE",
      "heading": "SECTION TITLE IN UPPERCASE",
      "type": "tagged_items | flowchart",
      "section_type": "tagged_items | flowchart",
      "bullets": ["Bullet point text"],
      "items": [{"label": "Badge Title", "text": "Detail text"}],
      "flows": [{"label": "PATHWAY NAME", "steps": [{"name": "Step Title", "subtext": "Step Details"}]}],
      "image": "/assets/...",
      "accent": "blue | red | purple | green | amber"
    }
  ],
  "bottom_panels": {
    "clinical_tip": {"title": "Clinical Tip", "text": "Consideration text"},
    "remember_mnemonic": {"title": "Remember", "text": "Mnemonic summary"}
  }
}"""

SYSTEM_PROMPT = MEDICAL_STUDY_NOTES_ARCHITECT_PROMPT


# --- Deterministic Parser (1:1 Source Fidelity) ---
def clean_heading_title(line: str) -> str:
    return re.sub(r'^(?:[0-9]+[\.\)]|#{1,6}\s+|[IVXLCDM]+[\.\)])\s*', '', line).strip(':').strip()

def is_major_heading(line: str) -> bool:
    """Accurately detects ONLY true top-level section headers, avoiding false card splits on sub-bullets."""
    stripped = line.strip()
    if not stripped:
        return False

    # 1. Explicitly numbered/lettered/markdown section headings: "1. ...", "1) ...", "I. ...", "A. ...", "## ..."
    if re.match(r'^(?:[0-9]+[\.\)]|#{1,6}\s+|[IVXLCDM]+[\.\)]|[A-Z][\.\)])\s+[A-Za-z]', stripped):
        return True

    # Strip bullets & leading symbols
    clean = re.sub(r'^[•\-\*\→●▪■◆➢►○✔✓]\s*', '', stripped)
    clean = re.sub(r'^(?:[0-9]+[\.\)]|#{1,6}\s+|[IVXLCDM]+[\.\)])\s*', '', clean).strip(':').strip()
    clean_lower = clean.lower()

    if not clean or '→' in clean or '->' in clean or len(clean) > 60:
        return False

    # Skip sentence narrative text ending with periods
    if clean.endswith(('.', '?', ';')) and len(clean.split()) > 3:
        return False

    # Strict top-level section domain titles ONLY
    top_level_headings = [
        "overview", "definition", "introduction", "what is it",
        "anatomy", "structure", "physiology",
        "pathophysiology", "pathophysiology pathway", "pathway", "mechanism", "pathogenesis",
        "causes", "cause", "etiology", "risk factors",
        "diagnostic workup", "diagnostics", "diagnostic", "labs", "laboratory", "evaluation", "assessment",
        "stages and phases", "stages & phases", "stages", "phases", "staging", "classification",
        "clinical manifestations", "manifestations", "signs & symptoms", "signs and symptoms",
        "nursing interventions", "nursing care", "interventions",
        "treatment and medications", "treatment & medications", "treatment", "management", "pharmacology", "medications",
        "complications", "prevention"
    ]

    for th in top_level_headings:
        if clean_lower == th or clean_lower == f"{th}:":
            return True

    # Standalone major title with clear heading prefix/formatting
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
        elif canonical_title.strip() in ["TREATMENT", "MEDICATIONS", "MEDICATION", "PHARMACOLOGY", "TREATMENT AND MEDICATIONS", "TREATMENT & MEDICATIONS", "MANAGEMENT"]:
            canonical_title = "TREATMENT & MEDICATIONS"

        if canonical_title not in grouped:
            sec_type = sec.get("type") or sec.get("section_type") or "tagged_items"
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
                "image": sec.get("image") or ""
            }
            if sec.get("isAnatomyCard"):
                grouped[canonical_title]["isAnatomyCard"] = True
            if sec.get("organ_type"):
                grouped[canonical_title]["organ_type"] = sec.get("organ_type")
            ordered_keys.append(canonical_title)

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
                card["bullets"].append(b.strip())
                existing_bullets.add(b.strip().lower())

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

    # Combine AI-generated sections with dynamic sections to ensure NO section heading is ever lost or omitted
    if parsed_ai and isinstance(parsed_ai, dict) and "sections" in parsed_ai:
        ai_secs = parsed_ai.get("sections") or []
        ai_titles_clean = {re.sub(r'^\d+[\.\)]\s*', '', (s.get("title") or s.get("heading") or "")).strip().lower() for s in ai_secs if isinstance(s, dict)}
        for dsec in dynamic_sections:
            d_title_clean = re.sub(r'^\d+[\.\)]\s*', '', (dsec.get("title") or dsec.get("heading") or "")).strip().lower()
            if "overview" in d_title_clean:
                d_title_clean = "overview"
            if d_title_clean and not any((t == d_title_clean or (t in d_title_clean and len(t) > 3) or (d_title_clean in t and len(d_title_clean) > 3)) for t in ai_titles_clean):
                ai_secs.append(dsec)
                ai_titles_clean.add(d_title_clean)
        sections_to_use = ai_secs
    else:
        sections_to_use = dynamic_sections

    DEFAULT_ANATOMY_ITEMS = {
        "lungs": [
            {"label": "PRIMARY STRUCTURE", "text": "Paired respiratory organs in thoracic cavity separated by mediastinum."},
            {"label": "ALVEOLAR MEMBRANE", "text": "Alveoli capillaries form thin membrane for O2 & CO2 gas exchange."},
            {"label": "SURFACTANT LAYER", "text": "Phospholipid lining reducing surface tension to prevent atelectasis."},
            {"label": "PLEURAL SPACE", "text": "Visceral & parietal pleura lubricated by serous fluid for smooth lung expansion."}
        ],
        "heart": [
            {"label": "MYOCARDIUM", "text": "Thick muscular middle layer driving systemic & pulmonary circulation."},
            {"label": "FOUR CHAMBERS", "text": "Right/Left Atria and Ventricles regulating unidirectional blood flow."},
            {"label": "VALVULAR APPARATUS", "text": "Atrioventricular & semilunar valves preventing retrograde blood flow."},
            {"label": "PERICARDIAL SAC", "text": "Double-walled sac protecting cardiac tissue & cushioning movement."}
        ],
        "kidney": [
            {"label": "NEPHRON UNITS", "text": "Functional units filtering metabolic wastes & maintaining fluid balance."},
            {"label": "GLOMERULUS", "text": "Capillary bed performing high-pressure blood plasma filtration."},
            {"label": "RENAL CORTEX & MEDULLA", "text": "Outer filtration layer & inner pyramid region concentrating urine."},
            {"label": "JUXTAGLOMERULAR APPARATUS", "text": "Regulates blood pressure via renin secretion and sodium balance."}
        ],
        "liver": [
            {"label": "HEPATIC LOBULES", "text": "Functional units processing nutrients, toxins, and bile synthesis."},
            {"label": "HEPATOCYTES & SINUSOIDS", "text": "Primary metabolic cells exchanging nutrients with portal circulation."},
            {"label": "BILIARY TREE", "text": "Duct system collecting and transporting bile for lipid digestion."},
            {"label": "KUPFFER CELLS", "text": "Specialized hepatic macrophages clearing pathogens & cell debris."}
        ],
        "brain": [
            {"label": "CEREBRAL CORTEX", "text": "Outer neural layer responsible for memory, reasoning, language, and sensory integration."},
            {"label": "BRAINSTEM & CNS", "text": "Regulates autonomic functions (respiration, cardiac rhythm, vasomotor control) with spinal cord."},
            {"label": "MOTOR & SENSORY PATHWAYS", "text": "Coordinates voluntary/involuntary motor movements & interprets somatic sensory inputs."},
            {"label": "BLOOD-BRAIN BARRIER", "text": "Selective endothelial barrier protecting brain tissue & maintaining neural homeostasis."}
        ],
        "stomach": [
            {"label": "GASTRIC MUCOSA", "text": "Parietal & chief cells secreting HCl acid & pepsinogen for digestion."},
            {"label": "MUSCULARIS EXTERNA", "text": "Three smooth muscle layers enabling churning & mechanical breakdown."},
            {"label": "PYLORIC SPHINCTER", "text": "Regulates controlled passage of acidic chyme into duodenum."},
            {"label": "RUGAE FOLDS", "text": "Mucosal folds allowing gastric expansion during fluid/food intake."}
        ],
        "vessels": [
            {"label": "TUNICA INTIMA & ENDOTHELIUM", "text": "Vascular lining regulating vascular tone, permeability, and smooth blood flow."},
            {"label": "TUNICA MEDIA", "text": "Smooth muscle and elastic fibers regulating vasodilation and vasoconstriction."},
            {"label": "TUNICA ADVENTITIA", "text": "Outer connective tissue anchoring blood vessels to surrounding tissues."},
            {"label": "CAPILLARY BED", "text": "Microvascular network facilitating tissue oxygenation and nutrient exchange."}
        ]
    }

    DEFAULT_STAGES_ITEMS = {
        "lungs": [
            {"label": "STAGE 1: EXUDATIVE PHASE", "text": "Alveolar edema, capillary congestion, and neutrophil infiltration (Days 1-7)."},
            {"label": "STAGE 2: PROLIFERATIVE PHASE", "text": "Type II pneumocyte hyperplasia and cellular tissue repair (Days 7-21)."},
            {"label": "STAGE 3: FIBROTIC PHASE", "text": "Extensive pulmonary fibrosis, lung remodeling, and compliance loss (>21 Days)."}
        ],
        "heart": [
            {"label": "STAGE A: AT RISK", "text": "High risk for heart failure without structural disease or symptoms."},
            {"label": "STAGE B: PRE-HEART FAILURE", "text": "Structural heart disease present without clinical signs or symptoms."},
            {"label": "STAGE C: SYMPTOMATIC HF", "text": "Structural heart disease with past or current heart failure symptoms."},
            {"label": "STAGE D: ADVANCED HF", "text": "Marked refractory symptoms at rest requiring specialized interventions."}
        ],
        "kidney": [
            {"label": "STAGE 1: NORMAL / HIGH GFR", "text": "Kidney damage with normal or elevated GFR (≥90 mL/min/1.73m²)."},
            {"label": "STAGE 2: MILD DECREASE", "text": "Mild reduction in renal function (GFR 60-89 mL/min/1.73m²)."},
            {"label": "STAGE 3: MODERATE DECREASE", "text": "Moderate GFR decline (Stage 3a: 45-59, Stage 3b: 30-44 mL/min)."},
            {"label": "STAGE 4: SEVERE DECREASE", "text": "Severe loss of kidney function (GFR 15-29 mL/min/1.73m²)."},
            {"label": "STAGE 5: KIDNEY FAILURE", "text": "End-stage renal disease (ESRD) requiring dialysis or transplant (GFR <15)."}
        ],
        "liver": [
            {"label": "STAGE 1: INFLAMMATION", "text": "Early liver cell swelling and hepatic inflammation."},
            {"label": "STAGE 2: FIBROSIS", "text": "Scar tissue begins to form around liver tissue and portal triads."},
            {"label": "STAGE 3: CIRRHOSIS", "text": "Permanent irreversible nodular scarring altering hepatic vasculature."},
            {"label": "STAGE 4: LIVER FAILURE", "text": "End-stage hepatic decompensation requiring liver transplantation."}
        ],
        "brain": [
            {"label": "STAGE 1: EARLY / MILD", "text": "Initial cognitive decline, subtle memory loss, and focal symptoms."},
            {"label": "STAGE 2: MODERATE / PROGRESSIVE", "text": "Increasing disorientation, speech difficulty, and impaired motor control."},
            {"label": "STAGE 3: SEVERE / ADVANCED", "text": "Profound neural deficits, loss of autonomy, and autonomic instability."}
        ],
        "stomach": [
            {"label": "STAGE 1: MUCOSAL DAMAGE", "text": "Superficial gastric mucosal irritation and acid erosion."},
            {"label": "STAGE 2: ULCER FORMATION", "text": "Deeper submucosal erosion forming active gastric ulceration."},
            {"label": "STAGE 3: COMPLICATED DISEASE", "text": "Risk of gastric perforation, bleeding, or pyloric stenosis."}
        ],
        "vessels": [
            {"label": "STAGE 1: ENDOTHELIAL DAMAGE", "text": "Vascular wall shear stress and lipid streak accumulation."},
            {"label": "STAGE 2: PLAQUE FORMATION", "text": "Fibrous cap formation causing luminal arterial narrowing."},
            {"label": "STAGE 3: OCCLUSION / RUPTURE", "text": "Critical tissue ischemia, thrombosis, or plaque rupture."}
        ]
    }

    DEFAULT_PATHWAY_ITEMS = {
        "lungs": {
            "title": f"{doc_organ_name.upper()} PATHOPHYSIOLOGY PATHWAY",
            "heading": f"{doc_organ_name.upper()} PATHOPHYSIOLOGY PATHWAY",
            "type": "flowchart",
            "section_type": "flowchart",
            "flows": [{
                "label": f"{doc_organ_name.upper()} PATHWAY",
                "steps": [
                    {"name": "INJURY INITIATION", "subtext": "Direct or indirect pulmonary insult triggers systemic inflammatory release."},
                    {"name": "ENDOTHELIAL DAMAGE", "subtext": "Neutrophil infiltration damages alveolar-capillary membrane barrier."},
                    {"name": "ALVEOLAR FLUID EXUDATE", "subtext": "Protein-rich exudative fluid fills alveoli and inactivates surfactant."},
                    {"name": "ATELECTASIS & HYPOXEMIA", "subtext": "Alveolar collapse leads to severe V/Q mismatch and refractory hypoxemia."}
                ]
            }],
            "accent": "purple"
        },
        "heart": {
            "title": "CARDIAC BLOOD FLOW PATHWAY",
            "heading": "CARDIAC BLOOD FLOW PATHWAY",
            "type": "flowchart",
            "section_type": "flowchart",
            "flows": [{
                "label": "CARDIAC CIRCULATION PATHWAY",
                "steps": [
                    {"name": "RIGHT ATRIUM", "subtext": "Receives deoxygenated venous return from vena cava."},
                    {"name": "RIGHT VENTRICLE", "subtext": "Pumps deoxygenated blood through pulmonary artery into lungs."},
                    {"name": "LEFT ATRIUM", "subtext": "Receives oxygen-rich blood returning from pulmonary veins."},
                    {"name": "LEFT VENTRICLE & AORTA", "subtext": "Propels oxygenated blood into systemic arterial circulation."}
                ]
            }],
            "accent": "purple"
        },
        "kidney": {
            "title": "RENAL FILTRATION PATHWAY",
            "heading": "RENAL FILTRATION PATHWAY",
            "type": "flowchart",
            "section_type": "flowchart",
            "flows": [{
                "label": "NEPHRON URINE FORMATION PATHWAY",
                "steps": [
                    {"name": "GLOMERULAR FILTRATION", "subtext": "High-pressure plasma filtration at Bowman capsule."},
                    {"name": "PROXIMAL REABSORPTION", "subtext": "Reabsorption of glucose, amino acids, water, and sodium."},
                    {"name": "LOOP OF HENLE COUNTERCURRENT", "subtext": "Establishes osmotic gradient concentrating tubular fluid."},
                    {"name": "COLLECTING DUCT EXCRETION", "subtext": "Final water reabsorption under ADH control and urine excretion."}
                ]
            }],
            "accent": "purple"
        },
        "liver": {
            "title": "HEPATIC CIRCULATION PATHWAY",
            "heading": "HEPATIC CIRCULATION PATHWAY",
            "type": "flowchart",
            "section_type": "flowchart",
            "flows": [{
                "label": "HEPATIC PROCESSING PATHWAY",
                "steps": [
                    {"name": "PORTAL INFLOW", "subtext": "Dual blood supply from hepatic artery and portal vein enters sinusoids."},
                    {"name": "HEPATOCYTE CLEARANCE", "subtext": "Metabolic transformation, toxin neutralization, and glycogen storage."},
                    {"name": "BILE SYNTHESIS", "subtext": "Secretes bile acids and conjugated bilirubin into bile canaliculi."},
                    {"name": "HEPATIC VEIN OUTFLOW", "subtext": "Cleared blood drains to inferior vena cava while bile drains to gallbladder."}
                ]
            }],
            "accent": "purple"
        },
        "brain": {
            "title": "NEURAL REFLEX & CONDUCTION PATHWAY",
            "heading": "NEURAL REFLEX & CONDUCTION PATHWAY",
            "type": "flowchart",
            "section_type": "flowchart",
            "flows": [{
                "label": "CENTRAL NERVOUS SYSTEM PATHWAY",
                "steps": [
                    {"name": "SENSORY RECEPTOR INPUT", "subtext": "Periphery detects environmental stimulus and transmits via afferent nerve."},
                    {"name": "SPINAL & CEREBRAL INTEGRATION", "subtext": "Interneurons process signal in cerebral cortex and thalamic relay."},
                    {"name": "EFFERENT MOTOR COMMAND", "subtext": "Motor cortex dispatches impulse down corticospinal tract."},
                    {"name": "TARGET EFFECTOR RESPONSE", "subtext": "Neuromuscular junction activates target muscle or autonomic gland."}
                ]
            }],
            "accent": "purple"
        },
        "stomach": {
            "title": "GASTRIC DIGESTIVE PATHWAY",
            "heading": "GASTRIC DIGESTIVE PATHWAY",
            "type": "flowchart",
            "section_type": "flowchart",
            "flows": [{
                "label": "GASTRIC SECRETION & MOTILITY PATHWAY",
                "steps": [
                    {"name": "CEPHALIC & GASTRIC PHASE", "subtext": "Vagal stimulation & gastrin release trigger HCl acid secretion."},
                    {"name": "MECHANICAL CHURNING", "subtext": "Muscularis externa contractions mix food into acidic chyme."},
                    {"name": "PROTEIN DIGESTION", "subtext": "Pepsin breaks down proteins in low pH environment."},
                    {"name": "PYLORIC EMPTYING", "subtext": "Controlled release of chyme through pyloric sphincter into duodenum."}
                ]
            }],
            "accent": "purple"
        },
        "vessels": {
            "title": "VASCULAR HEMODYNAMIC PATHWAY",
            "heading": "VASCULAR HEMODYNAMIC PATHWAY",
            "type": "flowchart",
            "section_type": "flowchart",
            "flows": [{
                "label": "VASCULAR PERFUSION PATHWAY",
                "steps": [
                    {"name": "ARTERIAL INFLOW", "subtext": "High pressure oxygenated surge delivered from aorta into major arteries."},
                    {"name": "ARTERIOLAR RESISTANCE", "subtext": "Smooth muscle tone regulates peripheral resistance and blood pressure."},
                    {"name": "CAPILLARY EXCHANGE", "subtext": "Microvascular diffusion delivers O2 and nutrients while picking up CO2."},
                    {"name": "VENOUS RETURN", "subtext": "Low pressure return facilitated by skeletal muscle pumps and valves."}
                ]
            }],
            "accent": "purple"
        }
    }

    # Check Anatomy rule
    has_anatomy = False
    for sec in sections_to_use:
        t_low = (sec.get("title", "") or sec.get("heading", "")).lower()
        if "anatomy" in t_low or ("structure" in t_low and not any(k in t_low for k in ["cell", "protect", "neuron"])):
            has_anatomy = True
            sec["isAnatomyCard"] = True
            if not sec.get("items") or len(sec.get("items")) == 0 or all(not (it.get("text") or "").strip() for it in sec.get("items")):
                sec["items"] = DEFAULT_ANATOMY_ITEMS.get(doc_subject, DEFAULT_ANATOMY_ITEMS["lungs"])
            break

    if not has_anatomy:
        anatomy_sec = {
            "id": "section_anatomy",
            "title": f"ANATOMY OF THE {doc_organ_name.upper()}",
            "heading": f"ANATOMY OF THE {doc_organ_name.upper()}",
            "type": "tagged_items",
            "section_type": "tagged_items",
            "bullets": [f"Anatomical organization of {doc_organ_name}."],
            "items": DEFAULT_ANATOMY_ITEMS.get(doc_subject, DEFAULT_ANATOMY_ITEMS["lungs"]),
            "flows": [],
            "isAnatomyCard": True,
            "organ_type": doc_subject,
            "accent": "blue"
        }
        if len(sections_to_use) >= 1:
            sections_to_use.insert(1, anatomy_sec)
        else:
            sections_to_use.append(anatomy_sec)

    # Check Stages fallback
    for sec in sections_to_use:
        t_low = (sec.get("title", "") or sec.get("heading", "")).lower()
        if any(k in t_low for k in ["stage", "phase"]):
            if not sec.get("items") or len(sec.get("items")) == 0 or all(not (it.get("text") or "").strip() for it in sec.get("items")):
                sec["items"] = DEFAULT_STAGES_ITEMS.get(doc_subject, DEFAULT_STAGES_ITEMS["lungs"])

    # Sort into Layout Order: Card 1 (Overview) -> Card 2 (Anatomy) -> Card 3 (Pathway) -> Cards 4+
    overview_card, anatomy_card = None, None
    pathway_cards, other_cards = [], []

    for idx, sec in enumerate(sections_to_use):
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

    # 1. Sanitize Overview card items to keep Overview card clean & concise (2-4 items max)
    if overview_card:
        leakage_labels = {"pathophysiology", "patho", "diagnostic", "nursing", "intervention", "treatment", "medication", "cause", "causes", "stage", "stages"}
        leakage_phrases = ["chest xray", "white out", "bronchoscopy", "blood culture", "abg blood", "v/q mismatch", "stiff lungs", "hypovolemia", "kayexalate"]
        
        clean_overview_items = []
        for it in overview_card.get("items", []):
            if isinstance(it, dict):
                lbl = (it.get("label") or "").strip().lower()
                txt = (it.get("text") or "").strip().lower()
                is_leaked = any(kl in lbl for kl in leakage_labels) or any(kp in txt for kp in leakage_phrases)
                if not is_leaked:
                    clean_overview_items.append(it)
        if clean_overview_items:
            overview_card["items"] = clean_overview_items[:4]
        else:
            dyn_overview = next((ds for ds in dynamic_sections if "overview" in (ds.get("heading") or "").lower()), None)
            if dyn_overview and dyn_overview.get("items"):
                overview_card["items"] = [it for it in dyn_overview["items"] if isinstance(it, dict) and not any(kl in (it.get("label") or "").lower() for kl in leakage_labels)][:4]

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

    if not pathway_cards:
        fallback_pw = DEFAULT_PATHWAY_ITEMS.get(doc_subject, DEFAULT_PATHWAY_ITEMS["lungs"])
        pathway_cards.append({
            "id": "section_pathway_fallback",
            "title": fallback_pw["title"],
            "heading": fallback_pw["heading"],
            "type": "flowchart",
            "section_type": "flowchart",
            "flows": fallback_pw["flows"],
            "bullets": [],
            "items": [],
            "accent": "purple"
        })

    ordered = [s for s in [overview_card, anatomy_card] if s and isinstance(s, dict)]
    ordered.extend([s for s in pathway_cards if s and isinstance(s, dict)])
    ordered.extend([s for s in other_cards if s and isinstance(s, dict)])

    # Asset Resolution for Cards
    for idx, sec in enumerate(ordered, start=1):
        if not sec or not isinstance(sec, dict):
            continue
        t_low = (sec.get("title", "") or sec.get("heading", "")).lower()
        if any(k in t_low for k in ["medication", "drug", "pharmacolog", "treatment", "therapy"]):
            sec["image"] = resolve_visual_asset(doc_subject, visual_type="primary", is_medication=True)
        elif sec.get("isAnatomyCard") or "anatomy" in t_low:
            sec["image"] = resolve_visual_asset(doc_subject, visual_type="cross_section")
        elif idx == 1 or "overview" in t_low:
            sec["image"] = resolve_visual_asset(doc_subject, visual_type="organ_3d")
        elif sec.get("type") == "flowchart" or sec.get("flows") or "pathway" in t_low or "patho" in t_low:
            sec["image"] = resolve_visual_asset(doc_subject, visual_type="pathway")

    clean_sections = sanitize_and_group_sections(ordered)

    bottom_panels = {
        "clinical_tip": {"title": "Clinical Tip", "text": f"Monitor hemodynamic status and organ perfusion parameters when assessing {doc_organ_name.lower()} function."},
        "remember_mnemonic": {"title": "Remember", "text": f"Recall the sequential blood flow and anatomical landmarks of the {doc_organ_name.lower()} during patient evaluations."}
    }
    if parsed_ai and isinstance(parsed_ai, dict) and "bottom_panels" in parsed_ai:
        if isinstance(parsed_ai["bottom_panels"], dict):
            bottom_panels.update(parsed_ai["bottom_panels"])

    return {
        "id": f"doc_{uuid.uuid4().hex[:8]}",
        "filename": filename,
        "topic": topic_name,
        "organ_subject": doc_subject,
        "organ_name": doc_organ_name,
        "original_text": raw_text.strip(),
        "sections": clean_sections,
        "cards": clean_sections,
        "bottom_panels": bottom_panels,
        "header": {
            "title": f"{topic_name} Visual Notes",
            "subtitle": "AI Visual Notes System • 1:1 Synchronized Study Sheet",
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

def build_visual_notes_from_text(file_name: str, raw_text: str) -> Dict[str, Any]:
    """Builds visual study notes schema using OpenAI API or deterministic engine."""
    truncated = raw_text[:15000] if len(raw_text) > 15000 else raw_text
    user_msg = f"Document Filename: {file_name}\n\nRAW STUDY NOTES CONTENT:\n{truncated}"
    openai_json = call_openai_api(MEDICAL_STUDY_NOTES_ARCHITECT_PROMPT, user_msg)
    if openai_json and isinstance(openai_json, dict):
        return build_modular_visual_dashboard_schema(file_name, raw_text, openai_json)
    return build_modular_visual_dashboard_schema(file_name, raw_text)


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

- Identify cards/sections containing too much information.
- Reduce long paragraphs into short bullets.
- Prefer 2–3 concise points where appropriate.
- Shorten unnecessarily long sentences.
- Keep important information.
- Remove repetition.
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

Only simplify the language:

- Rewrite complex medical, anatomical, or scientific jargon into clear, easy-to-read, plain English suitable for 1st-year beginner students.
- Whenever a complex medical term is used (e.g. Dyspnea, Pathophysiology, Atelectasis, Infarction, Etiology), keep the term but ALWAYS follow it with a clear parenthetical explanation (e.g. "Dyspnea (shortness of breath)", "Etiology (Causes)", "Hypoxemia (low blood oxygen)").
- Break down dense multi-clause sentences into short, simple sentences and clear bullet points.
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


def extract_color_from_instruction(instruction: str) -> str:
    """Extracts whatever light color phrase the user requested without static word lists."""
    text = instruction.lower().strip()
    for phrase in [
        "change the color theme to", "change color theme to", "change the color to",
        "change color to", "change theme to", "change color theme", "make it",
        "set theme to", "color theme", "theme color", "convert to", "switch to",
        "please change to", "theme", "color"
    ]:
        text = text.replace(phrase, "")
    clean = re.sub(r'[^a-z0-9\s\-]', '', text).strip()
    return clean if clean else "pastel pink"


MEDICAL_EXPLANATION_KB = {
    "alveolar membrane": "Thin respiratory barrier (~0.5 μm) separating alveolar air from pulmonary capillary blood, facilitating rapid O2 and CO2 gas diffusion.",
    "alveolar": "Air sac units in the lungs where oxygen enters the bloodstream and carbon dioxide is removed.",
    "membrane": "Cellular respiratory membrane facilitating exchange of respiratory gases between alveoli and capillaries.",
    "surfactant": "Lipoprotein fluid produced by Type II pneumocytes that lowers alveolar surface tension, preventing alveolar collapse (atelectasis).",
    "sepsis": "Systemic inflammatory response to severe infection causing widespread endothelial damage, vasodilation, and microvascular fluid leaks.",
    "atelectasis": "Partial or complete collapse of lung alveoli, impairing gas exchange and decreasing lung compliance.",
    "peep": "Positive End-Expiratory Pressure: Mechanical ventilation setting that maintains positive airway pressure during exhalation to keep alveoli open.",
    "hypoxemia": "Abnormally low partial pressure of oxygen in arterial blood, leading to tissue hypoxia.",
    "permeability": "Microvascular leakiness allowing fluid and plasma proteins to escape into interstitial and alveolar spaces.",
    "exudate": "Protein-rich inflammatory fluid that leaks into tissue or alveoli during acute injury.",
    "capillary": "Microscopic blood vessels where oxygen, nutrients, and waste products are exchanged between blood and tissues."
}

SIMPLIFIED_STUDENT_VOCAB = {
    "dyspnea": "shortness of breath (dyspnea)",
    "hypoxemia": "low blood oxygen levels (hypoxemia)",
    "hypoxia": "low tissue oxygen (hypoxia)",
    "edema": "fluid swelling (edema)",
    "tachycardia": "rapid heart rate (tachycardia)",
    "bradycardia": "slow heart rate (bradycardia)",
    "atelectasis": "air sac collapse (atelectasis)",
    "cyanosis": "bluish skin from low oxygen (cyanosis)",
    "surfactant": "fluid that keeps air sacs open (surfactant)",
    "permeability": "vessel leakiness",
    "exudate": "protein-rich fluid leak",
    "etiology": "underlying cause",
    "pathogenesis": "disease process step-by-step",
    "manifestations": "signs & symptoms",
    "ischemia": "lack of blood flow (ischemia)",
    "necrosis": "tissue death (necrosis)",
    "analgesic": "pain-relieving medicine",
    "antipyretic": "fever-reducing medicine",
    "hypertension": "high blood pressure",
    "hypotension": "low blood pressure",
    "auscultation": "listening with stethoscope",
    "perfusion": "passage of blood to tissues",
    "ventilation": "movement of air in/out of lungs"
}

def simplify_text_for_students(text: str) -> str:
    """Simplifies complex medical terms and formats text into 1st-year student friendly language without truncating points."""
    if not text or not isinstance(text, str):
        return ""
    clean = clean_markdown_stars(re.sub(r'<[^>]+>', '', text)).strip()

    # Replace complex medical terms with beginner-friendly descriptions across full text
    words = clean.split()
    for i, w in enumerate(words):
        w_clean = re.sub(r'[^a-zA-Z]', '', w).lower()
        if w_clean in SIMPLIFIED_STUDENT_VOCAB:
            words[i] = w.replace(re.sub(r'[^a-zA-Z]', '', w), SIMPLIFIED_STUDENT_VOCAB[w_clean])
    
    return " ".join(words)

def extract_elaboration_from_document_text(plan: dict, target_sec: dict, instruction: str = "") -> Tuple[str, str]:
    """Extracts real 1:1 clinical sentences from uploaded document text or provides clear clinical elaboration for custom prompts."""
    raw_text = (plan.get("original_text") or plan.get("raw_text") or plan.get("document_text") or "").strip()
    sec_title = (target_sec.get("title") or target_sec.get("heading") or "Clinical Section").strip()
    clean_sec_title = re.sub(r'^\d+[\.\)]\s*', '', sec_title).strip()

    # Clean instruction to identify user's requested topic
    inst_clean = instruction.lower()
    for phrase in ["what is", "tell me about", "i want to know more in details", "i want to know more", "details", "explain", "describe", "more about", "in detail"]:
        inst_clean = inst_clean.replace(phrase, "")
    query_topic = inst_clean.strip(" .?!,")

    # Keywords to search in document text
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
    elab_label = f"{label_topic} (CLINICAL DETAILS)"

    if extracted_sentences:
        elab_text = " ".join(extracted_sentences)
    else:
        # Check medical explanation KB for queried terms
        found_explanation = None
        for key, exp in MEDICAL_EXPLANATION_KB.items():
            if key in inst_clean or key in query_topic:
                found_explanation = exp
                break
        
        if found_explanation:
            elab_text = found_explanation
        else:
            elab_text = f"Detailed physiological and anatomical mechanisms regarding {query_topic or clean_sec_title}."

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

    # Rule 3: Concise / Shorter (Keeps 3-4 concise items/bullets/flows per card)
    if any(k in text_lower for k in ["concise", "shorter", "brief", "summarize", "short"]):
        for sec in sections:
            if isinstance(sec, dict):
                # 1. Truncate items to max 4 items, each max 1 sentence
                if sec.get("items") and isinstance(sec["items"], list):
                    sec["items"] = sec["items"][:4]
                    for it in sec["items"]:
                        if isinstance(it, dict) and it.get("text"):
                            sentences = [s.strip() for s in re.split(r'\.\s+', str(it["text"])) if s.strip()]
                            if sentences:
                                first_s = sentences[0]
                                if not first_s.endswith('.'):
                                    first_s += '.'
                                it["text"] = first_s

                # 2. Truncate bullets to max 4 bullets, each max 1 sentence
                if sec.get("bullets") and isinstance(sec["bullets"], list):
                    short_bullets = []
                    for b in sec["bullets"][:4]:
                        if isinstance(b, str) and b.strip():
                            sentences = [s.strip() for s in re.split(r'\.\s+', b) if s.strip()]
                            if sentences:
                                first_s = sentences[0]
                                if not first_s.endswith('.'):
                                    first_s += '.'
                                short_bullets.append(first_s)
                    sec["bullets"] = short_bullets

                # 3. Truncate flows to max 4 steps
                if sec.get("flows") and isinstance(sec["flows"], list):
                    sec["flows"] = sec["flows"][:4]

                # 4. Truncate content to max 1 sentence
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
                if sec.get("items") and isinstance(sec["items"], list):
                    for it in sec["items"]:
                        if isinstance(it, dict) and it.get("text"):
                            it["text"] = simplify_text_for_students(it["text"])
                        if isinstance(it, dict) and it.get("label"):
                            it["label"] = simplify_text_for_students(it["label"])
                if sec.get("bullets") and isinstance(sec["bullets"], list):
                    sec["bullets"] = [simplify_text_for_students(b) for b in sec["bullets"] if isinstance(b, str) and b.strip()]
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

            if target_sec.get("items") and isinstance(target_sec["items"], list):
                if not any(it.get("label") == elab_label for it in target_sec["items"] if isinstance(it, dict)):
                    target_sec["items"].append({
                        "label": elab_label,
                        "text": elab_text
                    })

    if "sections" in updated:
        updated["sections"] = sanitize_and_group_sections(updated["sections"])
        updated["cards"] = updated["sections"]

    return updated


def apply_ai_edit(plan: dict, instruction: str) -> dict:
    """Applies user's modification prompt to Visual Notes JSON schema dynamically using OpenAI API, falling back to deterministic engine."""
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
    
    compact_plan = {
        "title": plan.get("title", ""),
        "sections": [
            {
                "id": sec.get("id"),
                "title": sec.get("title") or sec.get("heading", ""),
                "type": sec.get("type") or sec.get("section_type", ""),
                "items": sec.get("items", []),
                "bullets": sec.get("bullets", []),
                "flows": sec.get("flows", [])
            }
            for sec in sections if isinstance(sec, dict)
        ]
    }

    user_msg = f"""CURRENT VISUAL NOTES JSON SCHEMA:
{json.dumps(compact_plan, indent=2)}

USER MODIFICATION INSTRUCTION:
"{instruction}"

Please return the updated valid JSON object with matching "sections" structure after applying the requested modification."""

    try:
        openai_json = call_openai_api(AI_EDIT_SYSTEM_PROMPT, user_msg)
        if openai_json and isinstance(openai_json, dict) and ("sections" in openai_json or "cards" in openai_json):
            new_secs = openai_json.get("sections") or openai_json.get("cards") or []
            if new_secs and isinstance(new_secs, list):
                updated = json.loads(json.dumps(plan))
                updated["sections"] = sanitize_and_group_sections(new_secs)
                updated["cards"] = updated["sections"]
                if openai_json.get("theme"):
                    updated["theme"] = openai_json["theme"]
                return updated
    except Exception as e:
        print(f"[AI Edit OpenAI Exception] {e}")

    return process_deterministic_ai_edit(plan, instruction)

def create_default_plan(filename: str = "Uploaded_Notes.pdf", text_content: str = "") -> dict:
    return build_modular_visual_dashboard_schema(filename, text_content)

def get_realistic_diagram_metadata(topic: str, text_content: str = "") -> dict:
    """Generates metadata for document topic and resolves appropriate asset."""
    clean = str(topic or "Medical Study Notes").strip().rsplit('.', 1)[0].replace("_", " ").replace("-", " ").title()
    doc_subject, doc_name = detect_primary_subject(clean, text_content)
    return {
        "illustration_suggested": True,
        "organ_system": f"{doc_name.upper()} & CLINICAL STUDY",
        "target_organ": doc_name,
        "subject": f"{doc_name} Clinical & Anatomical Concepts",
        "category": doc_subject,
        "style": "Clean Educational Medical Infographic",
        "asset_url": resolve_visual_asset(doc_subject, visual_type="cross_section"),
        "pins": []
    }


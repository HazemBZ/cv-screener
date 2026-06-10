#!/usr/bin/env python3
"""
CV Screener — Batch PDF CV extraction + scoring pipeline.

Extracts candidate info from text and scanned PDFs using OCR,
scores against configurable criteria, and outputs an Excel spreadsheet.

Usage:
    docker build -t cv-screener .
    docker run --rm -v /path/to/pdfs:/input -v /path/to/output:/output cv-screener \
        --input /input --output /output/results.xlsx [--criteria /app/criteria.yaml]
"""

import argparse
import csv
import io
import os
import re
import subprocess
import sys
import tempfile
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Optional

import pandas as pd
import yaml
from pdf2image import convert_from_path
from PIL import Image
# ---------------------------------------------------------------------------
# Configure local tesseract if bundled alongside the script
# ---------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_TESSERACT_BIN = os.path.join(_SCRIPT_DIR, "tesseract", "bin", "tesseract")
_TESSDATA = os.path.join(_SCRIPT_DIR, "tesseract", "share", "tessdata")
_TESSERACT_LIB = os.path.join(_SCRIPT_DIR, "tesseract", "lib")

import pytesseract
if os.path.exists(_TESSERACT_BIN):
    pytesseract.pytesseract.tesseract_cmd = _TESSERACT_BIN
    os.environ["TESSDATA_PREFIX"] = _TESSDATA
    os.environ.setdefault("LD_LIBRARY_PATH", "")
    os.environ["LD_LIBRARY_PATH"] = _TESSERACT_LIB + ":" + os.environ["LD_LIBRARY_PATH"]


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class CVEvaluation:
    file_name: str = ""
    candidate_name: str = ""
    email: str = ""
    phone: str = ""
    skills_found: str = ""
    skills_score: int = 0
    years_experience: float = 0.0
    experience_score: int = 0
    education: str = ""
    education_score: int = 0
    current_role: str = ""
    role_score: int = 0
    overall_score: int = 0
    notes: str = ""

    def to_row(self, columns: list[str]) -> list:
        mapping = {
            "File Name": self.file_name,
            "Candidate Name": self.candidate_name,
            "Email": self.email,
            "Phone": self.phone,
            "Skills Found": self.skills_found,
            "Skills Score": self.skills_score,
            "Years Experience": self.years_experience,
            "Experience Score": self.experience_score,
            "Education": self.education,
            "Education Score": self.education_score,
            "Current Role": self.current_role,
            "Role Score": self.role_score,
            "Overall Score": self.overall_score,
            "Notes": self.notes,
        }
        return [mapping.get(col, "") for col in columns]


# ---------------------------------------------------------------------------
# Text extraction
# ---------------------------------------------------------------------------

def extract_text_pdftotext(pdf_path: str) -> str:
    """Extract text via pdftotext (fast, for text-based PDFs)."""
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", pdf_path, "-"],
            capture_output=True, text=True, timeout=30
        )
        return result.stdout or ""
    except Exception:
        return ""


def extract_text_ocr(pdf_path: str, dpi: int = 300) -> str:
    """Extract text via OCR (for scanned/image PDFs)."""
    texts: list[str] = []
    try:
        images = convert_from_path(pdf_path, dpi=dpi, fmt="jpeg")
        for img in images:
            text = pytesseract.image_to_string(img)
            texts.append(text)
    except Exception as e:
        return f"[OCR ERROR: {e}]"
    return "\n".join(texts)


def extract_text(pdf_path: str) -> str:
    """
    Extract text from a PDF.
    Tries pdftotext first (fast). Falls back to OCR if result is too sparse.
    """
    text = extract_text_pdftotext(pdf_path)
    # If text extraction yielded very little, it's likely scanned → OCR
    if len(text.strip()) < 50:
        text = extract_text_ocr(pdf_path)
    return text


# ---------------------------------------------------------------------------
# Information extraction (regex-based)
# ---------------------------------------------------------------------------

EMAIL_RE = re.compile(r'[\w.+-]+@[\w-]+\.[\w.]+', re.IGNORECASE)

PHONE_RE = re.compile(
    r'(?:\+?\d{1,3}[-.\s]?)?\(?\d{2,4}\)?[-.\s]?\d{3,4}[-.\s]?\d{3,4}(?:\s*(?:ext|×|#)\s*\d+)?',
    re.IGNORECASE
)

EXPERIENCE_RE = re.compile(
    r'(\d+)\+?\s*(?:years?|yrs?|y(?:ear)?s?)\s*(?:of\s+)?(?:experience|exp|work)',
    re.IGNORECASE
)

DEGREE_KEYWORDS = [
    "bachelor", "b.s.", "b.tech", "b.e.", "b.a.",
    "master", "m.s.", "m.tech", "m.e.", "m.b.a.", "mba",
    "phd", "ph.d.", "doctorate", "doctor of",
    "associate", "diploma", "high school",
]


def find_email(text: str) -> str:
    m = EMAIL_RE.search(text)
    return m.group(0) if m else ""


def find_phone(text: str) -> str:
    m = PHONE_RE.search(text)
    return m.group(0).strip() if m else ""


def find_experience(text: str) -> float:
    """Extract years of experience from text."""
    matches = EXPERIENCE_RE.findall(text)
    if matches:
        # Take the highest number mentioned
        years = max(int(y) for y in matches)
        return float(years)
    # Fallback: look for year ranges like "2018 - 2024" implying duration
    year_range = re.findall(r'(19\d\d|20\d\d)\s*[-–to]+\s*(19\d\d|20\d\d|present|now)', text, re.IGNORECASE)
    if year_range:
        durations = []
        for start, end in year_range:
            end_num = 2026 if end.lower() in ("present", "now") else int(end)
            durations.append(end_num - int(start))
        if durations:
            return float(max(durations))
    return 0.0


def find_education(text: str) -> str:
    """Extract highest education level mentioned."""
    lines = text.splitlines()
    degrees_found: list[str] = []
    for line in lines:
        lower = line.lower()
        for kw in DEGREE_KEYWORDS:
            if kw in lower:
                degrees_found.append(line.strip())
                break
    if degrees_found:
        # Return the highest (last in list tends to be higher, but just return first found)
        return degrees_found[0][:120]
    return ""


def find_education_level(text: str) -> int:
    """Return numeric level: 0=unknown, 1=high_school, 2=bachelor, 3=master, 4=phd."""
    lower = text.lower()
    if any(kw in lower for kw in ["phd", "ph.d.", "doctorate", "doctor of"]):
        return 4
    if any(kw in lower for kw in ["master", "m.s.", "m.tech", "m.e.", "m.b.a.", "mba"]):
        return 3
    if any(kw in lower for kw in ["bachelor", "b.s.", "b.tech", "b.e.", "b.a."]):
        return 2
    if any(kw in lower for kw in ["associate", "diploma", "high school"]):
        return 1
    return 0


def find_current_role(text: str) -> str:
    """
    Heuristic: look for a line near "work experience" or after the header
    that looks like a job title.
    """
    lines = text.splitlines()
    # Strategy: find lines with common title indicators
    title_indicators = re.compile(
        r'(engineer|scientist|developer|analyst|manager|director|lead|'
        r'architect|specialist|consultant|intern|researcher|associate|'
        r'president|officer|head|chief|principal)',
        re.IGNORECASE
    )
    titles: list[str] = []
    capture = False
    for line in lines:
        stripped = line.strip()
        if not stripped or len(stripped) > 100:
            continue
        lower = stripped.lower()
        # Start capturing after "experience" header
        if re.search(r'(work\s+)?experience|employment|career|history', lower):
            capture = True
            continue
        if capture and title_indicators.search(stripped):
            titles.append(stripped)
        # Also capture lines that look like titles in the first ~30 lines
        if not capture and title_indicators.search(stripped):
            titles.append(stripped)

    return titles[0][:100] if titles else ""


def find_name(text: str) -> str:
    """
    Heuristic: the name is typically the first non-empty, non-boilerplate line
    that looks like a person's name (2-4 words, capitalized).
    """
    lines = text.splitlines()
    skip_patterns = re.compile(
        r'(curriculum\s*vitae|resume|cv|phone|email|address|linkedin|github|page)',
        re.IGNORECASE
    )
    name_pattern = re.compile(r'^[A-Z][a-záéíóúñüäöèà]*(?:\s+[A-Z][a-záéíóúñüäöèà]*){1,3}$')
    # Also match all-caps names (common in CVs)
    name_pattern_allcaps = re.compile(r'^[A-Z][A-Z\s]+$')

    for line in lines[:40]:  # Scan first 40 lines
        stripped = line.strip()
        if not stripped or len(stripped) > 60:
            continue
        if skip_patterns.search(stripped):
            continue
        if name_pattern.match(stripped) or name_pattern_allcaps.match(stripped):
            return stripped.strip()[:80]
    return ""


# ---------------------------------------------------------------------------
# Skills matching
# ---------------------------------------------------------------------------

def find_skills(text: str, skill_keywords: dict[str, int]) -> list[str]:
    """Find which configured skills appear in the CV text."""
    lower = text.lower()
    found: list[str] = []
    for skill in skill_keywords:
        # Word-boundary matching to avoid false positives
        pattern = r'\b' + re.escape(skill) + r'\b'
        if re.search(pattern, lower):
            found.append(skill)
    return found


# ---------------------------------------------------------------------------
# Scoring
# ---------------------------------------------------------------------------

def evaluate_cv(text: str, criteria: dict) -> CVEvaluation:
    weights = criteria.get("scoring_weights", {})
    skill_keywords = criteria.get("required_skills", {})
    min_exp = criteria.get("min_experience_years", 0)
    min_edu_level = {"high_school": 1, "bachelor": 2, "master": 3, "phd": 4}.get(
        criteria.get("min_education", "bachelor"), 0
    )
    preferred_roles = criteria.get("preferred_roles", [])
    qualifying_degrees = criteria.get("qualifying_degrees", [])
    output_columns = criteria.get("output_columns", [])

    result = CVEvaluation()

    # --- Basic info ---
    result.email = find_email(text)
    result.phone = find_phone(text)
    result.candidate_name = find_name(text)
    result.years_experience = find_experience(text)
    result.education = find_education(text)
    result.current_role = find_current_role(text)

    # --- Skills ---
    matched_skills = find_skills(text, skill_keywords)
    result.skills_found = ", ".join(matched_skills)
    if skill_keywords:
        max_skill_weight = sum(skill_keywords.values())
        earned = sum(skill_keywords[s] for s in matched_skills)
        result.skills_score = min(100, round(earned / max_skill_weight * 100)) if max_skill_weight else 0
    else:
        result.skills_score = 0

    # --- Experience ---
    edu_level = find_education_level(text)
    if min_exp > 0 and result.years_experience >= min_exp:
        result.experience_score = min(100, round((result.years_experience / (min_exp * 2)) * 100))
    elif min_exp > 0:
        result.experience_score = round((result.years_experience / min_exp) * 100)
    else:
        # No minimum: scale 0-20 years → 0-100
        result.experience_score = min(100, round(result.years_experience * 5))

    # --- Education ---
    if min_edu_level > 0 and edu_level >= min_edu_level:
        # Exceeds minimum → bonus
        result.education_score = min(100, 50 + (edu_level - min_edu_level) * 25)
    elif min_edu_level > 0:
        result.education_score = max(0, round((edu_level / min_edu_level) * 50))
    else:
        result.education_score = 50 if edu_level >= 2 else 20  # default

    # --- Role relevance ---
    lower_text = text.lower()
    role_matches = sum(1 for role in preferred_roles if role.lower() in lower_text)
    if preferred_roles:
        result.role_score = min(100, round((role_matches / len(preferred_roles)) * 100))
    else:
        result.role_score = 50

    # --- Overall ---
    total_weight = sum(weights.values()) or 100
    result.overall_score = round(
        (
            result.skills_score * weights.get("skills_match", 50) +
            result.experience_score * weights.get("experience", 25) +
            result.education_score * weights.get("education", 15) +
            result.role_score * weights.get("role_relevance", 10)
        ) / total_weight
    )

    # --- Notes for low-quality extractions ---
    notes_parts = []
    if not result.email:
        notes_parts.append("No email found")
    if not result.phone:
        notes_parts.append("No phone found")
    if not result.candidate_name:
        notes_parts.append("Name extraction uncertain")
    if result.years_experience == 0:
        notes_parts.append("Could not determine experience")
    if not result.education:
        notes_parts.append("No education found")
    result.notes = "; ".join(notes_parts)

    return result


# ---------------------------------------------------------------------------
# Batch processing
# ---------------------------------------------------------------------------

def process_one_pdf(pdf_path: str, criteria: dict) -> CVEvaluation:
    """Process a single PDF and return evaluation."""
    try:
        text = extract_text(pdf_path)
        ev = evaluate_cv(text, criteria)
        ev.file_name = os.path.basename(pdf_path)
        return ev
    except Exception as e:
        ev = CVEvaluation()
        ev.file_name = os.path.basename(pdf_path)
        ev.notes = f"ERROR: {e}"
        return ev


def process_batch(
    pdf_dir: str,
    criteria: dict,
    output_path: str,
    max_workers: int = 4,
) -> str:
    """Process all PDFs in a directory and write results to Excel."""
    pdf_dir = os.path.abspath(pdf_dir)
    output_path = os.path.abspath(output_path)

    # Find all PDFs
    pdf_files = sorted([
        os.path.join(pdf_dir, f)
        for f in os.listdir(pdf_dir)
        if f.lower().endswith(".pdf")
    ])

    if not pdf_files:
        print(f"❌ No PDF files found in {pdf_dir}", file=sys.stderr)
        sys.exit(1)

    print(f"📄 Found {len(pdf_files)} PDF(s) — starting extraction...")
    start = time.time()

    results: list[CVEvaluation] = []

    if max_workers > 1 and len(pdf_files) > 1:
        with ProcessPoolExecutor(max_workers=max_workers) as pool:
            futures = {
                pool.submit(process_one_pdf, pdf, criteria): pdf
                for pdf in pdf_files
            }
            done = 0
            total = len(futures)
            for f in as_completed(futures):
                done += 1
                pdf_name = os.path.basename(futures[f])
                try:
                    ev = f.result()
                    results.append(ev)
                except Exception as e:
                    ev = CVEvaluation()
                    ev.file_name = pdf_name
                    ev.notes = f"WORKER ERROR: {e}"
                    results.append(ev)
                print(f"  [{done}/{total}] {pdf_name} — score: {ev.overall_score}")
    else:
        for i, pdf in enumerate(pdf_files, 1):
            ev = process_one_pdf(pdf, criteria)
            results.append(ev)
            print(f"  [{i}/{len(pdf_files)}] {ev.file_name} — score: {ev.overall_score}")

    # Sort by overall score descending
    results.sort(key=lambda r: r.overall_score, reverse=True)

    # Build the CSV output (robust, no formatting issues)
    columns = criteria.get("output_columns", [])
    csv_rows = [columns]
    for ev in results:
        csv_rows.append(ev.to_row(columns))

    # Also produce Excel
    df = pd.DataFrame(csv_rows[1:], columns=csv_rows[0])
    df.to_excel(output_path, index=False, engine="openpyxl")

    elapsed = time.time() - start
    print(f"\n✅ Done in {elapsed:.1f}s — {len(pdf_files)} CVs processed")
    print(f"📊 Output: {output_path}")

    # Summary stats
    scores = [r.overall_score for r in results]
    print(f"📈 Score range: {min(scores)}–{max(scores)}  |  Mean: {sum(scores)/len(scores):.1f}")

    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def load_criteria(path: str) -> dict:
    """Load criteria from YAML file, with defaults for missing keys."""
    default = {
        "required_skills": {},
        "min_experience_years": 0,
        "min_education": "bachelor",
        "qualifying_degrees": [],
        "preferred_roles": [],
        "scoring_weights": {"skills_match": 50, "experience": 25, "education": 15, "role_relevance": 10},
        "output_columns": [
            "File Name", "Candidate Name", "Email", "Phone",
            "Skills Found", "Skills Score", "Years Experience",
            "Experience Score", "Education", "Education Score",
            "Current Role", "Role Score", "Overall Score", "Notes",
        ],
    }
    if path and os.path.exists(path):
        with open(path) as f:
            user = yaml.safe_load(f) or {}
            default.update(user)
    else:
        print(f"⚠️  Criteria file not found at '{path}', using built-in defaults")
    return default


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CV Screener — Batch PDF extraction + scoring to Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  %(prog)s --input ./cvs --output results.xlsx\n"
            "  %(prog)s --input ./cvs --output results.xlsx --criteria my_criteria.yaml\n"
            "  %(prog)s --input ./cvs --output results.xlsx --workers 8\n"
        ),
    )
    parser.add_argument("--input", "-i", required=True, help="Directory containing PDF CVs")
    parser.add_argument("--output", "-o", required=True, help="Output Excel file path (.xlsx)")
    parser.add_argument("--criteria", "-c", default="/app/criteria.yaml",
                        help="Path to criteria YAML config")
    parser.add_argument("--workers", "-w", type=int, default=4,
                        help="Number of parallel workers (default: 4)")
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    criteria = load_criteria(args.criteria)

    if not os.path.isdir(args.input):
        print(f"❌ Input directory not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    process_batch(
        pdf_dir=args.input,
        criteria=criteria,
        output_path=args.output,
        max_workers=args.workers,
    )


if __name__ == "__main__":
    main()

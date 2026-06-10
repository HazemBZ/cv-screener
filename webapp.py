#!/usr/bin/env python3
"""
CV Screener — Web UI. FastAPI app: upload PDFs -> scored results -> download Excel.
Usage: uv run uvicorn webapp:app --host 0.0.0.0 --port 8080
"""

import os
import sys
import tempfile
import time

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, Response, JSONResponse

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

from cv_screener import extract_text, evaluate_cv, load_criteria, CVEvaluation
import pandas as pd

app = FastAPI(title="CV Screener", version="0.2.0")

_latest_results: list[CVEvaluation] = []
_latest_criteria: dict = {}

_DEFAULT_CRITERIA = os.path.join(_SCRIPT_DIR, "criteria.yaml")

_HTML_PATH = os.path.join(_SCRIPT_DIR, "templates", "index.html")
with open(_HTML_PATH) as _fh:
    _INDEX_HTML = _fh.read()


def _load_default_criteria() -> dict:
    return load_criteria(_DEFAULT_CRITERIA)


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTMLResponse(content=_INDEX_HTML, status_code=200)


@app.post("/upload")
async def upload_files(files: list[UploadFile] = File(...)):
    """
    Upload PDF CVs, process them, return results as JSON.
    Accepts multiple PDF files. Processes synchronously.
    """
    global _latest_results, _latest_criteria

    if not files:
        raise HTTPException(400, "No files uploaded")

    criteria = _load_default_criteria()
    pdf_files = [f for f in files if f.filename and f.filename.lower().endswith(".pdf")]
    if not pdf_files:
        raise HTTPException(400, "No PDF files found in upload. Please upload PDF files only.")

    t0 = time.time()
    tmpdir = tempfile.mkdtemp(prefix="cv_screener_")
    saved_paths = []
    try:
        for f in pdf_files:
            dest = os.path.join(tmpdir, f.filename)
            content = await f.read()
            with open(dest, "wb") as fh:
                fh.write(content)
            saved_paths.append(dest)

        results: list[CVEvaluation] = []
        total = len(saved_paths)
        for i, path in enumerate(saved_paths, 1):
            ev = CVEvaluation()
            try:
                text = extract_text(path)
                ev = evaluate_cv(text, criteria)
            except Exception as e:
                ev.notes = f"ERROR: {e}"
            ev.file_name = os.path.basename(path)
            results.append(ev)

        results.sort(key=lambda r: r.overall_score, reverse=True)
        _latest_results = results
        _latest_criteria = criteria

        columns = criteria.get("output_columns", [])
        rows = []
        for ev in results:
            row = {}
            for col in columns:
                key = col.lower().replace(" ", "_")
                val = getattr(ev, key, "")
                row[col] = val
            rows.append(row)

        scores = [r.overall_score for r in results]
        elapsed = round(time.time() - t0, 2)
        summary = {
            "total": len(results),
            "mean": round(sum(scores) / len(scores), 1) if scores else 0,
            "min": min(scores) if scores else 0,
            "max": max(scores) if scores else 0,
            "time_s": elapsed,
        }

        return JSONResponse({"results": rows, "summary": summary})

    finally:
        for p in saved_paths:
            try:
                os.remove(p)
            except OSError:
                pass
        try:
            os.rmdir(tmpdir)
        except OSError:
            pass


@app.get("/download")
async def download_results():
    global _latest_results, _latest_criteria

    if not _latest_results:
        raise HTTPException(404, "No results available. Upload CVs first.")

    criteria = _latest_criteria
    columns = criteria.get("output_columns", [])
    rows = [ev.to_row(columns) for ev in _latest_results]

    df = pd.DataFrame(rows, columns=columns)

    output = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        df.to_excel(output.name, index=False, engine="openpyxl")
        output.seek(0)
        with open(output.name, "rb") as fh:
            data = fh.read()
    finally:
        os.unlink(output.name)

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=cv_screening_results.xlsx"},
    )


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "8080"))
    uvicorn.run("webapp:app", host="0.0.0.0", port=port, log_level="info")

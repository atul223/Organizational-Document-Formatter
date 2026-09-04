"""
app.py — FastAPI wrapper around the existing formatting pipeline.

This file does NOT modify, reimplement, or duplicate any pipeline logic.
It only:
  1. Accepts two uploaded .docx files (reference + target) over HTTP.
  2. Saves them to a private per-request temp directory.
  3. Calls format_doc.run(...) — the exact same function your
     `python format_doc.py --reference ... --input ... --output ...`
     CLI command calls — completely unchanged.
  4. Returns the resulting formatted .docx as a file download.

The pipeline code lives untouched in ./pipeline/ alongside this file.
"""
import os
import shutil
import sys
import tempfile
import traceback
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

# --- Make the untouched pipeline package importable (it uses flat imports
#     like `import policy_extractor`, so its own folder must be on sys.path) ---
PIPELINE_DIR = Path(__file__).parent / "pipeline"
sys.path.insert(0, str(PIPELINE_DIR))
import format_doc  # noqa: E402  (must import after sys.path insert)

APP_DIR = Path(__file__).parent
FRONTEND_DIR = APP_DIR.parent / "frontend"
JOBS_ROOT = Path(tempfile.gettempdir()) / "formatting_tool_jobs"
JOBS_ROOT.mkdir(exist_ok=True)

MAX_UPLOAD_BYTES = 25 * 1024 * 1024  # 25 MB per file — adjust as needed

app = FastAPI(title="Document Formatting Tool")

# Same-origin deployment (frontend served by this same app) needs no CORS,
# but this is left permissive+safe for an internal LAN tool. Tighten
# allow_origins to your actual internal URL if you split frontend/backend.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)


def _validate_docx(upload: UploadFile) -> None:
    if not upload.filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail=f"'{upload.filename}' is not a .docx file. Please upload Word .docx files only.",
        )


async def _save_upload(upload: UploadFile, dest: Path) -> None:
    size = 0
    with open(dest, "wb") as out:
        while True:
            chunk = await upload.read(1024 * 1024)
            if not chunk:
                break
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail="File too large.")
            out.write(chunk)


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/format")
async def format_document(
    reference: UploadFile = File(..., description="Reference document (formatting source)"),
    target: UploadFile = File(..., description="Document to be formatted"),
):
    """
    Runs the exact same pipeline as:
        python format_doc.py --reference <reference> --input <target> --output <result>
    and returns the resulting .docx file.
    """
    _validate_docx(reference)
    _validate_docx(target)

    job_id = uuid.uuid4().hex
    job_dir = JOBS_ROOT / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    reference_path = job_dir / "reference.docx"
    target_path = job_dir / "target.docx"
    output_path = job_dir / "formatted_output.docx"
    workdir = job_dir / "run_artifacts"

    try:
        await _save_upload(reference, reference_path)
        await _save_upload(target, target_path)

        # Same defaults as the CLI (format_doc.py's argparse defaults).
        result = format_doc.run(
            reference_path=str(reference_path),
            input_path=str(target_path),
            output_path=str(output_path),
            workdir=str(workdir),
            body_baseline_pt=11.0,
            confidence_threshold=0.7,
            fail_on_content_drift=True,
        )
    except SystemExit as e:
        # format_doc.run() calls sys.exit(2) on detected content drift —
        # surface that as a proper HTTP error instead of killing the server.
        report_path = workdir / "formatting_change_report.md"
        detail = "Formatting halted: content-integrity check failed (possible content drift)."
        if report_path.exists():
            detail += " See attached report for details."
        raise HTTPException(status_code=422, detail=detail) from e
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Formatting failed: {e}") from e

    if not output_path.exists():
        raise HTTPException(status_code=500, detail="Formatting completed but no output file was produced.")

    def _cleanup(path: Path = job_dir):
        shutil.rmtree(path, ignore_errors=True)

    out_name = f"formatted_{Path(target.filename).stem}.docx"
    return FileResponse(
        path=str(output_path),
        filename=out_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        background=BackgroundTask(_cleanup),
    )


# --- Serve the simple frontend (index.html + assets) at the site root ---
if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")

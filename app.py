"""SheetForge — Excel Report Merger & Compressor.

A deploy-ready Flask application.

Endpoints
---------
GET  /                     web UI
GET  /api/health           liveness probe (for uptime monitors)
POST /api/merge            merge uploaded workbooks (modes: stack | sheets)
POST /api/compress         size-compress .xlsx/.xlsm workbooks
POST /api/zip              bundle any uploaded files into one .zip
GET  /download/<token>     download a processed result (1h TTL)

Design notes
------------
* Stateless: every job lives in the local "work" directory as
  <token>.{ext} + <token>.json metadata; no database required.
  Works across multiple gunicorn workers out of the box.
* Files are deleted automatically one hour after they are created.
* All processing errors are returned as JSON with a friendly message;
  unexpected exceptions become a generic 500 JSON response.
"""
import json
import os
import re
import secrets
import time
import traceback
import uuid
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import HTTPException
from werkzeug.utils import secure_filename

from core.bundle import make_bundle
from core.compress import compress_workbook
from core.errors import ExcelToolError
from core.merge import merge_as_sheets, merge_stacked, read_tables

from jobs import create_job, get_job
from worker import start_merge_job

BASE_DIR = Path(__file__).resolve().parent
WORK_DIR = BASE_DIR / "work"
WORK_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Configuration (environment overridable)
# ---------------------------------------------------------------------------
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "100"))
MAX_FILES = int(os.environ.get("MAX_FILES", "30"))
JOB_TTL_SECONDS = int(os.environ.get("JOB_TTL_SECONDS", "3600"))

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".csv"}
COMPRESSIBLE_EXTENSIONS = {".xlsx", ".xlsm"}
TOKEN_RE = re.compile(r"^[0-9a-f]{32}$")

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
XLS_MIME = "application/vnd.ms-excel"
CSV_MIME = "text/csv"
ZIP_MIME = "application/zip"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.config["JSON_SORT_KEYS"] = False


# ---------------------------------------------------------------------------
# Job store (filesystem-based, multi-worker safe)
# ---------------------------------------------------------------------------
def _sweep_old_jobs():
    """Delete work files older than the TTL. Cheap; called on each request."""
    cutoff = time.time() - JOB_TTL_SECONDS
    try:
        for entry in os.scandir(WORK_DIR):
            try:
                if entry.is_file() and entry.stat().st_mtime < cutoff:
                    os.unlink(entry.path)
            except OSError:
                pass
    except OSError:
        pass


def _save_job(file_path: Path, download_name: str, mime: str) -> str:
    token = secrets.token_hex(16)
    meta = {
        "file": file_path.name,
        "name": download_name,
        "mime": mime,
        "created": time.time(),
    }
    meta_path = WORK_DIR / f"{token}.json"
    tmp = meta_path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(meta), encoding="utf-8")
    os.replace(tmp, meta_path)  # atomic, worker-safe
    return token


# ---------------------------------------------------------------------------
# Upload helpers
# ---------------------------------------------------------------------------
def _collect_uploads():
    """Validate + save uploaded files to the work dir; returns [(path, name)]."""
    files = request.files.getlist("files")
    files = [f for f in files if f and f.filename]
    if not files:
        raise ExcelToolError("No files were uploaded.")
    if len(files) > MAX_FILES:
        raise ExcelToolError(f"A maximum of {MAX_FILES} files per request is allowed.")

    saved = []
    try:
        for f in files:
            original = f.filename or "file"
            ext = Path(original).suffix.lower()
            if ext not in ALLOWED_EXTENSIONS:
                raise ExcelToolError(
                    f"Unsupported file type “{original}”. Allowed: "
                    f"xlsx, xlsm, xls, csv."
                )
            tmp = WORK_DIR / f"u_{uuid.uuid4().hex}{ext}"
            f.save(tmp)
            if tmp.stat().st_size == 0:
                tmp.unlink(missing_ok=True)
                raise ExcelToolError(f"“{original}” is empty — nothing to process.")
            saved.append((tmp, original))
    except Exception:
        for tmp, _ in saved:
            tmp.unlink(missing_ok=True)
        raise
    return saved


def _mime_for(ext: str) -> str:
    return {".xlsx": XLSX_MIME, ".xlsm": XLSX_MIME, ".xls": XLS_MIME,
            ".csv": CSV_MIME, ".zip": ZIP_MIME}.get(ext, "application/octet-stream")


def _first_stem(saved) -> str:
    stem = Path(secure_filename(saved[0][1])).stem if saved else ""
    return (re.sub(r"[^A-Za-z0-9_-]+", "_", stem).strip("_") or "report")[:60]


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/health")
def health():
    _sweep_old_jobs()
    return jsonify({"ok": True, "service": "sheetforge", "time": time.time()})


@app.post("/api/merge")
def api_merge():

    _sweep_old_jobs()

    try:

        mode = request.form.get("mode", "stack")

        if mode not in ("stack", "sheets"):
            raise ExcelToolError("Invalid merge mode.")

        strategy = request.form.get(
            "strategy",
            "union"
        )

        if strategy not in (
            "union",
            "common",
            "first",
        ):
            raise ExcelToolError(
                "Invalid column strategy."
            )

        header = (
            request.form.get("header", "1")
            == "1"
        )

        add_source = (
            request.form.get("add_source", "0")
            == "1"
        )

        dedupe = (
            request.form.get("dedupe", "0")
            == "1"
        )

        include_all = (
            request.form.get("include_all", "0")
            == "1"
        )

        # ---------------------------------------------------------
        # Collect uploaded files
        # ---------------------------------------------------------

        saved = _collect_uploads()

        # ---------------------------------------------------------
        # Create background job
        # ---------------------------------------------------------

        job_id = create_job()

        # ---------------------------------------------------------
        # Start background worker
        # ---------------------------------------------------------

        options = {
            "mode": mode,
            "strategy": strategy,
            "header": header,
            "add_source": add_source,
            "dedupe": dedupe,
            "include_all": include_all,
        }

        start_merge_job(
            job_id,
            saved,
            options,
        )

        # ---------------------------------------------------------
        # IMPORTANT:
        # Return immediately.
        # Do NOT wait for the Excel merge.
        # ---------------------------------------------------------

        return jsonify({
            "ok": True,
            "job_id": job_id,
            "status": "queued",
            "message": "Merge job started.",
        })

    except ExcelToolError as exc:

        return jsonify({
            "ok": False,
            "error": str(exc),
        }), 400

    except Exception as exc:

        traceback.print_exc()

        return jsonify({
            "ok": False,
            "error": "Could not start merge job.",
        }), 500


@app.get("/api/jobs/<job_id>")
def api_job_status(job_id):

    job = get_job(job_id)

    if not job:
        return jsonify({
            "ok": False,
            "error": "Job not found.",
        }), 404

    response = {
        "ok": True,
        "id": job["id"],
        "status": job["status"],
        "progress": job["progress"],
        "message": job["message"],
        "error": job.get("error"),
    }

    # ---------------------------------------------------------
    # Completed job
    # ---------------------------------------------------------

    if job["status"] == "completed":

        result = job.get("result") or {}

        response["result"] = {
            "name": result.get(
                "name",
                "merged.xlsx"
            ),
            "stats": result.get(
                "stats",
                {}
            ),
        }

        response["download"] = (
            f"/download/{job_id}"
        )

    return jsonify(response)

@app.post("/api/compress")
def api_compress():
    _sweep_old_jobs()
    try:
        preset = request.form.get("preset", "safe")
        if preset not in ("safe", "max"):
            raise ExcelToolError("Invalid compression preset.")

        def _clamp_int(raw, lo, hi, default):
            try:
                return max(lo, min(hi, int(raw)))
            except (TypeError, ValueError):
                return default

        max_dim = _clamp_int(request.form.get("max_dim", "1600"), 256, 4096, 1600)
        quality = _clamp_int(request.form.get("quality", "72"), 30, 95, 72)

        saved = _collect_uploads()
        try:
            for tmp, original in saved:
                if Path(original).suffix.lower() not in COMPRESSIBLE_EXTENSIONS:
                    raise ExcelToolError(
                        f"“{original}” cannot be size-compressed — the compressor "
                        f"works on .xlsx/.xlsm only. Use the ZIP tool for other "
                        f"formats."
                    )

            results = []
            for tmp, original in saved:
                out = WORK_DIR / f"{uuid.uuid4().hex}.xlsx"
                stats = compress_workbook(
                    str(tmp), str(out),
                    preset=preset, max_dim=max_dim, jpeg_quality=quality,
                )
                stem = Path(secure_filename(original)).stem or "file"
                name = f"{stem}_compressed.xlsx"
                token = _save_job(out, name, XLSX_MIME)
                results.append({
                    "file": original,
                    "token": token,
                    "name": name,
                    "download": f"/download/{token}",
                    "stats": stats,
                })
            return jsonify({"ok": True, "preset": preset, "results": results})
        finally:
            for tmp, _ in saved:
                tmp.unlink(missing_ok=True)
    except ExcelToolError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/zip")
def api_zip():
    _sweep_old_jobs()
    try:
        saved = _collect_uploads()
        try:
            out = WORK_DIR / f"{uuid.uuid4().hex}.zip"
            stats = make_bundle([(str(tmp), original) for tmp, original in saved],
                                str(out))
            name = f"{_first_stem(saved)}_bundle.zip"
            token = _save_job(out, name, ZIP_MIME)
            return jsonify({"ok": True, "token": token, "name": name,
                            "download": f"/download/{token}", "stats": stats})
        finally:
            for tmp, _ in saved:
                tmp.unlink(missing_ok=True)
    except ExcelToolError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.get("/download/<job_id>")
def download(job_id):

    job = get_job(job_id)

    if not job:
        return jsonify({
            "ok": False,
            "error": "Download not found.",
        }), 404

    if job["status"] != "completed":
        return jsonify({
            "ok": False,
            "error": "The merge is not completed yet.",
        }), 400

    result = job.get("result") or {}

    file_path = result.get("file")

    if not file_path:
        return jsonify({
            "ok": False,
            "error": "Result file is missing.",
        }), 404

    path = Path(file_path)

    if not path.is_file():
        return jsonify({
            "ok": False,
            "error": "Result file no longer exists.",
        }), 404

    return send_file(
        path,
        as_attachment=True,
        download_name=result.get(
            "name",
            "merged.xlsx"
        ),
        mimetype=XLSX_MIME,
    )

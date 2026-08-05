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
from __future__ import annotations

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
        strategy = request.form.get("strategy", "union")
        if strategy not in ("union", "common", "first"):
            raise ExcelToolError("Invalid column strategy.")
        header = request.form.get("header", "1") == "1"
        add_source = request.form.get("add_source", "0") == "1"
        dedupe = request.form.get("dedupe", "0") == "1"
        include_all = request.form.get("include_all", "0") == "1"

        saved = _collect_uploads()
        try:
            tables = []
            for tmp, original in saved:
                tables.extend(read_tables(tmp, original, include_all))
        finally:
            for tmp, _ in saved:
                tmp.unlink(missing_ok=True)

        if mode == "stack":
            wb, stats = merge_stacked(
                tables,
                header=header,
                strategy=strategy,
                add_source=add_source,
                dedupe=dedupe,
                include_all_sheets=include_all,
            )
        else:
            wb, stats = merge_as_sheets(tables, include_all_sheets=include_all)

        out = WORK_DIR / f"{uuid.uuid4().hex}.xlsx"
        wb.save(out)
        name = f"{_first_stem(saved)}_merged.xlsx"
        token = _save_job(out, name, XLSX_MIME)
        return jsonify({"ok": True, "token": token, "name": name,
                        "download": f"/download/{token}", "stats": stats})
    except ExcelToolError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


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


@app.get("/download/<token>")
def download(token: str):
    _sweep_old_jobs()
    if not TOKEN_RE.match(token):
        return jsonify({"ok": False, "error": "Invalid download token."}), 404
    meta_path = WORK_DIR / f"{token}.json"
    if not meta_path.is_file():
        return jsonify({"ok": False,
                        "error": "This download has expired or does not exist "
                                 "(files are deleted 1 hour after processing)."}), 404
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        file_path = WORK_DIR / meta["file"]
        if not file_path.is_file():
            return jsonify({"ok": False, "error": "Result file is missing."}), 404
        return send_file(file_path, as_attachment=True,
                         download_name=meta["name"], mimetype=meta["mime"])
    except Exception:
        traceback.print_exc()
        return jsonify({"ok": False, "error": "Could not serve the file."}), 500


# ---------------------------------------------------------------------------
# Error handlers
# ---------------------------------------------------------------------------
@app.errorhandler(413)
def too_large(_e):
    return jsonify({
        "ok": False,
        "error": f"Upload too large — the limit is {MAX_UPLOAD_MB} MB per request.",
    }), 413


@app.errorhandler(Exception)
def unexpected_error(e):
    if isinstance(e, HTTPException):
        return jsonify({"ok": False, "error": e.description}), e.code
    traceback.print_exc()
    return jsonify({"ok": False,
                    "error": "Unexpected server error. Please try again."}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, threaded=True, debug=False)

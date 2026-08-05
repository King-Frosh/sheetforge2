# SheetForge — Excel Report Merger & Compressor

A full-stack, deploy-ready web app that lets your users:

| Tool | What it does |
|---|---|
| **Merge** | Combine multiple Excel/CSV reports into **one workbook** — either *stack all rows into a single sheet* (columns auto-aligned by header) or *keep each file as its own sheet*. |
| **Compress** | True size compression for `.xlsx`/`.xlsm`: re-encodes embedded images, strips hidden junk, repacks the ZIP at maximum compression. A "Maximum" preset also removes conditional formatting, data validations and empty sheets. |
| **ZIP bundle** | Packs any number of files into one `.zip` for email/archiving. |

Stack = Flask · openpyxl · Pillow. No database. No accounts. Files are
deleted automatically 1 hour after processing.

---

## Quick start (local)

```bash
cd sheetforge
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

python app.py            # → http://localhost:5000
```

Run the test suite (8 groups, end-to-end incl. HTTP API):

```bash
pip install -r requirements-dev.txt
python tests/test_flows.py
```

---

## Deploying to the web

### Option A — Render.com (easiest, free tier works)

1. Push this folder to a GitHub repo.
2. In Render: **New → Web Service → connect the repo**.
3. Render auto-detects Python. The included [`render.yaml`](render.yaml) also
   supports **Blueprint** deploys (New → Blueprint) with zero config.
4. Start command (set automatically by `render.yaml`):

   ```bash
   gunicorn app:app --bind 0.0.0.0:$PORT --timeout 300 --workers 2 --threads 4
   ```

5. Health check: `https://your-app.onrender.com/api/health`

### Option B — Railway

1. Push to GitHub → **New Project → Deploy from GitHub repo**.
2. Railway auto-detects the Python app (Nixpacks). Set the start command:

   ```bash
   gunicorn app:app --bind 0.0.0.0:$PORT --timeout 300 --workers 2 --threads 4
   ```

### Option C — PythonAnywhere

1. Upload the code, create a virtualenv and `pip install -r requirements.txt`.
2. In **Web → WSGI configuration file**, point to this app:

   ```python
   from app import app as application
   ```

3. Free accounts need the code in `/home/<user>/mysite` with the static
   folder mapped: URL `/static` → `/home/<user>/mysite/static`.

### Option D — Any VPS / Docker

```bash
docker build -t sheetforge .
docker run -d -p 8000:8000 --name sheetforge sheetforge
```

---

## Configuration (environment variables)

| Variable | Default | Meaning |
|---|---|---|
| `MAX_UPLOAD_MB` | `100` | Max request size in MB (Flask `MAX_CONTENT_LENGTH`) |
| `MAX_FILES` | `30` | Max files per request |
| `JOB_TTL_SECONDS` | `3600` | How long processed files stay downloadable |
| `PORT` | `5000` | Local dev port (`$PORT` is used automatically on Render/Railway) |

---

## How the pieces work

### Merging (`core/merge.py`)
* Reads `.xlsx`, `.xlsm` (openpyxl), `.xls` (xlrd) and `.csv` (encoding +
  delimiter detection).
* **Stack mode**: header rows are matched by normalised name; column order
  = first-seen (union), shared-only (common), or first-file-only. Optionally
  adds a "Source File" column, removes exact duplicates (type-tolerant:
  `300` == `"300"`), and includes all sheets (tagged `file [sheet]`).
* Output gets a bold header, frozen first row, autofilter and auto-sized
  columns (small merges) or a memory-safe streaming writer (>25k rows).
* Bounded: 500k rows / 500 columns max per merge — errors are friendly.

### Compressing (`core/compress.py`)
An `.xlsx` is a ZIP package, so instead of re-exporting (which would destroy
images/charts/formatting) the compressor edits the package in place:
1. Images in `xl/media/` are re-encoded (JPEG quality, PNG optimize, max
   dimension) — only replaced when actually smaller.
2. Always-removed junk: `calcChain.xml`, `docProps/thumbnail.*`.
3. **Max preset**: strips `conditionalFormatting` / `dataValidations` from
   sheet XML and deletes genuinely empty sheets — with full relationship,
   content-type and part cleanup so Excel never shows a repair prompt.
4. Everything is repacked with DEFLATE level 9. If repacking ever gains
   nothing, the original file is handed back byte-for-byte.

### Architecture (`app.py`)
Stateless: each job = `work/<token>.xlsx` + `work/<token>.json`. No shared
memory, so it works across multiple gunicorn workers. Old files are swept
on every request. All user-facing errors return JSON with status 4xx; the
UI renders them as friendly messages.

---

## Project layout

```
sheetforge/
├── app.py                 # Flask app: routes, uploads, jobs, errors
├── core/
│   ├── merge.py           # stack + per-sheet merging
│   ├── compress.py        # in-place ZIP surgery + image re-encoding
│   ├── bundle.py          # ZIP bundling with duplicate-name handling
│   └── errors.py          # user-facing exceptions
├── templates/index.html   # single-page UI (no external CDNs/assets)
├── static/css|js/         # styles + frontend logic
├── tests/test_flows.py    # 8 end-to-end test groups
├── requirements.txt
├── Dockerfile
└── render.yaml            # one-click Render Blueprint config
```

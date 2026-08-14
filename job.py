import json
import os
import threading
import uuid
from pathlib import Path

WORK_DIR = Path(os.environ.get("WORK_DIR", "/app/work"))
WORK_DIR.mkdir(parents=True, exist_ok=True)

_jobs = {}
_lock = threading.Lock()


def create_job():
    job_id = uuid.uuid4().hex

    with _lock:
        _jobs[job_id] = {
            "id": job_id,
            "status": "queued",
            "progress": 0,
            "message": "Waiting to start",
            "result": None,
            "error": None,
        }

    return job_id


def update_job(job_id, **kwargs):
    with _lock:
        if job_id in _jobs:
            _jobs[job_id].update(kwargs)


def get_job(job_id):
    with _lock:
        return dict(_jobs.get(job_id, {}))

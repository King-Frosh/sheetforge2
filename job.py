import threading

_jobs = {}
_lock = threading.Lock()


def create_job():
    """
    Create a new background job and return its ID.
    """
    import uuid

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
    """
    Update the status/details of an existing job.
    """
    with _lock:
        if job_id in _jobs:
            _jobs[job_id].update(kwargs)


def get_job(job_id):
    """
    Retrieve a copy of a job.
    """
    with _lock:
        job = _jobs.get(job_id)

        if job is None:
            return None

        return dict(job)

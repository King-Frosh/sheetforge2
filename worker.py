import threading
import traceback
from pathlib import Path

from jobs import update_job
from core.errors import ExcelToolError
from core.merge import read_tables, merge_stacked, merge_as_sheets

WORK_DIR = Path("/app/work")
WORK_DIR.mkdir(parents=True, exist_ok=True)


def start_merge_job(job_id, uploaded_files, options):
    """
    Start the Excel merge in a background thread.
    """
    thread = threading.Thread(
        target=run_merge_job,
        args=(job_id, uploaded_files, options),
        daemon=True,
    )
    thread.start()


def _cleanup_input_files(uploaded_files):
    """Remove temporary uploaded files after they have been read."""
    for path, _original_name in uploaded_files:
        try:
            Path(path).unlink(missing_ok=True)
        except OSError:
            pass


def run_merge_job(job_id, uploaded_files, options):
    """Perform the actual merge operation in the background."""
    try:
        update_job(
            job_id,
            status="processing",
            progress=5,
            message="Preparing files...",
        )

        tables = []
        total_files = len(uploaded_files)

        # ---------------------------------------------------------
        # Read uploaded spreadsheets
        # ---------------------------------------------------------
        for index, (path, original_name) in enumerate(uploaded_files):
            update_job(
                job_id,
                progress=10 + int(((index) / max(total_files, 1)) * 40),
                message=f"Reading {original_name}...",
            )

            try:
                tables.extend(
                    read_tables(
                        str(path),
                        original_name,
                        options.get("include_all", False),
                    )
                )
            finally:
                # The source workbook is no longer needed once it has
                # been converted into the in-memory table representation.
                try:
                    Path(path).unlink(missing_ok=True)
                except OSError:
                    pass

        update_job(
            job_id,
            progress=50,
            message="Merging spreadsheets...",
        )

        # ---------------------------------------------------------
        # Merge spreadsheets
        # ---------------------------------------------------------
        mode = options.get("mode", "stack")

        if mode == "sheets":
            workbook, stats = merge_as_sheets(
                tables,
                include_all_sheets=options.get("include_all", False),
            )
        else:
            workbook, stats = merge_stacked(
                tables,
                header=options.get("header", True),
                strategy=options.get("strategy", "union"),
                add_source=options.get("add_source", False),
                dedupe=options.get("dedupe", False),
                include_all_sheets=options.get("include_all", False),
            )

        update_job(
            job_id,
            progress=85,
            message="Creating final Excel file...",
        )

        # ---------------------------------------------------------
        # Save result in the root work directory.
        # This allows the existing cleanup sweep in app.py to remove
        # old result files automatically after JOB_TTL_SECONDS.
        # ---------------------------------------------------------
        output_file = WORK_DIR / f"{job_id}.xlsx"
        workbook.save(output_file)

        update_job(
            job_id,
            status="completed",
            progress=100,
            message="Merge completed successfully.",
            result={
                "file": str(output_file),
                "name": "merged.xlsx",
                "stats": stats,
            },
        )

    except ExcelToolError as exc:
        _cleanup_input_files(uploaded_files)

        update_job(
            job_id,
            status="failed",
            progress=0,
            message="Merge failed.",
            error=str(exc),
        )

    except Exception as exc:
        _cleanup_input_files(uploaded_files)
        traceback.print_exc()

        update_job(
            job_id,
            status="failed",
            progress=0,
            message="Unexpected server error.",
            error=str(exc),
        )

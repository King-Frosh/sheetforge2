import threading
import traceback
from pathlib import Path

from jobs import update_job
from core.errors import ExcelToolError
from core.merge import read_tables, merge_stacked, merge_as_sheets


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


def run_merge_job(job_id, uploaded_files, options):
    """
    Perform the actual merge operation.
    """

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
        for index, item in enumerate(uploaded_files):

            path, original_name = item

            update_job(
                job_id,
                progress=10 + int(
                    ((index) / max(total_files, 1)) * 40
                ),
                message=f"Reading {original_name}...",
            )

            tables.extend(
                read_tables(
                    str(path),
                    original_name,
                    options.get("include_all", False),
                )
            )

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
                include_all_sheets=options.get(
                    "include_all",
                    False,
                ),
            )

        else:

            workbook, stats = merge_stacked(
                tables,
                header=options.get("header", True),
                strategy=options.get("strategy", "union"),
                add_source=options.get("add_source", False),
                dedupe=options.get("dedupe", False),
                include_all_sheets=options.get(
                    "include_all",
                    False,
                ),
            )

        update_job(
            job_id,
            progress=85,
            message="Creating final Excel file...",
        )

        # ---------------------------------------------------------
        # Save result
        # ---------------------------------------------------------
        job_dir = Path("/app/work") / job_id
        job_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        output_file = job_dir / "merged.xlsx"

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

        update_job(
            job_id,
            status="failed",
            progress=0,
            message="Merge failed.",
            error=str(exc),
        )

    except Exception as exc:

        traceback.print_exc()

        update_job(
            job_id,
            status="failed",
            progress=0,
            message="Unexpected server error.",
            error=str(exc),
        )

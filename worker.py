import threading
from pathlib import Path

from jobs import update_job
from core.merge import read_tables, merge_stacked


def start_merge_job(job_id, uploaded_files, options):

    thread = threading.Thread(
        target=run_merge,
        args=(job_id, uploaded_files, options),
        daemon=True,
    )

    thread.start()


def run_merge(job_id, uploaded_files, options):

    try:
        update_job(
            job_id,
            status="processing",
            progress=10,
            message="Reading spreadsheets...",
        )

        tables = []

        total = len(uploaded_files)

        for index, (path, filename) in enumerate(uploaded_files):

            tables.extend(
                read_tables(
                    str(path),
                    filename,
                    options["include_all"],
                )
            )

            progress = 10 + int((index + 1) / total * 40)

            update_job(
                job_id,
                progress=progress,
                message=f"Reading {filename}",
            )

        update_job(
            job_id,
            progress=60,
            message="Merging files...",
        )

        workbook, stats = merge_stacked(
            tables,
            header=options["header"],
            strategy=options["strategy"],
            add_source=options["add_source"],
            dedupe=options["dedupe"],
            include_all_sheets=options["include_all"],
        )

        update_job(
            job_id,
            progress=85,
            message="Creating output workbook...",
        )

        output = Path("/app/work") / f"{job_id}.xlsx"

        workbook.save(output)

        update_job(
            job_id,
            status="completed",
            progress=100,
            message="Merge completed",
            result={
                "file": output.name,
                "stats": stats,
            },
        )

    except Exception as exc:

        update_job(
            job_id,
            status="failed",
            error=str(exc),
            message="Merge failed",
        )

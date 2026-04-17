"""
function_app.py — Azure Function (Flex Consumption, Python v2)
--------------------------------------------------------------
Timer-triggered function that:
  1. Downloads 3 regional Excel workbooks from SharePoint via Graph API
  2. Generates ~84 text files (health summaries, parameter trends, indexes)
  3. Uploads them to the Foundry vector store, replacing stale files

Schedule: Every Monday 6:00 AM UTC  (configurable via TIMER_SCHEDULE app setting)
"""

import datetime
import logging
import os
import tempfile

import azure.functions as func

from sharepoint_client import SharePointClient
from foundry_client import FoundryVectorStoreClient
from report_generator import generate_all_reports

app = func.FunctionApp()

# ── Timer Trigger ────────────────────────────────────────────────────────────

@app.function_name(name="SyncSprintData")
@app.timer_trigger(
    schedule="%TIMER_SCHEDULE%",      # App Setting, e.g. "0 0 6 * * 1"
    arg_name="timer",
    run_on_startup=False,
)
def sync_sprint_data(timer: func.TimerRequest) -> None:
    """Main entry point – runs on schedule."""
    utc_now = datetime.datetime.now(datetime.timezone.utc).isoformat()
    logging.info("SyncSprintData triggered at %s", utc_now)

    if timer.past_due:
        logging.warning("Timer is past due – executing anyway.")

    try:
        _run_pipeline()
        logging.info("Pipeline completed successfully.")
    except Exception:
        logging.exception("Pipeline failed.")
        raise  # re-raise so Azure marks the invocation as failed


# ── HTTP Trigger (manual / testing) ──────────────────────────────────────────

@app.function_name(name="SyncSprintDataHttp")
@app.route(route="sync", auth_level=func.AuthLevel.FUNCTION)
def sync_sprint_data_http(req: func.HttpRequest) -> func.HttpResponse:
    """Manual trigger for testing – POST /api/sync"""
    logging.info("Manual HTTP trigger invoked.")
    try:
        _run_pipeline()
        return func.HttpResponse("Pipeline completed successfully.", status_code=200)
    except Exception as exc:
        logging.exception("Pipeline failed on manual trigger.")
        return func.HttpResponse(f"Pipeline failed: {exc}", status_code=500)


# ── Pipeline logic ───────────────────────────────────────────────────────────

# SharePoint file names to download (must match what's in the document library)
SP_FILES = [
    "FY26 Americas Sprint Checkpoint Tracker v2.0.xlsm",
    "FY26 EMEA Sprint Checkpoint Tracker v2.0.xlsm",
    "FY26 Asia Sprint Checkpoint Tracker v2.0.xlsm",
]


def _run_pipeline():
    """Download → Generate → Upload."""
    work_dir = tempfile.mkdtemp(prefix="sprint_sync_")
    download_dir = os.path.join(work_dir, "downloads")
    output_dir = os.path.join(work_dir, "output")
    os.makedirs(download_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    # ── Step 1: Download Excel files from SharePoint ─────────────────────────
    logging.info("Step 1/3: Downloading Excel files from SharePoint...")
    sp = SharePointClient()
    downloaded_paths = []
    for filename in SP_FILES:
        local_path = os.path.join(download_dir, filename)
        sp.download_file(filename, local_path)
        downloaded_paths.append(local_path)
        logging.info("  Downloaded: %s", filename)

    # ── Step 2: Generate text reports from Excel ─────────────────────────────
    logging.info("Step 2/3: Generating text reports...")
    generated_files = generate_all_reports(downloaded_paths, output_dir)
    logging.info("  Generated %d files.", len(generated_files))

    # ── Step 3: Upload to Foundry vector store ───────────────────────────────
    logging.info("Step 3/3: Syncing to Foundry vector store...")
    foundry = FoundryVectorStoreClient()
    foundry.sync_files(output_dir)
    logging.info("  Vector store sync complete.")

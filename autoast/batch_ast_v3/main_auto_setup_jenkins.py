# autoast is a script for batch processing the automated status tool
# author: csostad
# copyright Government of British Columbia
# Copyright 2019 Province of British Columbia

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at

#    http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""
Jenkins headless entry point for the AutoAST Batch Tool.

Workflow:
  1. Reads WATCH_DIR from .env — the root folder on the objectstore where users
     deposit shapefile submissions (e.g., \\objectstore2.nrs.bcgov\\GSS_Share\\
     authorizations\\batch_ast_tool).
  2. Scans WATCH_DIR for immediate subdirectories that contain shapefiles (or
     subfolders-of-shapefiles). Each subdirectory is treated as one user submission.
  3. Skips any submission whose outputs/BATCH_COMPLETE.json sentinel already exists
     (written at the end of a previous successful run). To force a rerun, the user
     deletes that file and waits for the next nightly Jenkins build.
  4. For each pending submission:
       a. Creates an outputs/ directory inside the submission folder.
       b. Collects shapefiles using the shared _collect_shapefiles() helper from
          multi_excel_setup.py — no logic is duplicated.
       c. Auto-detects NR Regions for every shapefile via a BCGW spatial join
          (detect_regions() from region_detector.py). No manual region selection.
       d. Writes Excel job files (same format as the interactive tool, max 8 jobs
          each) into outputs/.
       e. Runs the batch tool for each Excel file (same process_excel_file loop
          used by main_auto_setup.py and main.py).
       f. Counts pass/fail row outcomes across all Excel files.
       g. Writes outputs/BATCH_COMPLETE.json with a structured summary sentinel
          regardless of individual job success or failure.
  5. Writes a single orchestration log in WATCH_DIR/autoast_batch_logs_folder_YYYYMMDD/
     (visible to all users on the objectstore drive).

Required .env keys (in addition to the standard set):
  WATCH_DIR — root folder path to scan for user submissions.

Intended to be run nightly by Jenkins. No tkinter/GUI dependencies.

Entry-point sync rule (from copilot instructions):
  Any changes to shared orchestration logic (process_excel_file, AST_FACTORY
  instantiation, the __main__ startup sequence, per-spreadsheet banners, error
  guards, etc.) must also be applied to main.py and main_auto_setup.py.
"""

import os
import json
import datetime
from pathlib import Path
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook

# Standard batch-tool imports shared with main.py and main_auto_setup.py
from logging_setup import setup_logging
from database_connection import setup_bcgw
from toolbox_import import import_ast
from ast_factory import AST_FACTORY
from failed_job_tracker import FailedJobTracker

# Headless shapefile collection helper — imported rather than duplicated so
# any future change to discovery logic applies here automatically.
from multi_excel_setup import _collect_shapefiles

# Region detection via BCGW spatial join — replaces the manual GUI dropdown.
from region_detector import detect_regions


# ---------------------------------------------------------------------------
# Module-level constants — keep in sync with multi_excel_setup.py and
# ast_factory.py (AST_PARAMETERS / ADDITIONAL_PARAMETERS dicts).
# ---------------------------------------------------------------------------

# Sheet name inside every job Excel workbook — must match AST_FACTORY.XLSX_SHEET_NAME.
_XLSX_SHEET_NAME = "ast_config"

# Column headers in the exact order expected by AST_FACTORY.load_jobs().
# Must stay in sync with multi_excel_setup.create_job_excel_files() headers list.
_XLSX_HEADERS = [
    "region", "feature_layer", "crown_file_number", "disposition_number",
    "parcel_number", "output_directory", "output_dir_same_as_input",
    "dont_overwrite_outputs", "dont_run_conflicts_and_constraints_tab3",
    "suppress_map_creation_tab3", "open_output_directory", "full_path_hyperlinks",
    "ast_condition", "file_number",
]

# Maximum job rows per workbook — enforced by BCGW concurrent connection limit.
# Matches the hard limit in multi_excel_setup.py (8 jobs per file).
_MAX_JOBS_PER_WORKBOOK = 8

# Sentinel file name written inside outputs/ after a submission has been processed.
# Jenkins checks for this file to decide whether to skip the submission on the next run.
_SENTINEL_FILENAME = "BATCH_COMPLETE.json"


# ---------------------------------------------------------------------------
# process_excel_file
# Copied verbatim from main_auto_setup.py.  All three entry points (main.py,
# main_auto_setup.py, this file) carry their own copy per the project rule that
# all changes to shared orchestration logic must be applied to all three files.
# ---------------------------------------------------------------------------

def process_excel_file(excel_file, secrets, logger, current_path, failed_job_tracker, delete_transitory=False):
    '''
    This function takes a list of excel files and iterates over that list, applying the Batch AST Class (and hence the ast tool)
    too each row in each excel file. This is a workaround for multiprocessing issue with the BCGW sees too many db connections
    in batches of 8.
    failed_job_tracker: FailedJobTracker instance passed through to AST_FACTORY so all failures are recorded in the summary spreadsheet.
    delete_transitory: if True, transitory GDB/log files are deleted immediately after each job is verified COMPLETE (not in a
    bulk sweep at the end). Failed/FAILED_OUTPUTS jobs are never cleaned so their GDBs remain available for restart analysis.
    '''
    try:
        print(f"Main: Creating queuefile path for {excel_file}")
        logger.info(f"Main: Creating queuefile path for {excel_file}")

        qf = os.path.join(current_path, excel_file)

        # Create an instance of the AST_FACTORY class
        # Pass failed_job_tracker so AST_FACTORY logs failures to the summary spreadsheet
        # Pass delete_transitory so per-job cleanup is triggered only after verified COMPLETE
        ast = AST_FACTORY(qf, secrets[0], secrets[1], logger, current_path, failed_job_tracker, delete_transitory)

        if not os.path.exists(qf):
            print(f"Main: Queuefile for {excel_file} not found, creating new queuefile")
            logger.info(f"Main: Queuefile for {excel_file} not found, creating new queuefile")
            ast.create_new_queuefile()

        # Load jobs from the Excel file
        print(f"Main: Loading jobs from {excel_file}")
        logger.info(f"Main: Loading jobs from {excel_file}")
        jobs = ast.load_jobs()

        # Batch jobs
        print(f"Main: Batching jobs for {excel_file}")
        logger.info(f"Main: Batching jobs for {excel_file}")
        ast.batch_ast()

        # Reload failed jobs
        print(f"Main: Reloading failed jobs for {excel_file}")
        logger.info(f"Main: Reloading failed jobs for {excel_file}")
        ast.re_load_failed_jobs_V2()

        # Re-batch failed jobs
        print(f"Main: Re-batching failed jobs for {excel_file}")
        logger.info(f"Main: Re-batching failed jobs for {excel_file}")
        ast.batch_ast()

        print(f"\n{'='*100}")
        print(f"  >>> COMPLETED SPREADSHEET: {excel_file} <<<")
        print(f"{'='*100}\n")
        logger.info(f"\n{'='*100}")
        logger.info(f"  >>> COMPLETED SPREADSHEET: {excel_file} <<<")
        logger.info(f"{'='*100}\n")


    except Exception as e:
        print(f"Error processing {excel_file}: {e}")
        logger.error(f"Error processing {excel_file}: {e}")


# ---------------------------------------------------------------------------
# Sentinel helpers
# ---------------------------------------------------------------------------

def _is_submission_complete(submission_dir: Path) -> bool:
    """
    Returns True if this submission was already processed by a previous Jenkins run.

    The presence of outputs/BATCH_COMPLETE.json is the authoritative signal — it is
    written at the end of every run (even if some jobs failed) so Jenkins knows not
    to retry.  To force a rerun, the user deletes the sentinel file, fixes any broken
    shapefiles, and waits for the next nightly build.
    """
    # Check for the sentinel file at the expected location inside the submission folder.
    sentinel_path = submission_dir / "outputs" / _SENTINEL_FILENAME
    return sentinel_path.exists()


def _count_job_statuses(excel_files: list) -> tuple:
    """
    Scans every Excel job file produced for a submission and counts row-level outcomes.

    Reads the 'ast_condition' column from the 'ast_config' sheet of each workbook.
    The column is located by header name (not by index) so the count remains correct
    even if columns are ever reordered in a future schema change.  The header row
    (row 1) is always skipped.

    Returns:
        (total, complete, failed) — int counts of:
          total    : all data rows found across all Excel files
          complete : rows whose ast_condition value is exactly "COMPLETE"
          failed   : rows with any other ast_condition value (Failed, FAILED_OUTPUTS,
                     Queued, blank) — i.e., anything that is not a clean success
    """
    total = 0
    complete = 0
    failed = 0

    for excel_path in excel_files:
        try:
            # Open read-only for speed; data_only=True reads cached cell values rather
            # than formulas — safe here because the batch tool writes plain strings.
            wb = load_workbook(excel_path, read_only=True, data_only=True)

            if _XLSX_SHEET_NAME not in wb.sheetnames:
                # Expected sheet is absent — treat the whole file as unresolved.
                wb.close()
                continue

            ws = wb[_XLSX_SHEET_NAME]
            rows = list(ws.iter_rows(values_only=True))
            wb.close()

            if not rows:
                # Empty sheet — nothing to count.
                continue

            # Locate ast_condition by scanning the header row (index 0).
            # Using index() raises ValueError if the header is absent, caught below.
            header_row = rows[0]
            try:
                condition_col_idx = list(header_row).index("ast_condition")
            except ValueError:
                # Header not found — cannot determine status for this file; skip it.
                continue

            # Count every data row (row index 1 onward).
            for row in rows[1:]:
                if row is None:
                    continue
                total += 1
                condition_value = row[condition_col_idx]
                if condition_value == "COMPLETE":
                    complete += 1
                else:
                    # Anything other than COMPLETE is treated as failed/incomplete.
                    failed += 1

        except Exception as e:
            # A read failure on one file must not abort the count — log it and continue
            # so the sentinel is still written with whatever partial data we have.
            print(f"WARNING: Could not read status counts from {excel_path}: {e}")

    return total, complete, failed


def _write_completion_sentinel(outputs_dir: Path, submission_dir: Path,
                               total: int, complete: int, failed: int,
                               excel_files: list, logger) -> None:
    """
    Writes BATCH_COMPLETE.json into outputs_dir after a submission has been processed.

    The sentinel is ALWAYS written regardless of individual job pass/fail status.
    Its presence tells Jenkins to skip this submission on all subsequent nightly runs.
    To force a rerun: delete this file, fix any failing shapefiles, and wait for the
    next nightly Jenkins build.

    JSON payload fields:
      status              — "ALL_COMPLETE" if every job row is COMPLETE; else "PARTIAL_FAILURE"
      submission_folder   — name of the user's submission directory (for dashboard display)
      timestamp           — ISO-8601 wall-clock time at time of writing
      jobs_total          — total job rows found across all Excel files
      jobs_complete       — rows with ast_condition == "COMPLETE"
      jobs_failed         — rows with any other ast_condition value
      excel_files         — list of Excel file paths for this submission (relative to outputs_dir
                            when possible so the JSON stays valid if the drive is remounted)
      rerun_instructions  — human-readable string explaining how to trigger a rerun
    """
    # Determine overall outcome.  ALL_COMPLETE requires at least one job and zero failures.
    status = "ALL_COMPLETE" if (total > 0 and failed == 0) else "PARTIAL_FAILURE"

    # Build portable Excel file references — relative to outputs_dir where possible so
    # the JSON remains valid if the objectstore is remounted at a different drive letter.
    portable_excel_paths = []
    for ef in excel_files:
        ef_path = Path(ef)
        try:
            # is_relative_to() requires Python 3.9+ — available in ArcGIS Pro 3 / Python 3.9+.
            portable_excel_paths.append(
                str(ef_path.relative_to(outputs_dir)) if ef_path.is_relative_to(outputs_dir) else str(ef)
            )
        except ValueError:
            # Fallback: store absolute path if relative_to raises.
            portable_excel_paths.append(str(ef))

    # Structured payload — all fields are simple JSON-safe types so any consumer
    # (Jenkins post-build script, dashboard, Power BI) can parse without knowledge
    # of internal batch-tool data structures.  The 'status' field alone is sufficient
    # for a dashboard to colour-code a row green/red.
    payload = {
        "status": status,
        "submission_folder": submission_dir.name,
        "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
        "jobs_total": total,
        "jobs_complete": complete,
        "jobs_failed": failed,
        "excel_files": portable_excel_paths,
        "rerun_instructions": (
            "To force a rerun: delete this file, fix any failed shapefiles, "
            "and wait for the next nightly Jenkins run."
        ),
    }

    sentinel_path = outputs_dir / _SENTINEL_FILENAME
    try:
        # indent=2 keeps the file human-readable in a text editor and in browser previews.
        with open(sentinel_path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2)
        print(f"Sentinel written: {sentinel_path}  [status={status}]")
        logger.info(f"Sentinel written: {sentinel_path}  [status={status}]")
    except Exception as e:
        # If the sentinel cannot be written, Jenkins will retry the submission on the
        # next nightly run — which is the safe failure mode (retry rather than silently
        # skip).  Re-raise so the caller's exception handler records this in the log.
        print(f"ERROR: Could not write completion sentinel {sentinel_path}: {e}")
        logger.error(f"ERROR: Could not write completion sentinel {sentinel_path}: {e}")
        raise


# ---------------------------------------------------------------------------
# Headless Excel workbook writer
# ---------------------------------------------------------------------------

def _write_excel_files(jobs: list, region_map: dict, output_dir: Path) -> list:
    """
    Writes Excel job workbooks to output_dir from collected shapefile jobs and
    the region_map returned by detect_regions().

    This is the headless equivalent of the workbook-writing section of
    multi_excel_setup.create_job_excel_files().  It uses the same column headers,
    the same 8-job split logic, and the same jobs_N.xlsx naming convention so
    AST_FACTORY can consume the files without any modifications.

    Parameters:
        jobs       : list of (Path, Path) — (shapefile path, output subfolder path)
        region_map : dict {int index -> str region name} from detect_regions()
        output_dir : Path — the outputs/ directory to write workbooks into

    Returns:
        list[str] — absolute string paths to every Excel file created; same
        contract as create_job_excel_files() so the caller loop is unchanged.
    """
    excel_files = []        # Accumulated list of file paths to return to the caller
    workbook_counter = 1    # Used to produce names: jobs_1.xlsx, jobs_2.xlsx, …
    job_counter = 0         # Tracks how many rows are in the current open workbook

    # Start the first workbook.
    wb = Workbook()
    ws = wb.active
    ws.title = _XLSX_SHEET_NAME
    ws.append(_XLSX_HEADERS)   # Row 1: column headers

    for i, (shp_path, output_subfolder) in enumerate(jobs):
        # Use the detected region; fall back to empty string if detection failed
        # for this shapefile (warning already logged by process_submission).
        region = region_map.get(i, "")

        # Append one data row — defaults match the interactive tool defaults so
        # the AST tool behaves identically whether run from the GUI or Jenkins.
        ws.append([
            region,
            str(shp_path),
            "",       # crown_file_number — blank; user fills if needed
            "",       # disposition_number
            "",       # parcel_number
            str(output_subfolder),
            "false",  # output_dir_same_as_input
            "false",  # dont_overwrite_outputs
            "false",  # dont_run_conflicts_and_constraints_tab3
            "false",  # suppress_map_creation_tab3
            "false",  # open_output_directory — must be False in batch (no desktop to open)
            "true",   # full_path_hyperlinks
            "",       # ast_condition — blank; AST_FACTORY writes this at runtime
            "",       # file_number
        ])
        job_counter += 1
        print(f"  Job {i + 1}: {shp_path.name} -> region='{region}'")

        # Split into a new workbook every _MAX_JOBS_PER_WORKBOOK rows.
        # Each workbook maps to one BCGW connection slot; exceeding 8 causes errors.
        if job_counter == _MAX_JOBS_PER_WORKBOOK:
            batch_excel_path = output_dir / f"jobs_{workbook_counter}.xlsx"
            wb.save(str(batch_excel_path))
            excel_files.append(str(batch_excel_path))
            print(f"  Saved: {batch_excel_path}")
            # Reset counters and open a new workbook for the next batch.
            workbook_counter += 1
            job_counter = 0
            wb = Workbook()
            ws = wb.active
            ws.title = _XLSX_SHEET_NAME
            ws.append(_XLSX_HEADERS)

    # Save any remaining rows that did not fill a complete workbook of _MAX_JOBS_PER_WORKBOOK.
    if job_counter > 0:
        batch_excel_path = output_dir / f"jobs_{workbook_counter}.xlsx"
        wb.save(str(batch_excel_path))
        excel_files.append(str(batch_excel_path))
        print(f"  Saved: {batch_excel_path}")

    return excel_files


# ---------------------------------------------------------------------------
# Per-submission orchestrator
# ---------------------------------------------------------------------------

def process_submission(submission_dir: Path, secrets: list, sde_path: str,
                       logger, current_path: str) -> None:
    """
    Orchestrates one complete submission end-to-end:
      1. Creates outputs/ directory inside the submission folder.
      2. Collects shapefiles via the shared _collect_shapefiles() helper.
      3. Detects NR Regions via BCGW spatial join (detect_regions()).
      4. Writes Excel job files into outputs/.
      5. Runs the batch tool (same process_excel_file loop as main_auto_setup.py).
      6. Counts pass/fail outcomes and writes the BATCH_COMPLETE.json sentinel
         unconditionally — Jenkins always stops retrying after one full attempt.

    Parameters:
        submission_dir : Path   — immediate subdirectory of WATCH_DIR containing the user's shapefiles
        secrets        : list   — [BCGW_USER, BCGW_PASS] from setup_bcgw()
        sde_path       : str    — path to bcgw.sde; passed to detect_regions() and worker processes
        logger                  — orchestration logger instance
        current_path   : str    — absolute path of this script's directory (required by AST_FACTORY)
    """
    # ------------------------------------------------------------------
    # Submission banner — matches the style used by main_auto_setup.py
    # ------------------------------------------------------------------
    print(f"\n{'='*100}")
    print(f"{'='*100}")
    print(f"  >>> PROCESSING SUBMISSION: {submission_dir.name} <<<")
    print(f"{'='*100}")
    print(f"{'='*100}\n")
    logger.info(f"\n{'='*100}")
    logger.info(f"{'='*100}")
    logger.info(f"  >>> PROCESSING SUBMISSION: {submission_dir.name} <<<")
    logger.info(f"{'='*100}")
    logger.info(f"{'='*100}\n")

    # ------------------------------------------------------------------
    # Step 1: Create outputs/ directory
    # ------------------------------------------------------------------
    output_dir = submission_dir / "outputs"
    # exist_ok=True is intentional — if outputs/ exists from a prior incomplete run
    # (no sentinel was written) we reuse it rather than overwriting partial results.
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output directory: {output_dir}")
    logger.info(f"Output directory: {output_dir}")

    # ------------------------------------------------------------------
    # Step 2: Collect shapefiles
    # ------------------------------------------------------------------
    jobs = _collect_shapefiles(submission_dir, output_dir)
    print(f"Collected {len(jobs)} shapefile(s) for submission '{submission_dir.name}'")
    logger.info(f"Collected {len(jobs)} shapefile(s) for submission '{submission_dir.name}'")

    if not jobs:
        # No shapefiles found — write a sentinel with 0 jobs so Jenkins does not
        # retry this empty folder every night.  The user must add shapefiles and
        # delete the sentinel to trigger processing.
        msg = (
            f"WARNING: No shapefiles found in submission '{submission_dir.name}'. "
            "Skipping batch run. Sentinel written with 0 jobs."
        )
        print(msg)
        logger.warning(msg)
        _write_completion_sentinel(output_dir, submission_dir, 0, 0, 0, [], logger)
        return

    # ------------------------------------------------------------------
    # Step 3: Auto-detect NR Regions via BCGW spatial join
    # ------------------------------------------------------------------
    print(f"Detecting NR Regions for {len(jobs)} shapefile(s) via BCGW...")
    logger.info(f"Detecting NR Regions for {len(jobs)} shapefile(s) via BCGW...")

    shp_paths = [shp for shp, _ in jobs]
    region_map = detect_regions(shp_paths, sde_path, logger)

    # Warn for any shapefile that could not be matched to a region — the cell will be
    # left blank in the Excel file and the AST tool will handle the missing value.
    for i, (shp_path, _) in enumerate(jobs):
        if not region_map.get(i):
            warn = f"WARNING: Could not detect region for '{shp_path.name}' — cell will be blank"
            print(warn)
            logger.warning(warn)

    # ------------------------------------------------------------------
    # Step 4: Write Excel job files into outputs/
    # ------------------------------------------------------------------
    print("Writing Excel job files...")
    logger.info("Writing Excel job files...")

    excel_files = _write_excel_files(jobs, region_map, output_dir)
    print(f"Wrote {len(excel_files)} Excel file(s) to {output_dir}")
    logger.info(f"Wrote {len(excel_files)} Excel file(s) to {output_dir}")

    # ------------------------------------------------------------------
    # Step 5: Run the batch tool for each Excel file
    # ------------------------------------------------------------------
    # FailedJobTracker is initialised pointing to the submission's outputs/ so the
    # failed-jobs summary spreadsheet is co-located with the other outputs and visible
    # to the submitting user on the objectstore drive.
    failed_job_tracker = FailedJobTracker(str(output_dir), logger)

    # delete_transitory is always False for Jenkins runs — transitory GDBs must be
    # retained so that any failed job can be diagnosed by the user or requeued manually.
    # If automatic cleanup is needed in future, add DELETE_TRANSITORY=true to .env.
    delete_transitory = False

    for excel_file in excel_files:
        print(f"\n{'='*100}")
        print(f"{'='*100}")
        print(f"  >>> PROCESSING SPREADSHEET: {excel_file} <<<")
        print(f"{'='*100}")
        print(f"{'='*100}\n")
        logger.info(f"\n{'='*100}")
        logger.info(f"{'='*100}")
        logger.info(f"  >>> PROCESSING SPREADSHEET: {excel_file} <<<")
        logger.info(f"{'='*100}")
        logger.info(f"{'='*100}\n")

        # excel_file is an absolute path string.  os.path.join(current_path, excel_file)
        # inside process_excel_file() discards current_path when excel_file is absolute,
        # so the correct absolute path (including UNC paths on the objectstore) is used.
        process_excel_file(
            excel_file, secrets, logger, current_path, failed_job_tracker, delete_transitory
        )

    # ------------------------------------------------------------------
    # Step 6: Count outcomes and write the completion sentinel
    # ------------------------------------------------------------------
    total, complete, failed_count = _count_job_statuses(excel_files)

    summary = (
        f"Submission '{submission_dir.name}': "
        f"{complete}/{total} jobs COMPLETE, {failed_count} failed/incomplete"
    )
    print(summary)
    logger.info(summary)

    # Sentinel is always written so Jenkins stops retrying after one full attempt.
    # The JSON payload records pass/fail counts for dashboard consumption.
    _write_completion_sentinel(
        output_dir, submission_dir, total, complete, failed_count, excel_files, logger
    )


# ---------------------------------------------------------------------------
# Main block
# ---------------------------------------------------------------------------

if __name__ == '__main__':

    # ------------------------------------------------------------------
    # Step 1: Load .env — all credentials and paths come from here;
    # nothing is hardcoded below per project rules.
    # ------------------------------------------------------------------
    load_dotenv()

    # ------------------------------------------------------------------
    # Step 2: Validate required .env keys
    # ------------------------------------------------------------------
    # WATCH_DIR is the objectstore root that Jenkins monitors.
    # Per project rules: missing required config must raise immediately — no fallback.
    WATCH_DIR = os.getenv("WATCH_DIR")
    if not WATCH_DIR:
        raise ValueError(
            "Required .env key 'WATCH_DIR' is missing or empty. "
            "Set WATCH_DIR to the objectstore root where users deposit shapefile submissions "
            r"(e.g., \\objectstore2.nrs.bcgov\GSS_Share\authorizations\batch_ast_tool)."
        )

    watch_path = Path(WATCH_DIR)
    if not watch_path.is_dir():
        raise ValueError(
            f"WATCH_DIR '{WATCH_DIR}' does not exist or is not accessible. "
            "Ensure the objectstore drive is mounted and the UNC path is correct."
        )

    # ------------------------------------------------------------------
    # Step 3: Set up logging — single top-level log written into WATCH_DIR
    # ------------------------------------------------------------------
    # Passing WATCH_DIR as the custom log path routes the orchestration log into
    # WATCH_DIR/autoast_batch_logs_folder_YYYYMMDD/ so it is visible to all users
    # on the objectstore drive alongside the submission outputs.
    logger = setup_logging(WATCH_DIR)
    logger.info(f"Jenkins run started. WATCH_DIR: {WATCH_DIR}")

    # ------------------------------------------------------------------
    # Step 4: Import the AST toolbox — done once for the entire run
    # ------------------------------------------------------------------
    template = import_ast(logger)

    # ------------------------------------------------------------------
    # Step 5: Establish a single shared BCGW connection
    # ------------------------------------------------------------------
    # One connection is created here in the main process and reused for region
    # detection (detect_regions) and all worker subprocesses (mp_worker.py)
    # via the SDE_FILE_PATH environment variable.  Per project rules: do not
    # create per-submission or per-worker connections.
    current_path = os.path.dirname(os.path.realpath(__file__))
    secrets, sde_connection, sde_path = setup_bcgw(logger)

    # Publish the SDE file path so mp_worker.py subprocesses can find the connection.
    os.environ["SDE_FILE_PATH"] = sde_path
    logger.info(f"SDE connection established at: {sde_path}")

    # ------------------------------------------------------------------
    # Step 6: Scan WATCH_DIR for pending submissions
    # ------------------------------------------------------------------
    # A submission is any immediate subdirectory of WATCH_DIR.
    # Submissions are sorted alphabetically so the processing order is deterministic.
    pending = []    # Subdirectories that need to be processed this run
    skipped = []    # Subdirectories whose sentinel already exists (already complete)

    for entry in sorted(watch_path.iterdir()):
        # Skip any non-directory items (loose files) at the watch dir root.
        if not entry.is_dir():
            continue

        # Skip the log folder that setup_logging() creates inside WATCH_DIR — it is
        # not a submission and must never be treated as one.
        if entry.name.startswith("autoast_batch_logs_folder_"):
            continue

        # Check for the completion sentinel.  If present, this submission was fully
        # processed on a previous Jenkins run and should not be reprocessed.
        if _is_submission_complete(entry):
            skipped.append(entry.name)
            logger.info(f"Skipping '{entry.name}' — BATCH_COMPLETE.json sentinel found")
            print(f"  Skipping '{entry.name}' — already complete")
        else:
            pending.append(entry)

    logger.info(
        f"Scan complete: {len(pending)} pending submission(s), {len(skipped)} already complete"
    )
    print(f"\nScan complete: {len(pending)} pending submission(s), {len(skipped)} skipped\n")

    if not pending:
        # Nothing to do — log cleanly so Jenkins marks the build as success (exit 0).
        print("No pending submissions found. Nothing to do.")
        logger.info("No pending submissions found. Jenkins run complete.")
    else:
        # ------------------------------------------------------------------
        # Step 7: Process each pending submission sequentially
        # ------------------------------------------------------------------
        # Submissions are processed one at a time rather than in parallel because
        # each submission already saturates the BCGW connection pool with up to 8
        # concurrent worker processes.  Running two submissions simultaneously would
        # exceed the 8-connection hard limit.
        for submission_dir in pending:
            process_submission(submission_dir, secrets, sde_path, logger, current_path)

        print("\nAll pending submissions processed.")
        logger.info("All pending submissions processed. Jenkins run complete.")

# AST Batch Tool — Development and Troubleshooting Notes

> Comprehensive development log covering the AST Batch Tool from initial setup through production deployment. Entries are in **reverse chronological order** (newest first).

---

## May 4, 2026 — Integration of New `auto_status` Runner (v3 Alpha Branch)

### Overview
Began adapting the batch tool to call the new developer-maintained AST tool (`fcbc_auto_status_tool.pyt` / `auto_status` package) instead of the old `ast.atbx`. This replaces the ArcPy `ImportToolbox` / `arcpy.fcbc_auto_status_tool.AutomatedStatusTool` call pattern with a direct Python function call to `auto_status.analysis_tool.run(raw, sde=Path(sde_path))`.

### Branch
`ast_alpha_runner_version`

### New AST Tool Location
`\\giswhse.env.gov.bc.ca\whse_np\corp\script_whse\python\Utility_Misc\Ready\statusing_tools_arcpro\Tools\src\auto_status\tools\fcbc_auto_status_tool.pyt`

A read-only local reference copy of the `auto_status` package lives in `autoast/ignorefolder/` (excluded from git via `.gitignore`).

### Architecture: Runner Direct-Call with Shared SDE
The new tool's entry point is `auto_status.analysis_tool.run(raw_params, *, sde=None)` in `runner.py`. If `sde` is supplied as a `Path` to a pre-made `.sde` file, the runner skips its internal BCGW connection entirely. This fits cleanly with the existing shared-SDE architecture:
- `database_connection.setup_bcgw()` creates `connection/bcgw.sde` once in the main process
- Path is set via `os.environ["SDE_FILE_PATH"]`
- Each worker receives `sde_path` as a function argument and passes `Path(sde_path)` to `run_auto_status(raw, sde=Path(sde_path))`

The old `arcpy.ImportToolbox` approach and `params[]` list-building code were commented out (marked `# version2`) rather than deleted.

### `sys.path` Injection for `auto_status` Package
`mp_worker.py` derives the package root from the `TOOLBOX` env var path:
```python
auto_status_src = str(Path(ast_toolbox_path.strip()).parents[2])  # resolves to .../Tools/src/
sys.path.insert(0, auto_status_src)
from auto_status.analysis_tool import run as run_auto_status  # type: ignore
```
The `# type: ignore` suppresses Pylance's `reportMissingImports` — the import is dynamic and cannot be statically resolved.

### Parameter Changes (Old → New)
The new tool uses renamed parameters. Old `AST_PARAMETERS` dict commented out; new dict and Excel column headers updated:

| Index | Old name | New name |
|-------|----------|----------|
| 6 | `output_directory_same_as_input` | `output_dir_same_as_input` |
| 7 | `dont_overwrite_outputs` | `dont_overwrite_outputs` (same) |
| 8 | `skip_conflicts_and_constraints` | `dont_run_conflicts_and_constraints_tab3` |
| 9 | `suppress_map_creation` | `suppress_map_creation_tab3` |
| 10 | `add_maps_to_current` | `open_output_directory` (must be `False` in batch) |
| 11 | `run_as_fcbc` | `full_path_hyperlinks` |

Three parameters are hardcoded in `mp_worker.py` (not from Excel): `open_output_directory=False`, `fcbc_spreadsheet_formatting=True`, `debug=False`.

### New Output Directory Structure
The new runner's `prepare_run_paths()` creates a `YYYYMMDD-Status-{suffix}` subdirectory **inside** the job's `output_directory`:
- `output_directory = ...\outputs\632\`
- Actual outputs written to: `...\outputs\632\20260504-Status-632\`

The suffix is derived from `crown_file_number` if provided, otherwise from the AOI shapefile stem (e.g., `632` from `632.shp`).

### New Output File/Folder Names
The new runner produces (inside the `YYYYMMDD-Status-*` subdir):
- `aoi_boundary.gdb`
- `maps/` (HTML map files)
- `one_status_common_datasets_aoi.gdb`
- `one_status_tabs_1_and_2_datasets.gdb`
- `{run_folder_name}.xlsx` — single merged Excel (e.g., `20260504-Status-632.xlsx`)
- `auto_status.log`
- `input_list.json`

The old tool's `mapx_files/`, `automated_status_sheet.xlsx`, `one_status_common_datasets_aoi.xlsx`, `one_status_tabs_1_and_2.xlsx` are **not produced** by the new tool.

### First Test Run — 4 Jobs (10:57 AM)
Ran `main_auto_setup.py` with `jobs_1.xlsx` (jobs 632, 634, 635, 636). All 4 workers returned `'Success'` and ran for ~15 minutes. However, all 4 were marked `FAILED_OUTPUTS` because `output_validator.py` was looking in `output_directory` directly (e.g., `632\`) instead of inside the `20260504-Status-632\` subdirectory.

`re_load_failed_jobs_V2` immediately requeued all 4 jobs with `dont_overwrite_outputs=True`. The retry (11:14 AM) also failed because `dont_overwrite_outputs=True` causes `prepare_run_paths()` to set `run_dir = base_dir` (no subdir created), so the runner wrote a **second set of outputs** directly to `632\` root. This created the double-output situation visible in the file explorer screenshots.

### Fixes Applied
Three issues identified and fixed:

**1. `output_validator.py` — wrong subdirectory detection pattern**
My initial fix used `d.startswith('Status_')` but the actual folder starts with the date (`20260504-Status-632`). Corrected to `'-Status-' in d or '-DEBUG-' in d`.

**2. `output_validator.py` — `required_items` wrong for new tool**
Removed `mapx_files`, `automated_status_sheet.xlsx`, `one_status_common_datasets_aoi.xlsx`, `one_status_tabs_1_and_2.xlsx`. Added a scan for any non-empty `.xlsx` file in the run directory instead of a fixed filename.

**3. `ast_factory.py` `re_load_failed_jobs_V2` — `FAILED_OUTPUTS` retry needed `output_directory` update**
When retrying a `FAILED_OUTPUTS` job with `dont_overwrite_outputs=True`, `prepare_run_paths()` sets `run_dir = base_dir = output_directory`. For the runner to land inside the correct existing `YYYYMMDD-Status-*` dir, `output_directory` must be updated to point at that subdir before the retry. The fix scans for the most recently modified `YYYYMMDD-Status-*` subdir and updates `job['output_directory']` to it before setting `dont_overwrite_outputs=True`.

### `multi_excel_setup.py` Column Headers Updated
Old column header names replaced with new parameter names to match the new tool. Existing `jobs_1.xlsx` files created before this change still have old column names — must be regenerated or manually renamed before re-running.

### `main.py` and `main_auto_setup.py` — `FailedJobTracker` Added
Both entry points updated to initialize `FailedJobTracker` and pass it through to `process_excel_file()`. This was a carry-over from the February work, completed at the start of this session.

### Known State Before Next Test Run
- `jobs_1.xlsx` on disk has `dont_overwrite_outputs: True` for all 4 rows (written during the failed retry). Must be reset to `false` before re-running.
- The `20260504-Status-*` subdirectories from run 1 exist and contain partial outputs.
- The second set of outputs at `632\` root (from the incorrect retry) also exists.

---

## February 4, 2026 — Skeena 2026-2028 Production Batch Run

- Ran `main_finish_jobs.py` on Skeena 2026-2028 batch using V2 with the new timeout and failed job tracking features
- Processed 25 Excel files (`jobs_1.xlsx` through `jobs_25.xlsx`) with 8 jobs per file
- Used `python_geopandas` clone environment
- Job timeout set to 21,600 seconds (6 hours) with process health checks every 60 seconds
- Failed Job Tracker created spreadsheet at output location logging all `missing_outputs` failures
- Jobs were monitored with the new failure classification system (timeout, crash, worker_error, missing_outputs)

---

## February 2, 2026 — Timeout Issue V2 Branch: Failed Job Tracker & Health Checks

- Implemented new module: `failed_job_tracker.py`
  - Creates a timestamped Excel file (e.g., `failed_jobs_20260202_143052.xlsx`)
  - Logs each failure with details: spreadsheet name, job index, failure type, timestamp, and all job parameters
  - Color-codes failures by type (timeout=red, crash=orange, worker_error=yellow)
  - Provides summary statistics at the end
- Re-enabled job timeout at 6 hours (21,600 seconds)
- Added process health checks every 60 seconds to detect crashed workers
- Improved failure classification: `timeout`, `crash`, `worker_error`, `unknown`
- Key benefit: crashed jobs (like 6408157) are now detected within 60 seconds instead of hanging indefinitely
- One failure no longer stops the entire batch

---

## January 23, 2026 — Fixing ORA-02391 Database Session Limit Errors

### Problem
- Jobs 0 and 5 failed with generic `ERROR 000582: Error occurred during execution` after ~31 minutes of processing
- 8 concurrent workers in multiprocessing mode
- XTools module import warnings appeared in logs but were **not the cause** (harmless warning)

### Root Cause
- **Actual Error:** `ORA-02391: exceeded simultaneous SESSIONS_PER_USER limit`
- Each worker called `inactive_dispositions.execute_process()` which created its **own** pyodbc database connection to BCGW
- Resulted in 8+ simultaneous Oracle sessions for the same user
- The connection failure was masked by a bare `except:` clause, causing a generic `ERROR 000582` instead of showing the actual Oracle error
- Evidence from logs: `pyodbc.Error: ('HY000', '[HY000] [Oracle][ODBC][Ora]ORA-02391: exceeded simultaneous SESSIONS_PER_USER limit`

### Surgical Fix
- Modified `inactive_dispositions.py`:
  - Added optional `sde_connection` parameter to `execute_process()`
  - Created new `read_query_sde()` function using `arcpy.ArcSDESQLExecute` to execute SQL via existing SDE connection (no new session created)
  - Falls back to pyodbc if SDE connection not provided (backwards compatible)
  - Removed import-time SDE validation that prevented testing
- Modified `one_status_tabs_one_and_two_arcpro.py`:
  - Pass `self.sde_connection` to `execute_process()` (line 345)
- **Connection method:** `arcpy.ArcSDESQLExecute(sde_path)` — uses existing SDE connection file created by `database_connection.setup_bcgw()`
- **Verification:** Created `test_inactive_dispositions.py` — DataFrames from both methods (pyodbc vs SDE) were identical
- Recommendations: Deploy to production, monitor first batch run, consider removing pyodbc fallback in future

---

## December 11, 2025 — Ideas for BCGW Credential Interface

- Concept: Create an interface where the user inputs username and password; the script stores them in a variable that is deleted later

---

## October 28, 2025 — Hyperlinks Troubleshooting (Continued)

- Hyperlink tool runs in 1 second but doesn't apply hyperlinks correctly
- Copied outputs to csostad root, ran tool — **it works**, tested the links
- Copied outputs to `csostad`, ran hyperlinks, then copied outputs to gss share
- Also copied from gss `projects/work` to `sandbox/csostad` to test if shorter file path fixes the issue
- **Finding:** If outputs are copied to sandbox and hyperlinks tool is run, it works. If outputs are in `projects/work`, the tool seems to not work
- However, upon further testing, both locations seemed to work — may have made a mistake, needs retesting
- Sunny ran a status (663) to `gss/projects` — tool ran in 1 second and hyperlinks didn't work
- Sunny ran the hyperlink tool for outputs in `SharedWork\gr_2025_26` directory and it **worked**

---

## September 25, 2025 — Alpha .atbx Performance Comparison

- Confirmed that directing the batch tool to Steve's `.atbx` causes BCGW connection pop-ups
- Ran modified `.atbx` in ArcGIS Pro — runs successfully on its own but uses user credentials
- **Performance comparison on file 632:**
  - Batch alpha: **44.02 minutes**
  - Steve's P Corp Alpha: **53 minutes**
- Need to check file size of outputs to verify all versions produce the same results

---

## September 23, 2025 — Repointing Batch Tool to P Corp Alpha

- Updated `.env` file TOOLBOX path to: `\\giswhse.env.gov.bc.ca\whse_np\corp\script_whse\python\Utility_Misc\Ready\statusing_tools_arcpro\alpha\Alpha_Statusing_Tools_ArcPRO.atbx`
- Opened the `.atbx` in P Corp and viewed the properties

---

## July 22, 2025 — Goals

- Test V2, V3, and Alpha
- Pull a copy of the AST alpha into ArcPro, change the naming convention, and test it from within ArcPro
- Pull batch tool into a geoprocessing tool and test running `main.py`

---

## July 14-15, 2025 — Alpha AST with Multi Excel: Fiona Backslash Fix

- **Problem:** Fiona module fails when paths contain mixed forward/backslash combinations
  - `fiona._err.CPLE_OpenFailedError` — path `/work/srm/nel/.../v3 4 JOBS\outputs\636\one_status_common_datasets_aoi.gdb` does not exist
  - The `create_multi_excel` output was producing forward slashes, but the AST tool expected backslashes
- **XTools warning** (harmless): `ModuleNotFoundError: No module named 'XTools.XToolsAGP.Core'` — this is a non-issue with the python_geopandas clone
- **Fix:** Modified `create_multi_excel` to use only backslashes in all path outputs
- **Result:** Alpha AST with Multi Excel ran **successfully** after the fix

<!-- Screenshot: Verbose intersecting layer output showing 4 concurrent workers processing layers 88-240 -->

---

## July 11, 2025 — Fiona Module Path Issues

- Fiona module issues with forward/backslash path combinations in `create_multi_excel` output
- V2: **Working**
- V3: Working on single Excel, issues with multi-Excel paths

---

## July 4, 2025 — Setting Up Final Folium AST in Batch Tool

- Copied V3 to v3 Alpha
- Changed all `sys.path.append` file paths to point to the new alpha location
- Need to change TOOLBOX in `.env`
- **Discovery:** When changing everything except TOOLBOX in `.env`, the tool still worked but produced PDF maps instead of HTML/folium maps
- Scripts added to the folder: `.env`, `aoi_utilities.py`, `ast_config`, `ast_factory`, `database_connection.py`, `logging_setup.py`, `main_auto_setup.py`, `main.py`, `mp_worker.py`, `multi_excel_setup.py`, `toolbox_import.py`
- Need to update `.env` TOOLBOXALIAS from `alphaast` to `AlphaStatusingToolsArcPRO`
- Also need to update `mp_worker.py` to correct alias
- Brought in alpha from `P:\corp\script_whse\python\Utility_Misc\Ready\statusing_tools_arcpro\alpha`
- Copied to GitHub repo `ast_batch_tool\alpha`
- Opened Steve's Alpha and V3 Breville side by side in compare window and implemented changes into Steve's alpha
- Modifications to `ast_sheet_call_routine`:
  - Imported DotEnv to circumvent BCGW connections
  - Implemented logging
  - Commented out BCGW connection check
  - Commented out line 572 (cleanup temporary SDE file)
  - Imported dotenv to share BCGW credentials
  - In `GETINACTIVES`: find the secrets file and assign BCGW user/pw to username and password
- Fix relative hyperlinks for view maps and overview maps
- `OUT Call Routine arcpro.py`: Added get env SDE file path, commented out `oracleCreds`

<!-- Screenshot: ArcGIS Pro toolbox properties showing alpha_ast.atbx configuration -->
<!-- Screenshot: Side-by-side code comparison of Steve's Alpha and V3 Breville -->
<!-- Screenshot: DotEnv import modifications for BCGW connection circumvention -->
<!-- Screenshot: OUT Call Routine arcpro.py SDE path changes -->

### Toolbox Naming Issues
- `Alpha Statusing Tools ArcPRO.tbx` — wasn't working
- In V2 and V3, used `alpha_ast.atbx`
- Renamed new `.tbx` to `alpha_ast.TBX` — **didn't work**
- Opened ArcPro, imported old `.tbx`, used convert tool (right click), saved as `alpha_ast.atbx`
- **Discovery:** The script may not be able to import `.tbx` files — only `.atbx` works
- Removed the hyphen in the label to `alphaast`

---

## June 30, 2025 — Updating V3 with V2 Changes

- Merging V2 improvements (BCGW handling and multi-Excel support) into V3

---

## June 16, 2025 — Documentation for Batch Tool Spreadsheet Creator

- Documented the batch tool spreadsheet creator functionality
- Updated `sys.path.append` references to point to correct directories

---

## June 12-13, 2025 — Multi-Excel Handling Milestone

- Created `create_job_excel_files()` function
  - Uses Tkinter dialogs for folder selection
  - Automatically discovers shapefiles in the selected folder
  - Creates batches of 8 jobs per Excel file
  - Names files sequentially: `jobs_1.xlsx`, `jobs_2.xlsx`, etc.
- Tested on 27 jobs — produced 4 Excel files (3 × 8 jobs + 1 × 3 jobs)
- This solves the Oracle session limit issue by limiting concurrent database connections

---

## June 11, 2025 — Oracle SESSIONS_PER_USER Limit & Multi-Excel Concept

- **Error:** `ORA-02391: exceeded simultaneous SESSIONS_PER_USER limit` in `inactive_dispositions.py`
- When running 8+ concurrent workers, each creates its own database connection
- Oracle's per-user session limit is reached, causing `ERROR 000582`
- **Solution concept:** Split large batches into multiple Excel files of 8 jobs each, process sequentially
- Tested on 27 jobs in batches of 8

---

## June 9-10, 2025 — V2 BCGW Connection Issues & Cross-Version Testing

- V2 failing due to too many BCGW connections
- Tested multiple branches and backups
- V3 on Main failed with Fiona error (forward/backslash paths)
- **V1 on gss resources: SUCCESS**
- Ran Jordan's 6 jobs on V2 overnight — **SUCCESS**
- V2 branch status: working intermittently depending on batch size

---

## May 15, 2025 — Hyperlinks Failure Investigation

- Chris ran `gr_2025_26` batch of 8 statuses to workarea
- In the past, this location should have worked, but this time it failed
- Theory: Perhaps Batch Tool V2 is the issue because the BCGW connection script was edited
- Tested by running Batch Tool V1 to the same location — hyperlink tool **worked**
- One hyperlink output showing up with URL-encoded spaces (`%20`) in the path
- If link is manually edited and corrected to UNC path, the hyperlink works
- Status 136 is working for the client at `\\objectstore2.nrs.bcgov\GSS_Share\projects\gr_2025_26_batch_status_request\Updated_outputs\136_south coast`

---

## May 5-9, 2025 — Hyperlinks Troubleshooting Notes

- **May 5:** Ran Batch Tool V3 to `\WildLifePermittingTest\AST_OUTPUTS\V3outputs` — hyperlinks don't work
- Ran Batch Tool V3 with outputs to `\WildLifePermittingTest\AST_OUTPUTS\2\` — **hyperlinks work**
- **May 6:** Ran hyperlinks tool on Windfarm AST outputs using the Windfarm Project Master `.aprx` template — **worked**
- Successfully ran `gr_2025_26` batch on V3 using March 10th branch with geospatial geopandas clone
- **May 9:** Ran batch status on `gr_2025_26` with outputs going directly in the gss projects

---

## March 20, 2025 — Training Video & 15-Status Milestone

- Recorded a training video for the batch tool
- Completed **15 full statuses in 2 hours 10 minutes** using V3
- This demonstrates significant time savings compared to manual processing

---

## March 19, 2025 — V2 Code Changes for BCGW Handling

### Files Modified
- **`inactive_dispositions.py`:** Added `SDE_FILE_PATH` from environment variables
- **`one_status_tabs_one_and_two_arcpro.py`:** Implemented dotenv for BCGW credentials
- **`universal_overlap_tool_arcpro.py`:** Map creation path changed to T: drive
- **`universal_overlap_tool_call_routine.py`:** Updated BCGW connection handling

---

## March 17, 2025 — BCGW Connection Script Changes

- Changed BCGW connections in script to use centralized approach
- **Errors encountered:** `MakeQueryTable ERROR 000793` and `ERROR 000840`
- These errors were related to the new BCGW connection method not properly resolving table references

---

## March 13, 2025 — AST Background Information Documentation

- Created documentation explaining what the AST tool does and how it works
- Background information for onboarding and reference

---

## March 12, 2025 — V1 Regular Errors & Share Drive Issues

- V1 producing regular errors during batch processing
- Read/write issues when outputting to share drive locations
- Tested different output paths to isolate the problem

---

## March 10, 2025 — BCGW Fix Branch: Centralized Database Connection

- Created the "BCGW fix" branch
- **Approach:** Use a single `.sde` connection file shared across all multiprocessing workers
- Instead of each worker creating its own BCGW connection, all workers reference the same pre-established SDE connection
- This reduces the number of Oracle sessions and prevents connection pop-ups

---

## March 7, 2025 — V3 Testing with 15 Shapefiles

- Tested V3 (Breville/folium maps) with 15 shapefiles
- SDE connection logging showed SQLite DBMS connections being created
- BCGW credential pop-ups occurring during batch processing, interrupting automation
- Ran 4-7 jobs with varying success rates
- Some jobs completed successfully while others failed due to connection issues

<!-- Screenshot: SDE connection logs showing SQLite DBMS connections -->
<!-- Screenshot: BCGW credential pop-up dialog -->

---

## February 19, 2025 — Version Status Summary

| Version | Status |
|---------|--------|
| V1 (ToastMaster) | **Working** |
| V2 (Cuisinart) — All-in-one | Not Working |
| V2 (Cuisinart) — Modularized | Not Working |
| V3 (Breville/Folium) | **Working** |

---

## February 7, 2025 — Modularization Branch

- V1 (ToastMaster): **Working**
- V2 (Cuisinart): **Working** after modularization
- V3 (Breville): Troubleshooting in progress
- Modularization effort: splitting monolithic scripts into separate, importable modules

---

## January 10-20, 2025 — Alpha Modularization Testing

- Testing alpha version with modularized code structure
- Map generation errors encountered during testing
- Iterating on module imports and dependency resolution

---

## January 9, 2025 — Alpha 4 Testing & Toolbox Alias Discovery

- Testing alpha version 4 of the batch tool
- **Key discovery:** ArcPy toolbox naming convention
  - When importing a toolbox, ArcPy uses the toolbox name as a suffix: `arcpy.MakeAutomatedStatusSpreadsheet_ast`
  - The `_ast` suffix indicates the toolbox is named "ast"
  - Use `arcpy.ImportToolbox()` with an alias parameter to control the namespace
- This discovery was critical for the multiprocessing workers to correctly invoke the toolbox tool

---

## December 19, 2024 — Alpha Version Setup

- Setting up a new alpha version of the AST tool
- Configured `.env` file and toolbox paths
- **Discovery:** `.tbx` (legacy) vs `.atbx` (ArcGIS Pro native) format
  - The batch tool requires `.atbx` format
  - `.tbx` files need to be converted using ArcPro's built-in converter (right-click > Save As `.atbx`)

---

## December 11, 2024 — Cariboo Jobs Troubleshooting

- Troubleshooting Cariboo region batch jobs
- **Slow mode vs fast mode behavior:** Different execution speeds observed depending on BCGW connection method
- Identified that BCGW creates temporary connection files in `AppData\Local\Temp`
- These temp connections can accumulate and cause issues if not cleaned up

---

## November 19-21, 2024 — V1/V2 Testing & Multiprocessing Issues

### Issues Discovered
1. **Multiprocessing first-pass failures:** First job in each batch often fails while subsequent jobs succeed — likely a race condition during initialization
2. **`dont_overwrite_outputs` GDB conflict:** When this flag is enabled, multiple workers try to check/access the same geodatabase simultaneously, causing file locks
3. **WinError 123 — Trailing newline in paths:** Excel file paths containing trailing newline characters (`\n`) cause `WinError 123: The filename, directory name, or volume label syntax is incorrect`
   - **Root cause:** Excel cells containing line breaks
   - **Fix:** Strip whitespace/newlines from paths read from Excel
4. **Folder creation issues:** Output folder creation fails when multiple workers attempt to create the same directory tree simultaneously

---

## July 30, 2024 — Batch Error on 4-File Test

- Ran Evan's batch of 4 files
- **Error:** `WinError 32` — file lock error
  - The process cannot access the file because it is being used by another process
  - Caused by multiple workers trying to write to the same output location or access the same GDB

---

## July 10, 2024 — Initial AutoAST Setup

- Initial setup of the AutoAST batch processing tool
- Configured `dotenv` for environment variable management (BCGW credentials, toolbox paths)
- Received tester data from Evan Breton:
  - Crown Land files with DID numbers
  - Used for initial testing and validation
- Tool description: Automates the creation of AST (Automated Status Tool) spreadsheets in batch mode using multiprocessing
- Each worker process imports the ArcGIS Pro toolbox and runs `MakeAutomatedStatusSpreadsheet` with job-specific parameters

---

## Key Recurring Issues Reference

### BCGW/Oracle Connection Problems
- **ORA-02391:** Exceeded `SESSIONS_PER_USER` limit when running 8+ concurrent workers, each creating its own connection
- **Solution evolution:** dotenv credentials → centralized SDE connection → `arcpy.ArcSDESQLExecute` reusing single SDE file

### Fiona/GeoPandas Path Issues
- Fiona expects consistent path separators (all forward or all backslash)
- Mixed paths (e.g., `//server/path\subfolder`) cause `CPLE_OpenFailedError`
- Fix: Normalize all paths to backslashes before passing to Fiona

### Toolbox Import Issues
- `.tbx` (legacy) format cannot be imported by the batch tool — must use `.atbx` (ArcGIS Pro native)
- Toolbox alias is critical: `arcpy.ImportToolbox(path, alias)` — alias determines the calling convention (e.g., `arcpy.alphaast.MakeAutomatedStatusSpreadsheet`)

### Hyperlinks Tool Behavior
- Works correctly when outputs are in local/workarea paths
- Unreliable when outputs are in deeply nested `gss/projects` paths
- URL-encoded spaces (`%20`) in paths can break hyperlinks
- V1 (unedited) hyperlinks work more reliably than V2 (edited BCGW scripts)

### XTools Warning (Non-Issue)
- `ModuleNotFoundError: No module named 'XTools.XToolsAGP.Core'` — harmless warning from `python_geopandas` clone, does not affect processing

---

*Note: Screenshots referenced in the original document could not be extracted during conversion. Screenshot locations are indicated with HTML comments throughout the document.*

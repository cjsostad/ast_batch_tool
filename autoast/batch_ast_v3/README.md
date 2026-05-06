# autoast
This will be a tool to manage and automate status processes from a queue. 

# requirements
geopandas  
openpyxl  
arcpy  
automated status tool

## .env Configuration Keys

All credentials and paths are read from a `.env` file in the script directory (never hardcoded).

| Key | Required by | Description |
|-----|-------------|-------------|
| `BCGW_USER` | all entry points | BCGW Oracle database username |
| `BCGW_PASS` | all entry points | BCGW Oracle database password |
| `TOOLBOX` | all entry points | Absolute path to the AST toolbox file |
| `TOOLBOXALIAS` | all entry points | Toolbox alias used by arcpy |
| `TEMPLATE` | all entry points | Path to the Excel template file |
| `SECRET_FILE` | all entry points | Path to a secondary credentials file (optional) |
| `SDE_FILE_PATH` | workers | Set at runtime by the main process; do not set manually |
| `WATCH_DIR` | `main_auto_setup_jenkins.py` | Root folder on the objectstore where users deposit shapefile submissions for automated nightly processing (e.g., `\\objectstore2.nrs.bcgov\GSS_Share\authorizations\batch_ast_tool`) |
| `WORKER_LOG_DIR` | `mp_worker.py` | Directory where per-job ArcPy worker logs are written. When set, overrides the default (next to the script) so Jenkins runs write logs to a user-visible location (e.g., `\\objectstore2.nrs.bcgov\GSS_Share\authorizations\batch_ast_tool\worker_log_files`). Optional — omit to keep logs next to the script. |

You need a test excel spreadsheet. 
If file_number is filled out, the script will run the FW Setup tool on the feature layer. Enter the file number of the permit and it
will create shapefiles and .kml files in the appropriate directory the way the old FW Setup did. 

If file number is left blank, the script will pass the raw shapefile or .kml into the Ast Toolbox.
Be sure to update the output directory to the output where you want the results of the AST Toolbox to be placed. 


## Development History

_This project has been under active development since July 2024, originally in the bcgov/gss_authorizations repository. On July 16, 2025, it was migrated to the cjsostad/ast_batch_tool repository for enhanced version control. History below is presented in reverse chronological order (newest first)._

---

### March 2026

#### March 12, 2026 - .gitignore Refactor and File Cleanup Utilities
- **Refactored .gitignore**: Converted 188+ individual file entries to pattern matching
  - Added `*.xlsx`, `*.log`, `*.pyc`, `*.pyo`, `*.pyd`, `__pycache__/`, `.env`, `*.sde` patterns
  - More maintainable and comprehensive ignore rules
- **Renamed DeleteASTFiles_arcpro.py** to `delete_transitory_data.py` for better clarity
- **Updated kmlToShpInFolders.py** source folder path to sandbox location

### February 2026

#### February 5, 2026 - Delete Transitory Data Script
- **New Module**: `delete_transitory_data.py`
  - Recursively deletes transitory data from output folders (6 items: aoi_boundary.gdb, mapx_files, 2 GDBs, 2 Excel files)
  - Can run standalone or be imported as a module
  - Hardcoded default path with option to override
  - Comprehensive error handling and statistics reporting

#### February 4, 2026 - Output Validation System and Job Verification
- **New Module**: `output_validator.py`
  - Validates that all required output files and folders exist and are not empty
  - Checks for 8 required items: geodatabases, maps, mapx files, and Excel files
  - Integrated into both `ast_factory.py` and `main_finish_jobs.py`
- **Enhanced ast_factory.py**: 
  - Real-time output validation during job processing
  - Jobs now marked as `FAILED_OUTPUTS` if outputs are missing/invalid
  - Missing outputs logged to failed_job_tracker with 'missing_outputs' type
  - Detailed logging for success and failure paths
- **Updated main_finish_jobs.py**:
  - Fixed syntax error in `__init__` method
  - Replaced hardcoded spreadsheet list with dynamic discovery using `os.listdir()`
  - Fixed worksheet name from 'Sheet1' to 'ast_config'
  - Added output verification for jobs marked COMPLETE
  - Enhanced `has_incomplete_jobs()` to validate outputs
- **Updated .gitignore**: General cleanup
- **Updated failed_job_tracker.py**: 
  - Added light purple color coding for missing_outputs failures
  - Added missing_outputs count to summary report

#### February 2, 2026 - Failed Job Tracking and Timeout Management
- **New Module**: `failed_job_tracker.py`
  - Automatically tracks failed jobs and creates timestamped Excel spreadsheets
  - Logs multiple failure types: timeout, crash, worker_error, missing_outputs, unknown
  - Color-codes failures by type (timeout=red, crash=orange, worker_error=yellow, missing_outputs=light purple)
  - Captures full job details for easy retry
  - Provides summary statistics at the end
  - Integrated into batch processing workflow
- **Re-enabled Job Timeout**: Set to 6 hours (21,600 seconds)
  - Jobs exceeding limit are terminated and marked as Failed
  - Timeout failures logged to failed job tracker
- **Process Health Checks**: 
  - Every 60 seconds, system checks if worker processes are still alive
  - Detects processes that crash without reporting
  - Crashed processes detected within 60 seconds instead of waiting indefinitely
- **Enhanced Worker Log Preservation**: 
  - Comprehensive logging of all multiprocessing worker activities
  - Individual log files for each worker process (includes PID and job index)
  - Detailed execution traces for troubleshooting failed jobs
  - Logs preserved in timestamped folders for historical analysis
  - Enables post-mortem debugging of crashed or hung processes

### January 2026

#### January 29, 2026 - Timeout and Retry Controls
- **Disabled Job Timeout Monitoring**: 
  - Commented out the 3-hour timeout that was killing legitimate long-running jobs
  - Jobs now run to completion naturally without being terminated
  - Added logging: "Timeout monitoring DISABLED - jobs will run to completion"
- **Disabled Automatic Retries**: 
  - Changed `max_retry_attempts` from 2 to 0 in `main_finish_jobs.py`
  - Failed jobs no longer retried automatically
- **Enhanced Spreadsheet Progress Logging**: 
  - Added prominent 100-character banners when starting/completing each spreadsheet
  - Makes it easy to track progress in logs
  - Applied to both `main_finish_jobs.py` and `main_auto_setup.py`
- **Custom Logging Path Selection**: 
  - Added Tkinter GUI dialog for selecting custom log file locations
  - Modified `multi_excel_setup.py` to support user-defined log paths
  - Updated `logging_setup.py` to write to custom directories
  - Modified `main_auto_setup.py` and `main_finish_jobs.py` to accept custom log paths
  - Logs now saved to user-specified directory instead of default location
  - Integration tested and verified working
  - Improves organization for large batch processing operations

#### January 27, 2026 - Smart Job Completion and Timeout Adjustments
- **New Module**: `main_finish_jobs.py`
  - Intelligently processes only incomplete Excel files
  - `has_incomplete_jobs()` function checks ast_condition column before processing
  - Skips workbooks where all jobs are already COMPLETE
  - Supports both initial processing and retry of failed jobs
  - Up to 2 additional retry attempts for failed jobs (later disabled)
  - Reduces wasted processing time by skipping completed work
  - Provides detailed feedback on job status
  - Better progress tracking: Shows number of incomplete jobs per workbook
  - Tracks which workbooks were processed vs skipped
  - Final summary of processed/skipped workbooks
  - Improved logging throughout with timestamps and descriptive messages
- **Job Timeout Reduced**: From 24 hours to 3 hours (10,800 seconds)
  - Added descriptive print statements showing the new timeout
  - Better timeout handling with descriptive messages
  - Prevents indefinitely hanging jobs
- **Enhanced .gitignore**: Added more patterns to improve repository management

#### January 26, 2026 - SDE Connection Improvements
- **Modified**: `inactive_dispositions.py` and `one_status_tabs_one_and_two_arcpro.py`
  - Added `sde_connection` parameter to `execute_process()` function (optional for backwards compatibility)
  - New `read_query_sde()` function uses `arcpy.ArcSDESQLExecute`
  - Reuses existing SDE connection instead of creating new sessions
  - Modified logic to use SDE connection when provided, fall back to pyodbc if not
  - Removed import-time SDE validation that prevented testing
  - Prevents "too many database connections" errors
- **Modified**: `one_status_tabs_one_and_two_arcpro.py`
  - Line 345: Pass `self.sde_connection` to `inactives.execute_process()`

#### January 22, 2026 - KML Conversion Tool
- **New Tool**: `kmlToShpInFolders.py` (192 lines)
  - Batch converts KML files to shapefiles from folder structures
  - Recursively processes nested folder hierarchies
  - Automatically creates output directory structure matching input
  - Simplifies data preparation workflow for AST processing
  - Widens Tkinter dialog window for better usability
  - Located in `autoast/ast supporting tools/` folder

### December 2025

#### December 10, 2025 - PathLib Migration and Code Cleanup
- **Repathed All Modules**: Migrated to use `pathlib.Path` instead of string-based paths
  - Tool now portable - can be moved without breaking file path references
  - Replaced hardcoded absolute paths with dynamic path resolution
  - More robust path handling across different environments and network drives
  - Updated modules: `automated_status_sheet_call_routine_arcpro.py`, `logging_setup.py`, `main.py`, `one_status_tabs_one_and_two_arcpro.py`, `universal_overlap_tool_arcpro.py`
  - Deleted `arcpro_main.py` (consolidated into main.py)
  - Removed `multi_excel_setup_OLD.py` (obsolete version)
- **Repository Cleanup**: 
  - Cleaned up V1 and V3 directories (deleted old versions)
  - Prepared repository for sharing on P: Drive
  - More maintainable codebase with single active version (V2)

### September 2025

#### September 25, 2025 - Standalone Mode Fix
- **Added `load_dotenv()`** to `automated_status_sheet_call_routine_arcpro.py`
  - Critical fix for running AST tool in standalone mode in ArcGIS Pro
  - Properly loads environment variables from `.env` file when not using batch processing
  - Ensures database credentials and configuration are available
  - Enables single-job execution directly from ArcGIS Pro toolbox
  - Previously only worked when called from batch processing scripts
  - Tested and verified working in standalone ArcGIS Pro environment

#### September 23, 2025 - BCGW Connection and Testing
- **Updated SDE file path** in `.env` file
  - Corrected database connection string for BCGW access
  - Fixed path to SDE connection file
- **Comprehensive Testing**: 
  - Tested AST Alpha version - verified working on September 23, 2025
  - BCGW Connection tested and verified (prompts for credentials on first run)
  - Successful connection to provincial database
- **Initial Full System Test**: 
  - Script tested with real data but not run to full completion
  - Validation of end-to-end workflow
  - Identified and resolved connection issues

### July 2025

#### July 23, 2025 - Troubleshooting Enhancements
- **Added `arcpy.AddMessage()` calls** throughout multiple scripts
  - Enables real-time progress monitoring in ArcGIS Pro interface
  - Messages appear in ArcGIS Pro geoprocessing results window
  - Improved debugging capabilities during development
  - Better visibility into script execution flow for end users
  - Added to critical decision points and processing stages
  - Helps identify where scripts fail or hang during execution

#### July 22, 2025 - Multi-Version Testing and ArcGIS Pro Integration
- **Comprehensive Version Testing**: Alpha, V2, V3 all verified working
  - Multi-version validation ensures backward compatibility
  - Each version tested with production data
- **Created `arcpro_main.py`** (148 lines): New entry point for ArcGIS Pro toolbox integration
  - Allows batch tool to be called from within ArcGIS Pro as a toolbox script
  - Bridges Python script and ArcGIS Pro toolbox interface
  - Enables GUI-based batch processing workflow
  - Prepared foundation for batch_arcpro_gui branch development
- **PathLib Implementation**: Applied to `multi_excel_setup.py`
  - Migrated from string-based paths to `pathlib.Path` objects
  - Renamed original to `multi_excel_setup_OLD.py` for reference
  - Improved cross-platform path handling
  - More Pythonic and modern code structure
- **UI Enhancements**: 
  - Widened Tkinter dialog window for better usability
  - Improved visibility of long file paths
  - Enhanced user experience for file selection dialogs

#### July 16, 2025 - Initial Commit
- **Repository Created**: First commit of AST Batch Tool to GitHub
- **Core Architecture Established**:
  - `ast_factory.py` - Main batch processing factory with multiprocessing support
  - `automated_status_sheet_call_routine_arcpro.py` - Core AST calling routine for ArcGIS Pro
  - `main_auto_setup.py` - Main setup and execution script
  - `main.py` - Alternative entry point for batch processing
  - `multi_excel_setup.py` - Excel file batch generator for splitting large jobs
  - `mp_worker.py` - Multiprocessing worker process handler
  - `logging_setup.py` - Centralized logging configuration
  - `database_connection.py` - BCGW database connection management
  - `toolbox_import.py` - ArcGIS toolbox import utilities
- **Supporting Modules**:
  - `aoi_utilities.py` - Area of Interest helper functions
  - `one_status_tabs_one_and_two_arcpro.py` - Status tab generation
  - `universal_overlap_tool_arcpro.py` - Overlap analysis tool
  - `inactive_dispositions.py` - Inactive disposition handling
- **Configuration**: 
  - `.env` file support for environment variables
  - SDE connection file management
  - Excel-based job queue configuration
- **Foundation**: Established comprehensive architecture for automated, parallel batch processing of AST jobs with robust error handling and logging

---

### Pre-Migration History (bcgov/gss_authorizations Repository)

_The following history documents development in the original bcgov/gss_authorizations repository before the project was migrated to its own repository on July 16, 2025._

#### July 11, 2025 - Pre-Migration Testing
- **Final Testing Before Repository Migration**: Tested V2 and V3 on single july_7.xlsx file
  - Verified both versions working correctly
  - Confirmed batch processing functionality
  - Validated multiprocessing and logging
  - Last commit before migrating to new GitHub repository

#### July 2, 2025 - V3 Multi-Excel Validation
- **Tested V3 After Multi-Excel Enhancements**: Comprehensive testing successful
  - Multi-spreadsheet processing verified
  - Confirmed backward compatibility with V2

### June 2025

#### June 30, 2025 - Cross-Version Testing and Maintenance
- **Extensive Testing**: V2 with `main_multi_excel_setup.py` working
- **File Management**: Cleaned out many old Excel test files
- **.env File Recovery**: Discovered `.env` missing from V2 due to `.gitignore`, restored from backup
- **Code Organization**: Moved `main.py` and `main_multi_set_up.py` from V2 to V3 for testing

#### June 16, 2025 - Environment File Security
- **Untracked .env Files**: Removed `.env` from version control for V1, V2, and V3
  - Used `git rm --cached` to untrack without deleting
  - Added `.env` to `.gitignore` for all versions
  - Improved security by preventing credential exposure

#### June 13, 2025 - Multi-Excel Setup Development (Major Update)
- **Merged `Running-Multiple-Excel-Sheets-on-V2` Branch into Main**
- **New Module**: `multi_excel_setup.py`
  - Sets up Excel files for the batch tool into worksheets of 8 jobs each
  - Tkinter dialogue box for adding region information
  - Automatically groups jobs for optimal processing
- **Created `main_auto_setup.py`**: Combines `main.py` with `multi_excel_setup.py`
  - Creates all spreadsheets in groups of 8 with common region
  - Feeds directly into batch processing ("the toaster")
- **Code Organization**: Moved copy of `multi_excel_setup.py` to supporting tools folder for team access

#### June 11-12, 2025 - Multi-Excel Batch Processing Testing
- **Successfully tested multi-spreadsheet processing**:
  - Initial test: Excel 1 and 2 (2 jobs of 2) successful
  - Full test: V2 multi excel on 3 Excel sheets (mat_excel 1, 2, 3) all completed
  - Moved `create_excel.py` into V2 folder, renamed to `multi_excel_setup.py`
- **File Organization**: Renamed the Hyperlinks Folder

### May 2025

#### May 30, 2025 - Relative Hyperlinks Implementation (Major Feature)
- **Renamed V2** to "V2_Cuisinart_MultiP_PdfMaps" to reflect hyperlinks change
- **Major Update - Relative Hyperlinks**:
  1. Created beginning of Toaster GUI for ArcGIS Pro
  2. **Edited UOT Tool**: Modified Universal Overlap Tool to create **relative hyperlinks** instead of absolute paths — critical for portability and sharing of output files
  3. **Redirected AST Call Routine** to call new UOT tool; old UOT Tool preserved as original

#### May 20-21, 2025 - Hyperlinks QA/QC and UOT Refinements
- **UOT Tool Fix**: Changed 2 lines of code around line 1212, initial test successful
- **QA/QC Hyperlinks Script**: Finished script that writes the IDIR of the person who ran status for tracking; yes/no dialogue still not fully functional
- **Added 2 new scripts**: Single-folder and multi-folder hyperlink changers for AST outputs
- **Backup commit**: Created backup before forking to fix hyperlinks

#### May 5-7, 2025 - Hyperlinks and Excel Setup Foundations
- **Created hyperlinks .py files**: Foundation for relative hyperlinks implementation
- **Added `excel_setup.py`**: Auto-populates the AST spreadsheet

### April 2025

#### April 15, 2025 - Development Tools Enhancement
- **Added TODO Tree settings** and examples for knowledge sharing
- **Workspace configuration updates** and file cleanup

### March 2025

#### March 17-20, 2025 - BCGW Connection Overhaul (Major Refactoring)
- **Simplified database connection handling**:
  - Created BCGW (.sde) connection only once in `main.py`
  - Passed connection path via environment variable (`SDE_FILE_PATH`)
  - Updated `automated_status_sheet_call_routine_arcpro.py`, `one_status_tabs_one_and_two_arcpro.py`, `universal_overlap_tool_call_routine_arcpro.py`, and `mp_worker.py` to reference environment variable
  - Added validation checks for .sde connection path before use
- **Tested V1, V2, and V3**: All working with new BCGW connection handling
- **Added `clear_BCGW_Keychain_Credentials.py`** script for credential management
- **Added BatchFactory documentation**
- **Created BCGW Edit Version of V3**: Tested and working

### February 2025

#### February 6-19, 2025 - V2 Foundation and Multi-Version Testing
- **Merged `ast_Alpha_modularization_v2` branch into main** — the foundational merge
  - Established V2 "Cuisinart" architecture with multiprocessing
  - Modularized codebase for better maintainability
  - PDF map generation capabilities
- **Tested V1, V2, and V3**: All tested and working
  - Changed TOOLBOX in V3 to `TOOLBOX_ALPHA` in `.env` file
- **Also added**: RepathLayerTool.pyt for Cartographic Standards Group, KFN Watermap

### January 2025

#### January 23-24, 2025 - V2 Folder Organization
- **Organized folder structure** to include V2 — tested and working
- **Cleaned up transient files** prior to merge with main
- **Merged `ast_Alpha_modularization_v2` into main**: Created backup before merging

#### January 6-10, 2025 - Batch Factory Creation
- **Started creating generic Batch Factory** (`batch_factory.py`)
  - New approach: read header and create list, then read header and rows, add to dictionary, then batch the job
  - Created `main.py` entry point
  - Moved Batch Factory up one directory in folder tree
  - Multiple iterations to resolve import and toolbox connection issues
- **Tested and working** by January 10

#### January 2-3, 2025 - Alpha Toolbox Migration
- **Converted .tbx to .atbx** using ArcGIS Pro for better compatibility
- **Alpha 4 Testing**: Redirected call sheet routine to Steve's alpha folder — confirmed scripts don't need to live in auto_ast folder
- **Recreated Alpha Statusing .atbx**: Renamed work `ast.atbx` back to alpha
- **Restored MAIN branch** to working condition
- **Cleaned up imports** in all submodules

### December 2024

#### December 27-29, 2024 - Alpha Testing and Recovery
- **First commit on Alpha 3**: Got system back to a working state
- **Retested**: Working on `\Ready` and `ast.atbx`

#### December 16-17, 2024 - Modularization Milestone (Major Architecture Change)
- **Successfully modularized the AST Factory** into separate modules:
  - `toolbox_import.py` - Toolbox import utilities
  - `mp_worker.py` - Multiprocessing worker handler
  - `database_connection.py` - Database connections
  - `logging_setup.py` - Logging configuration
- **Created V2 of Fish and Wildlife toolbox** for future expansion
- **Created New Fish and Wildlife toolbox**: Plans to hold all F&W and possibly Lands & Water tools

#### December 13, 2024 - Completed Jobs Fix
- **Merged `FixCompletedJobsRunningTwice` branch into main**
- **Dev version tested**: Works unless more than 4 maps are created
- **Added Excel file name to log folder name** for easier tracking

#### December 10-11, 2024 - Cariboo Testing
- **Ran Cariboo Replacement Jobs**: All completed in one pass on Quick Settings
- **Identified issue**: Script says it's changing "don't overwrite outputs" to true but doesn't change them

### November 2024

#### November 19-25, 2024 - V2 Cuisinart Development and Indexing Resolution
- **After extensive testing, the indexing issue may be resolved**
  - Merging back to main after hard-copy backup
- **Created V2 "Cuisinart" development version** (`ast_v2_Cuisinart_DEV.py`)
  - Multiple iterations of backup, testing, and refinement
- **Still tracking indexing issue**: Autoast_Indexing on main completes a job then completes it a second time
- **Key fix attempts**: Moved skip-if-complete and failed logic, adjusted loop indentation

#### November 13-18, 2024 - Output Directory and Excel Improvements
- **Re-added handling** to create output directory if one doesn't exist (defaults to T: drive)
- **Excel results improving**: Updated indexing and commented out problematic "continue" statement
- **Test of 4 jobs designed to fail**: Ran successfully after fixes

#### November 1-6, 2024 - Logging and Indexing Fixes
- **Updated commenting and logging extensively** — cleaner and more descriptive
- **Refined logic** for handling True/False/blank condition of "Don't Overwrite Outputs"
- **Updated Load Jobs** with new indexing fix
- **Added `add_job_result`** to reload failed jobs

### October 2024

#### October 28-31, 2024 - Index Normalization
- **Changed all `job_index`** to remove -1 or +1 adjustments
- **Making progress**: Worker processes failing quickly due to exceptions, added `import arcpy` to worker
- **Cleaned up extraneous files**

#### October 16-22, 2024 - Deep Indexing Investigation
- **Discovered**: `excel_row_index = job_index + 2` must stay at +2
- **Multiple test iterations** with start=1 vs start=2 in reload jobs
- **Changed `index from +1 to +2`** in add_job_result line 502
- **Changed "completed" to "Success"**: Jobs now running properly

#### October 7-10, 2024 - V1/V2 Reconciliation (Major Debugging)
- **Major Changes and Cleanup**: Recovered AutoAst V1 (pre-multiprocessing/timeout) as best known working version
- **Reverted to V1 with working indexing**, then step-by-step moved V2 changes in:
  - All new BatchAST functions, new call routine
  - Changed `logger.info` to `self.logger.info`
  - All new `reload_failed_jobs` (referred to as RLFJ_V2)
- **Added search BCGW for keyword script**

#### October 1-3, 2024 - Batch_ast_v2 Branch Merge
- **Merged `Batch_ast_v2` into main**: Timeout now working, logging working on worker processes
- **Added reload failed jobs function**: Tested on sheet with all COMPLETE, did fail as expected
- **AST condition** being changed to requeued on failed jobs, don't overwrite outputs set to true

### September 2024

#### September 16-24, 2024 - Timeout and Multiprocessing Implementation
- **Timeout now timing out and WORKING**: Created copy of logging before adding `process_mp` function
- **Experimenting with V1 and V2**: V1 timing accurately and producing failed error when exceeding timeout
- **Added Draw.io diagram flowchart** documenting batch processing flow
- **Logging fixed**: Worker logs correctly capturing arcpy messages for each worker job
- **Added map automation tools** and field calculator scripts for others to use

#### September 11-17, 2024 - Multiprocessing Begins
- **Created draft copies of multiprocessing**: Bare-bones main and worker, both worked
- **Batch AST V2 seems to be working**: Had to change logging into worker function, fix index issue, change from global variables to function calls
- **Added logger** to AST Factory class
- **Fixed error in classifying input type**: Script working but logging needs tweaks

#### September 4-9, 2024 - Failed Jobs and Threading Attempts
- **Merged `rerun_failed_jobs` branch into main**
- **Fixed error** where completed wasn't being updated on Excel sheet
- **Added job threading** — didn't work; discovered ArcPy doesn't support job threading
- **Activated `reload_failed_jobs`** and `rebatch_failed_jobs` (not yet fully implemented)
- **Removed duplicate for loop** from classify input type; added handling for blank rows
- **Fixed issue** where jobs marked COMPLETE were rerunning

#### September 4, 2024 - LAST KNOWN GOOD Checkpoint
- **Stable checkpoint** established before timeout and threading experiments

### August 2024

#### August 22-30, 2024 - Error Handling and Job Management
- **Added error handling** to `load_jobs()`: immediate updating of Excel sheet before `batch_jobs` runs
- **Fixed Arcpy errors**: ERROR 000622 (Parameters not valid) and ERROR 000628 (Cannot set input into DO_NOT_USE_Debug)
- **Moved AST Condition** from AST Parameters to Additional Parameters
- **Added re-run jobs function**: Compared AST edits, fixed rerun failed jobs
- **Moved .shp and kml check** to `classify_input_type`

#### August 7-9, 2024 - Logging and Environment Configuration
- **Added basic Logging** to the script
- **Moved all file paths to .env file**: Contact developer for new `.env` file
- **Moved Excel file path** to top of script for team members to easily find and replace
- **Added `capture_arcpy_messages` method** to capture messages from AST toolbox
- **Merged `capture_asttoolbox_arcpy_messages` branch into main**
- **Ran overnight batch** using Sunny's 5 WMU Excel sheet

#### August 1, 2024 - Initial Logging
- **Initial script for logging** created

### July 2024 - Project Inception

#### July 23-30, 2024 - Shapefile Handling and FW Setup (Major Feature)
- **KML AOI working**: Fixed KML handling for Area of Interest
- **Added Shapefile handling**: Tested and working on july_23.xlsx sheet
- **Merged `cs_fix_kml` branch into main**
- **Added Mike's FW Setup tool** to create AOI from shapefiles
- **Merged FW Setup branch into main**: Runs FW Setup on spreadsheets containing file numbers
- **Updated README.md**: First documentation update

#### July 8-15, 2024 - Project Creation (First Commits)
- **Project inception**: First commits to create the automated AST batch processing tool
- **Successfully imported AST toolbox** and test-printed file path
- **Script completed without errors**: Got through BCGW Connect, DOTENV files; AOI.KML initially failing
- **Resolved issue** of importing and executing `ast_call_routine` — tested on simple test data and working
- **Added function** to run from toolbox or from script for testing purposes
- **`ast.start_ast_tb(jobs)` ran successfully** and produced outputs — first successful end-to-end run
- **Merged `csostadStartAst` branch into main**
- **Added test Excel files**: `.aoi` kml not yet working, but .xlsx with known working jobs confirmed




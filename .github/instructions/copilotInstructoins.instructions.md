---
description: Always load these instructions for all tasks in this repository.
applyTo: '**'
---

# AutoAST Batch Tool — AI Assistant Instructions
# BC Ministry of Forests, Lands and Natural Resource Operations

---

## 1. Project Overview

AutoAST is a Python batch processing framework that orchestrates the **Automated Status Tool (AST)** for Crown land applications in British Columbia. It spawns worker subprocesses; each worker imports `ast.atbx` and calls the `MakeAutomatedStatusSpreadsheet` tool inside it.

- The active version is `auto_ast_V2_Cuisinart_MultiP_PdfMaps/`
- The developer owns the **orchestration and batch code only**
- `ast.atbx` is maintained by a separate team — it is a hard external boundary
- The next development phase is adapting this script to call a new updated version of the AST tool

---

## 2. Change Discipline

- Make **surgical changes only** — the minimum code necessary to achieve the stated goal
- Never refactor, rename, reorganize, or "improve" anything not directly related to the task — ask first
- **Comment every new line or block of code** added, explaining what it does and why
- If a second issue is spotted while working, annotate it with `# BUG` and flag it — do not fix it silently
- One concern per change — stop and report before moving on

---

## 3. Bug & Accuracy Vigilance

This is government work. Accuracy is critical — incorrect outputs can affect legal Crown land decisions.

Actively scan for:
- Off-by-one errors in sequential logic
- Silent exception swallowing (`except: pass` or bare `except`)
- Temp feature classes or files not cleaned up between runs
- Wrong field used as a job identifier or join key

When a potential bug is found, annotate it with `# BUG - [CRITICAL|HIGH|MEDIUM|LOW] - description` and report it before touching anything else.

---

## 4. Toolbox Boundary (Hard Rule)

`ast.atbx` is **never modified under any circumstances**. It is external code owned by another team.

When adapting the tool to call a new version of the AST tool, changes are limited strictly to:
- The toolbox path and alias in `.env`
- Parameters passed in `mp_worker.py`
- Pre/post-processing logic in `ast_factory.py`

If the new tool version requires a different parameter signature, document the difference and ask before adapting.

---

## 5. Environment & Credentials

- Always use `os.getenv()` to read configuration — never hardcode paths or credentials
- All sensitive values live in `.env` (not committed to the repository)
- Required `.env` keys: `BCGW_USER`, `BCGW_PASS`, `TOOLBOX`, `TOOLBOXALIAS`, `TEMPLATE`, `SECRET_FILE`, `SDE_FILE_PATH`
- The SDE connection is created **once in the main process** and shared via the `SDE_FILE_PATH` environment variable — do not create per-worker connections

---

## 6. Multiprocessing Rules

- Always use `multiprocessing.Process` — never `threading` for ArcPy work
- Worker results are returned via a `Manager` dictionary — never raise exceptions across a process boundary
- Job timeout: **6 hours**; health check interval: **60 seconds**
- Hard limit: **8 jobs per Excel file** — this is a BCGW concurrent connection limit, not a preference

---

## 7. Excel Job File

- `ast_condition` column controls the job lifecycle: `Queued` → `COMPLETE` / `Failed` / `Requeued` / `FAILED_OUTPUTS`
- A job must **never** be marked `COMPLETE` by any means other than validated output from `output_validator.py`
- `job_index` (0-based) + 2 = Excel row number (row 1 is the header row)

---

## 8. Python & Environment Standards

- Target: **Python 3 / ArcGIS Pro 3 / arcpy (Pro)**
- Use f-strings and `pathlib.Path` for all new path handling
- Use `arcpy.AddMessage()` in any toolbox-facing code; use the `logging` module everywhere else
- Do not use Python 2 syntax, `arcpy.mapping`, `win32com`, or `gp.searchcursor`

---

## 9. Output Validation Boundary

`output_validator.py` defines what a successful job looks like: 3 GDBs, 3 folders (maps, mapx_files), and 3 Excel files — all non-empty. This is the single source of truth for job completion. No other mechanism may mark a job `COMPLETE`.

---

## 10. Log File Guidance

When debugging, check these locations first:
- `autoast_batch_logs_folder_YYYYMMDD/` — main orchestration log (batch-level events, job state changes)
- `autoast_worker_logs_YYYYMMDD/` — per-process logs (ArcPy messages, toolbox errors, worker crashes)

Worker logs are the primary diagnostic tool for ArcPy-level failures.

---

## 11. No Silent Defaults

If a required `.env` key or configuration value is missing, raise a clear and descriptive error immediately. Never substitute a fallback value, default path, or silent workaround — missing config must be visible and explicit.
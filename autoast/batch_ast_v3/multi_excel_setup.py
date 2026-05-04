'''
This script will prep a folder of shapefiles to be run through the batch tool. 
Due to limitations with BCGW connections and multiprocessing, any excel file greater than 8 jobs MAY
cause an issue with the BCGW and too many connections. This script will break up your list of shapefiles
into excel files with a maximum of 8 jobs per file.

# Its still early days but it can populate your excel file with the region. If you need a variety of regions,
just choose a region and then you will have to edit it by hand. A pain in the ass for now but to be improved later.

Opens a window asking the user to choose a main directory where the input data is stored.
Shows a dropdown menu asking the user to select a region (e.g., "Northeast", "Skeena", etc.).
The script goes through each subfolder inside the selected directory. 
If it finds a .shp (shapefile) inside a subfolder, it:
Notes the shapefile's path
Creates a matching output folder
Writes an entry to an Excel sheet with default settings
It adds up to 8 jobs per Excel workbook. If more than 8 shapefiles are found, it starts a new workbook.
Each Excel file is saved in a folder called outputs within the selected main directory.
'''

import os
import logging
from pathlib import Path
from openpyxl import Workbook
from tkinter import Tk, filedialog, Label, Button, StringVar, OptionMenu, Entry, Frame
from dotenv import load_dotenv


# Sentinel value used to trigger auto region detection via BCGW SpatialJoin.
# Must match exactly what appears in the tkinter dropdown list below.
AUTO_DETECT_OPTION = "Detect automatically (regions vary)"


def _collect_shapefiles(main_dir, output_dir):
    """
    Scan main_dir for shapefiles and return a list of (shp_path, output_subfolder) tuples.

    Search strategy (in order):
      1. Look inside each immediate subfolder of main_dir for a .shp file.
         One shapefile per subfolder is expected; the first found is used.
      2. If no shapefiles are found in any subfolder, fall back to scanning
         main_dir itself for .shp files directly.

    This single collection pass is shared by both the manual-region and
    auto-detect code paths, so shapefile discovery logic is not duplicated.

    Parameters:
        main_dir   : Path  — directory selected by the user
        output_dir : Path  — outputs/ directory (already created by caller)

    Returns:
        list of (Path, Path) — (shapefile path, output subfolder path)
    """
    jobs = []

    # Pass 1: subfolders of main_dir
    for entry in sorted(os.listdir(main_dir)):
        subfolder_path = main_dir / entry

        # Skip the outputs folder created by this script
        if entry == "outputs":
            continue

        if subfolder_path.is_dir():
            # Take the first .shp found in this subfolder
            shapefile = next((f for f in os.listdir(subfolder_path) if f.endswith(".shp")), None)
            if shapefile:
                shp_path = subfolder_path / shapefile
                output_subfolder = output_dir / entry
                output_subfolder.mkdir(parents=True, exist_ok=True)
                jobs.append((shp_path, output_subfolder))

    # Pass 2: if no subfolders contained shapefiles, check main_dir directly
    if not jobs:
        for filename in sorted(os.listdir(main_dir)):
            if filename.endswith(".shp"):
                shp_path = main_dir / filename
                # All flat-dir shapefiles share the same output subfolder (named after main_dir)
                output_subfolder = output_dir / main_dir.name
                output_subfolder.mkdir(parents=True, exist_ok=True)
                jobs.append((shp_path, output_subfolder))

    return jobs


def create_job_excel_files():
    """
    Creates Excel job files for the AST Batch tool.

    Prompts the user via a tkinter GUI to:
      - Select the input directory (flat folder of .shp files, or folder of subfolders each
        containing a .shp file)
      - Choose a region (or select 'Detect automatically (regions vary)' to use BCGW spatial
        analysis to determine the NR Region for each shapefile automatically)
      - Optionally choose a custom log directory

    Shapefiles are grouped into Excel workbooks of max 8 rows each (BCGW concurrent
    connection limit). Each row contains the full set of parameters for one AST batch job.

    Returns:
        tuple (list[str], str)
            - List of absolute paths to the created Excel files (fed directly to main_auto_setup.py)
            - Custom log directory path string (empty string = use default)
    """
    print("Running Create Job Excel Files")

    # ------------------------------------------------------------------ #
    # GUI — directory selection                                            #
    # ------------------------------------------------------------------ #
    root = Tk()
    root.withdraw()  # Hide root during folder dialog

    main_dir = filedialog.askdirectory(title="Select the main directory containing your shapefiles")
    if not main_dir:
        print("No directory selected. Exiting.")
        root.destroy()
        exit()
    print(f"Selected directory: {main_dir}")
    main_dir = Path(main_dir)

    # ------------------------------------------------------------------ #
    # GUI — region selection and log path                                  #
    # ------------------------------------------------------------------ #
    root.deiconify()
    root.title("AutoAST Batch Setup")
    root.geometry("650x380")

    region_var = StringVar(root)
    region_var.set("Northeast")  # Default selection

    Label(root, text="Select a region (or choose auto-detect if your AOIs span multiple regions):").pack(pady=10)

    # "Detect automatically" is the last option; selecting it triggers BCGW spatial analysis
    regions = [
        "Northeast", "Cariboo", "Kootenay_Boundary", "Skeena",
        "South_Coast", "Thompson_Okanagan", "West_Coast", "Omineca",
        AUTO_DETECT_OPTION
    ]
    OptionMenu(root, region_var, *regions).pack(pady=10)

    Label(root, text=(
        "Select a directory for log files (optional)\n"
        "Leave blank to write logs in the default location (script directory)\n"
        "Ensure you have write access to the selected path"
    )).pack(pady=10)

    log_frame = Frame(root)
    log_frame.pack(pady=5)
    log_path_var = StringVar(root)
    Entry(log_frame, textvariable=log_path_var, width=60).pack(side="left", padx=5)

    def browse_log_path():
        selected_dir = filedialog.askdirectory(title="Select directory for log files")
        if selected_dir:
            log_path_var.set(selected_dir)

    Button(log_frame, text="Browse...", command=browse_log_path).pack(side="left")

    def confirm_selection():
        root.quit()
        root.destroy()

    Button(root, text="Confirm", command=confirm_selection).pack(pady=20)
    root.mainloop()

    selected_region = region_var.get()
    log_path = log_path_var.get()
    print(f"Selected region: {selected_region}")
    print(f"Log path: {log_path if log_path else 'Default location'}")

    # ------------------------------------------------------------------ #
    # Collect shapefiles (both flat-folder and subfolder scenarios)        #
    # ------------------------------------------------------------------ #
    output_dir = main_dir / "outputs"
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output directory: {output_dir}")

    jobs = _collect_shapefiles(main_dir, output_dir)
    print(f"\nCollected {len(jobs)} shapefile(s) for processing")

    if not jobs:
        print("\nWARNING: No shapefiles found in subfolders OR main directory!")
        print("  1. Ensure the directory contains .shp files (in subfolders or directly)")
        print(f"  2. Selected directory was: {main_dir}")
        return [], log_path

    # ------------------------------------------------------------------ #
    # Region detection — auto-detect via BCGW, or use manual selection    #
    # ------------------------------------------------------------------ #
    if selected_region == AUTO_DETECT_OPTION:
        print("\nAuto-detect selected — connecting to BCGW to determine NR Regions...")

        # Credentials live in .env; load them here before calling setup_bcgw.
        # setup_bcgw requires a logger; we use a console-only logger since the
        # main log file has not been created yet (that happens after this function returns).
        load_dotenv()
        _temp_logger = logging.getLogger("region_setup")
        if not _temp_logger.handlers:
            _temp_logger.addHandler(logging.StreamHandler())
            _temp_logger.setLevel(logging.INFO)

        # Import here (not at module top) to keep this optional dependency lazy —
        # users who never use auto-detect don't need arcpy available at import time.
        from database_connection import setup_bcgw  # noqa: PLC0415
        from region_detector import detect_regions   # noqa: PLC0415

        _, _, sde_path = setup_bcgw(_temp_logger)
        print(f"BCGW connection established at: {sde_path}")

        shp_paths = [shp for shp, _ in jobs]
        region_map = detect_regions(shp_paths, sde_path, _temp_logger)

        # Warn for any shapefile that came back with an empty region
        for i, (shp_path, _) in enumerate(jobs):
            if not region_map.get(i):
                print(f"WARNING: Could not detect region for '{shp_path.name}' — cell will be blank")
    else:
        # Manual selection: apply the same region to every row; no BCGW call needed.
        region_map = {i: selected_region for i in range(len(jobs))}
        print(f"Using manual region '{selected_region}' for all {len(jobs)} shapefile(s)")

    # ------------------------------------------------------------------ #
    # Excel workbook writing                                               #
    # ------------------------------------------------------------------ #
    # Updated column names to match fcbc_auto_status_tool.pyt parameter index map
    headers = [
        "region", "feature_layer", "crown_file_number", "disposition_number",
        "parcel_number", "output_directory", "output_dir_same_as_input",
        "dont_overwrite_outputs", "dont_run_conflicts_and_constraints_tab3",
        "suppress_map_creation_tab3", "open_output_directory", "full_path_hyperlinks",
        "ast_condition", "file_number"
    ]

    excel_files = []
    workbook_counter = 1
    job_counter = 0

    wb = Workbook()
    ws = wb.active
    ws.title = "ast_config"
    ws.append(headers)

    for i, (shp_path, output_subfolder) in enumerate(jobs):
        region = region_map.get(i, "")  # Empty string if detection failed for this shapefile
        ws.append([
            region,
            str(shp_path),
            "",   # crown_file_number — blank; user fills if needed
            "",   # disposition_number
            "",   # parcel_number
            str(output_subfolder),
            "false",  # output_dir_same_as_input
            "false",  # dont_overwrite_outputs
            "false",  # dont_run_conflicts_and_constraints_tab3
            "false",  # suppress_map_creation_tab3
            "false",  # open_output_directory
            "true",   # full_path_hyperlinks
            "",       # ast_condition — blank; batch tool sets this at runtime
            ""        # file_number
        ])
        job_counter += 1
        print(f"Job {i + 1} added: {shp_path.name} -> region='{region}'")

        # Split into a new workbook every 8 jobs (BCGW concurrent connection limit)
        if job_counter == 8:
            batch_excel_path = output_dir / f"jobs_{workbook_counter}.xlsx"
            wb.save(str(batch_excel_path))
            excel_files.append(str(batch_excel_path))
            print(f"Saved: {batch_excel_path}")
            workbook_counter += 1
            job_counter = 0
            wb = Workbook()
            ws = wb.active
            ws.title = "ast_config"
            ws.append(headers)

    # Save any remaining rows that didn't fill a full workbook of 8
    if job_counter > 0:
        batch_excel_path = output_dir / f"jobs_{workbook_counter}.xlsx"
        wb.save(str(batch_excel_path))
        excel_files.append(str(batch_excel_path))
        print(f"Saved: {batch_excel_path}")

    print(f"\nTotal shapefiles processed: {len(jobs)}")
    print(f"Total Excel files created: {len(excel_files)}")

    return excel_files, log_path

if __name__ == "__main__":
    excel_files, log_path = create_job_excel_files()
    print(f"List of excel file paths is: {excel_files}")
    print(f"Log path: {log_path if log_path else 'Default location'}")

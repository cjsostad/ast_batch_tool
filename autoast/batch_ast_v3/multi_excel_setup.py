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
from pathlib import Path
from openpyxl import Workbook
from tkinter import Tk, filedialog, Label, Button, StringVar, OptionMenu, Entry, Frame

def create_job_excel_files():
    '''
    Creates Excel files for the AST Batch tool. Creates the Excel files in jobs of 8 to workaround the max allowed BCGW connections.
    Also prompts the user and populates the region.
    '''
    
    print("Running Create Job Excel Files")
    
    # Initialize Tkinter root window
    root = Tk()
    root.withdraw()  # Hide the main window during file dialog

    # Prompt user to select the main directory that contains the shapefiles
    main_dir = filedialog.askdirectory(title="Select the main directory ")
    if not main_dir:
        print("No directory selected. Exiting.")
        root.destroy()
        exit()
    print(f"Selected directory: {main_dir}")
    main_dir = Path(main_dir)

    # Create a dropdown menu for region selection
    root.deiconify()
    print("Creating a dropdown menu for region selection...")
    root.title("Select Region and Log Path")
    root.geometry("650x350")  # Make window larger to fit all elements
    region_var = StringVar(root)
    region_var.set("Northeast")  # Default region

    Label(root, text="Select a region:").pack(pady=10)

    regions = [
        "Northeast", "Cariboo", "Kootenay_Boundary", "Skeena",
        "South_Coast", "Thompson_Okanagan", "West_Coast", "Omineca"
    ]

    OptionMenu(root, region_var, *regions).pack(pady=10)

    Label(root, text="Please select a directory for your log files\nLeave blank to write logs in the default location (script directory)\nMake sure you have write access to the selected path").pack(pady=10)
    
    # Create a frame for log path entry and browse button
    log_frame = Frame(root)
    log_frame.pack(pady=5)
    
    log_path_var = StringVar(root)
    Entry(log_frame, textvariable=log_path_var, width=60).pack(side="left", padx=5)
    
    def browse_log_path():
        """Open a directory browser dialog to select log path"""
        selected_dir = filedialog.askdirectory(title="Select directory for log files")
        if selected_dir:
            log_path_var.set(selected_dir)
            print(f"Selected log path: {selected_dir}")
    
    Button(log_frame, text="Browse...", command=browse_log_path).pack(side="left")

    def confirm_selection():
        root.quit()
        root.destroy()

    Button(root, text="Confirm", command=confirm_selection).pack(pady=20)
    root.mainloop()
    region = region_var.get()
    log_path = log_path_var.get()
    print("Selected region:", region)
    print("Log path:", log_path if log_path else "Default location")

    # Define paths for the output directory
    output_dir = main_dir / "outputs"
    os.makedirs(output_dir, exist_ok=True)
    print(f"Output directory created: {output_dir}")

    # Initialize counters and file list
    job_counter = 0
    workbook_counter = 1
    excel_files = []

    # Create the first workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "ast_config"

    # version2 - old column names for ast.atbx / alphaast toolbox
    # headers = [
    #     "region", "feature_layer", "crown_file_number", "disposition_number",
    #     "parcel_number", "output_directory", "output_directory_same_as_input",
    #     "dont_overwrite_outputs", "skip_conflicts_and_constraints",
    #     "suppress_map_creation", "add_maps_to_current", "run_as_fcbc",
    #     "ast_condition", "file_number"
    # ]

    # Updated column names to match fcbc_auto_status_tool.pyt (AutomatedStatusTool) parameter index map
    headers = [
        "region", "feature_layer", "crown_file_number", "disposition_number",
        "parcel_number", "output_directory", "output_dir_same_as_input",
        "dont_overwrite_outputs", "dont_run_conflicts_and_constraints_tab3",
        "suppress_map_creation_tab3", "open_output_directory", "full_path_hyperlinks",
        "ast_condition", "file_number"
    ]
    ws.append(headers)
    print("Headers added to the workbook.")

    # Iterate through subfolders in the main directory looking for shapefiles
    print(f"\nSearching for shapefiles in subfolders of: {main_dir}")
    subfolders_found = 0
    shapefiles_found = 0
    
    for subfolder in os.listdir(main_dir):
        subfolder_path = main_dir / subfolder

        # Skip the outputs folder (created by this script)
        if subfolder == "outputs":
            print(f"Skipping outputs folder: {subfolder_path}")
            continue

        if subfolder_path.is_dir():
            subfolders_found += 1
            print(f"Processing subfolder: {subfolder_path}")
            shapefile = next((f for f in os.listdir(subfolder_path) if f.endswith(".shp")), None)
            if shapefile:
                shapefiles_found += 1
                shapefile_path = subfolder_path / shapefile
                print(f"Found shapefile: {shapefile_path}")

                output_subfolder = output_dir / subfolder
                os.makedirs(output_subfolder, exist_ok=True)
                print(f"Output subfolder created: {output_subfolder}")

                # Append job details to the worksheet using Windows-style backslashes
                ws.append([
                    region,
                    str(shapefile_path),
                    "",
                    "",
                    "",
                    str(output_subfolder),
                    "false",
                    "false",
                    "false",
                    "false",
                    "false",
                    "true",
                    "",
                    ""
                ])

                job_counter += 1
                print(f"Job added. Current job counter: {job_counter}")

                if job_counter == 8:
                    batch_excel_path = output_dir / f"jobs_{workbook_counter}.xlsx"
                    wb.save(batch_excel_path)
                    excel_files.append(str(batch_excel_path))
                    print(f"Batch jobs template saved to: {batch_excel_path}")
                    workbook_counter += 1
                    job_counter = 0

                    wb = Workbook()
                    ws = wb.active
                    ws.title = "ast_config"
                    ws.append(headers)  # headers already updated to new column names above
                    print("New workbook created.")
            else:
                print(f"No shapefile found in: {subfolder_path}")
    
    print(f"\nSummary: Found {subfolders_found} subfolder(s), {shapefiles_found} shapefile(s)")
    
    # If no shapefiles found in subfolders, check the main directory itself
    if shapefiles_found == 0:
        print(f"\nNo shapefiles found in subfolders. Checking main directory: {main_dir}")
        main_dir_shapefiles = [f for f in os.listdir(main_dir) if f.endswith(".shp")]
        
        if main_dir_shapefiles:
            print(f"Found {len(main_dir_shapefiles)} shapefile(s) in main directory")
            for shapefile in main_dir_shapefiles:
                shapefiles_found += 1
                shapefile_path = main_dir / shapefile
                print(f"Found shapefile: {shapefile_path}")
                
                # Use the main directory name for output subfolder
                main_dir_name = main_dir.name
                output_subfolder = output_dir / main_dir_name
                os.makedirs(output_subfolder, exist_ok=True)
                print(f"Output subfolder created: {output_subfolder}")
                
                # Append job details to the worksheet
                ws.append([
                    region,
                    str(shapefile_path),
                    "",
                    "",
                    "",
                    str(output_subfolder),
                    "false",
                    "false",
                    "false",
                    "false",
                    "false",
                    "true",
                    "",
                    ""
                ])
                
                job_counter += 1
                print(f"Job added from main directory. Current job counter: {job_counter}")
                
                if job_counter == 8:
                    batch_excel_path = output_dir / f"jobs_{workbook_counter}.xlsx"
                    wb.save(batch_excel_path)
                    excel_files.append(str(batch_excel_path))
                    print(f"Batch jobs template saved to: {batch_excel_path}")
                    workbook_counter += 1
                    job_counter = 0

                    wb = Workbook()
                    ws = wb.active
                    ws.title = "ast_config"
                    ws.append(headers)
                    print("New workbook created.")
        else:
            print("\nWARNING: No shapefiles found in subfolders OR main directory!")
            print("Make sure:")
            print("  1. Your directory contains .shp files (either in subfolders or directly)")
            print("  2. You selected the correct directory")
            print(f"\nSelected directory was: {main_dir}")

    if job_counter > 0:
        batch_excel_path = output_dir / f"jobs_{workbook_counter}.xlsx"
        wb.save(batch_excel_path)
        excel_files.append(str(batch_excel_path))
        print(f"Batch jobs template saved to: {batch_excel_path}")
    
    print(f"\nTotal shapefiles processed: {shapefiles_found}")
    print(f"Total Excel files created: {len(excel_files)}")

    return excel_files, log_path

if __name__ == "__main__":
    excel_files, log_path = create_job_excel_files()
    print(f"List of excel file paths is: {excel_files}")
    print(f"Log path: {log_path if log_path else 'Default location'}")

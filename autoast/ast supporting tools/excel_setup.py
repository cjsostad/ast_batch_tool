import os
from openpyxl import Workbook
from tkinter import Tk
from tkinter.filedialog import askdirectory

def populate_jobs_template():
    # Prompt user to select the main directory
    Tk().withdraw()  # Hide the root Tkinter window
    main_dir = askdirectory(title="Select the main directory provided by the client")
    if not main_dir:
        print("No directory selected. Exiting.")
        return

    # Define paths for the output directory and the new Excel file
    output_dir = os.path.join(main_dir, "outputs")
    os.makedirs(output_dir, exist_ok=True)
    batch_excel_path = os.path.join(main_dir, "jobs.xlsx")

    # Create a new workbook and set the sheet name to "ast_config"
    wb = Workbook()
    ws = wb.active
    ws.title = "ast_config"

    # Add the required headers
    headers = [
        "region", "feature_layer", "crown_file_number", "disposition_number", 
        "parcel_number", "output_directory", "output_directory_same_as_input", 
        "dont_overwrite_outputs", "skip_conflicts_and_constraints", 
        "suppress_map_creation", "add_maps_to_current", "run_as_fcbc", 
        "ast_condition", "file_number"
    ]
    ws.append(headers)

    # Iterate through subfolders in the main directory
    for subfolder in os.listdir(main_dir):
        subfolder_path = os.path.join(main_dir, subfolder)
        if os.path.isdir(subfolder_path):
            # Look for a shapefile in the subfolder
            shapefile = next((f for f in os.listdir(subfolder_path) if f.endswith(".shp")), None)
            if shapefile:
                shapefile_path = os.path.join(subfolder_path, shapefile)

                # Create a corresponding output subfolder
                output_subfolder = os.path.join(output_dir, subfolder)
                os.makedirs(output_subfolder, exist_ok=True)

                # Populate the Excel sheet with the shapefile path and default values
                ws.append([
                    "",  # region
                    shapefile_path,  # feature_layer
                    "",  # crown_file_number
                    "",  # disposition_number
                    "",  # parcel_number
                    output_subfolder,  # output_directory
                    "false",  # output_directory_same_as_input
                    "false",  # dont_overwrite_outputs
                    "false",  # skip_conflicts_and_constraints
                    "false",  # suppress_map_creation
                    "false",  # add_maps_to_current
                    "true",  # run_as_fcbc
                    "",  # ast_condition
                    ""  # file_number
                ])

    # Save the populated Excel sheet
    wb.save(batch_excel_path)
    print(f"Batch jobs template saved to: {batch_excel_path}")

if __name__ == "__main__":
    populate_jobs_template()
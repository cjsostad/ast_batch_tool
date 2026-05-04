"""
Output Validator Module
Validates that all required output files and folders exist and are not empty.
Used by both ast_factory.py and main_finish_jobs.py to ensure consistency.
"""

import os


def verify_job_outputs(output_directory, logger=None):
    '''
    Verifies that all required output files and folders exist and are not empty.
    
    Args:
        output_directory: Path to the job's output directory
        logger: Logger instance (optional)
    
    Returns:
        (is_valid, missing_items) where is_valid is True if all outputs exist and are valid,
        and missing_items is a list of missing/invalid items
    '''
    # Required GDB and folder outputs from the new auto_status runner
    # version2: removed mapx_files (not created by new tool)
    required_items = {
        'aoi_boundary.gdb': 'folder',
        'maps': 'folder',
        'one_status_common_datasets_aoi.gdb': 'folder',
        'one_status_tabs_1_and_2_datasets.gdb': 'folder',
    }
    
    missing_items = []

    # Check if output directory exists
    if not os.path.exists(output_directory):
        return (False, ['output_directory'])

    # The new runner (auto_status) creates a YYYYMMDD-Status-{suffix} subdirectory inside
    # the job's output_directory. Scan for it and validate inside it.
    # version2: folder name pattern is YYYYMMDD-Status-{suffix}, not Status_{suffix}
    subdirs = [
        d for d in os.listdir(output_directory)
        if os.path.isdir(os.path.join(output_directory, d))
        and ('-Status-' in d or '-DEBUG-' in d)
    ]
    if not subdirs:
        # No YYYYMMDD-Status- run subdirectory found — outputs are missing
        return (False, ['YYYYMMDD-Status- run subdirectory (missing)'])
    # Use the most recently modified Status_ subfolder
    subdirs.sort(
        key=lambda d: os.path.getmtime(os.path.join(output_directory, d)),
        reverse=True
    )
    check_directory = os.path.join(output_directory, subdirs[0])
    if logger:
        logger.info(f"Output Validator: Resolved run directory to {check_directory}")

    # Check each required item inside the Status_ run subdirectory
    for item_name, item_type in required_items.items():
        item_path = os.path.join(check_directory, item_name)

        if not os.path.exists(item_path):
            missing_items.append(f"{item_name} (missing)")
            continue
        
        if item_type == 'folder':
            # Check if folder is not empty
            try:
                if not os.listdir(item_path):
                    missing_items.append(f"{item_name} (empty folder)")
            except Exception as e:
                missing_items.append(f"{item_name} (error: {e})")
        
        elif item_type == 'file':
            # Check if file is not empty (size > 0)
            try:
                if os.path.getsize(item_path) == 0:
                    missing_items.append(f"{item_name} (empty file)")
            except Exception as e:
                missing_items.append(f"{item_name} (error: {e})")
    
    # The new runner creates a single merged xlsx named after the run folder (e.g. 20260504-Status-632.xlsx)
    # version2: check for any non-empty .xlsx file instead of fixed names
    xlsx_files = [
        f for f in os.listdir(check_directory)
        if f.lower().endswith('.xlsx')
        and os.path.isfile(os.path.join(check_directory, f))
        and os.path.getsize(os.path.join(check_directory, f)) > 0
    ]
    if not xlsx_files:
        missing_items.append('output .xlsx (missing or empty)')

    return (len(missing_items) == 0, missing_items)

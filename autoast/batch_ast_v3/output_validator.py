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
    required_items = {
        'aoi_boundary.gdb': 'folder',
        'maps': 'folder',
        'mapx_files': 'folder',
        'one_status_common_datasets_aoi.gdb': 'folder',
        'one_status_tabs_1_and_2_datasets.gdb': 'folder',
        'automated_status_sheet.xlsx': 'file',
        'one_status_common_datasets_aoi.xlsx': 'file',
        'one_status_tabs_1_and_2.xlsx': 'file'
    }
    
    missing_items = []

    # Check if output directory exists
    if not os.path.exists(output_directory):
        return (False, ['output_directory'])

    # The new runner (auto_status) creates a Status_<name>_<run_id> subdirectory inside
    # the job's output_directory. Detect that subdirectory and validate inside it instead.
    check_directory = output_directory
    try:
        subdirs = [
            d for d in os.listdir(output_directory)
            if os.path.isdir(os.path.join(output_directory, d))
            and d.startswith('Status_')
        ]
        if subdirs:
            # Use the most recently modified Status_ subfolder
            subdirs.sort(
                key=lambda d: os.path.getmtime(os.path.join(output_directory, d)),
                reverse=True
            )
            check_directory = os.path.join(output_directory, subdirs[0])
            if logger:
                logger.info(f"Output Validator: Resolved run directory to {check_directory}")
    except Exception as e:
        if logger:
            logger.warning(f"Output Validator: Could not scan for Status_ subdir: {e}")

    # Check each required item inside the resolved directory
    for item_name, item_type in required_items.items():
        item_path = os.path.join(check_directory, item_name)  # Validate inside resolved run dir

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
    
    return (len(missing_items) == 0, missing_items)

###############################################################################################################################################################################
# Set up logging

import os
import datetime
import logging
import arcpy


def setup_logging(custom_log_path=None):
    ''' 
    Set up logging for the script 
    
    Parameters:
        custom_log_path (str, optional): Custom directory path where log files should be written.
                                         If None or empty, uses the default location (script directory).
    '''
    arcpy.AddMessage("Setting up Logging")
    
    # Determine the base path for log folder
    if custom_log_path and custom_log_path.strip():
        # Use the custom path provided by the user
        base_path = custom_log_path.strip()
        arcpy.AddMessage(f"Using custom log path: {base_path}")
        print(f"Using custom log path: {base_path}")
    else:
        # Use the default path (script directory)
        base_path = os.path.dirname(os.path.realpath(__file__))
        arcpy.AddMessage("Using default log path (script directory)")
        print("Using default log path (script directory)")
    
    # Verify the base path exists and is writable
    if not os.path.exists(base_path):
        error_msg = f"Error: Log path does not exist: {base_path}"
        arcpy.AddError(error_msg)
        raise ValueError(error_msg)
    
    # Create the log folder filename
    log_folder = os.path.join(base_path, f'autoast_batch_logs_folder_{datetime.datetime.now().strftime("%Y%m%d")}')

    # Create the log folder if it doesn't exist
    try:
        if not os.path.exists(log_folder):
            os.mkdir(log_folder)
            arcpy.AddMessage(f"Created log folder: {log_folder}")
    except Exception as e:
        error_msg = f"Error creating log folder at {log_folder}: {e}"
        arcpy.AddError(error_msg)
        raise
    
    # Check if the log folder was created successfully
    assert os.path.exists(log_folder), "Error creating log folder, check permissions and path"

    # Create the log file path with the date and time appended
    log_file = os.path.join(log_folder, f'ast_batch_log_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.log')

    # Set up logging config to DEBUG level
    logging.basicConfig(filename=log_file, 
                        level=logging.DEBUG, 
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    # Add a StreamHandler so all logger calls also print to the console,
    # removing the need for separate print() calls in the main scripts.
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
    logging.getLogger().addHandler(console_handler)

    # Create the logger object and set to the current file name
    logger = logging.getLogger(__name__)

    print(f"Logging set up - Log file: {log_file}")
    logger.info(f"Logging set up - Log file: {log_file}")
    
    return logger
###############################################################################################################################################################################
###############################################################################################################################################################################
# Set up logging

import os
import datetime
import logging


def setup_logging():
    ''' Set up logging for the script '''
    # Create the log folder filename
    log_folder = f'autoast_logs_{datetime.datetime.now().strftime("%Y%m%d")}'

    # Create the log folder in the current directory if it doesn't exits
    if not os.path.exists(log_folder):
        os.mkdir(log_folder)
    
    # Check if the log folder was created successfully
    assert os.path.exists(log_folder), "Error creating log folder, check permissions and path"

    # Create the log file path with the date and time appended
    log_file = os.path.join(log_folder, f'ast_log_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.log')



    # Set up logging config to DEBUG level
    logging.basicConfig(filename=log_file, 
                        level=logging.DEBUG, 
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    # Create the logger object and set to the current file name
    logger = logging.getLogger(__name__)

    print("Logging set up")
    logger.info("Logging set up")
    
    return logger
###############################################################################################################################################################################
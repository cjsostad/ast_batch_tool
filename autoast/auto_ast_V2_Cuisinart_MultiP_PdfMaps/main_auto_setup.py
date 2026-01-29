# autoast is a script for batch processing the automated status tool
# author: csostad and wburt
# copyright Governent of British Columbia
# Copyright 2019 Province of British Columbia

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at 

#    http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.


import os
from dotenv import load_dotenv
from logging_setup import setup_logging
from database_connection import setup_bcgw
from toolbox_import import import_ast
from ast_factory import AST_FACTORY
from multi_excel_setup import create_job_excel_files

# snippet to run multiple terminal windows & "P:\corp\python_ast\python.exe" "W:\srm\nel\Local\Geomatics\Workarea\csostad\GitHubAutoAST\gss_authorizations\autoast\auto_ast_v3_Breville_folium_maps\main.py"


# Automatically creates the excel files in groups of 8 with a common region......





# Mandatory function that feeds the list of excel files into the Toaster
def process_excel_file(excel_file, secrets, logger, current_path):
    '''
    This function takes a list of excel files and iterates over that list, applying the Batch AST Class (and hence the ast tool)
    too each row in each excel file. This is a workaround for multiprocessing issue with the BCGW sees too many db connections
    in batches of 8 
    
    '''
    try:
        print(f"Main: Creating queuefile path for {excel_file}")
        logger.info(f"Main: Creating queuefile path for {excel_file}")

        qf = os.path.join(current_path, excel_file)

        # Create an instance of the AST_FACTORY class
        ast = AST_FACTORY(qf, secrets[0], secrets[1], logger, current_path)

        if not os.path.exists(qf):
            print(f"Main: Queuefile for {excel_file} not found, creating new queuefile")
            logger.info(f"Main: Queuefile for {excel_file} not found, creating new queuefile")
            ast.create_new_queuefile()

        # Load jobs from the Excel file
        print(f"Main: Loading jobs from {excel_file}")
        logger.info(f"Main: Loading jobs from {excel_file}")
        jobs = ast.load_jobs()

        # Batch jobs
        print(f"Main: Batching jobs for {excel_file}")
        logger.info(f"Main: Batching jobs for {excel_file}")
        ast.batch_ast()

        # Reload failed jobs
        print(f"Main: Reloading failed jobs for {excel_file}")
        logger.info(f"Main: Reloading failed jobs for {excel_file}")
        ast.re_load_failed_jobs_V2()

        # Re-batch failed jobs
        print(f"Main: Re-batching failed jobs for {excel_file}")
        logger.info(f"Main: Re-batching failed jobs for {excel_file}")
        ast.batch_ast()

        print(f"Main: AST Factory for {excel_file} COMPLETE")
        logger.info(f"Main: AST Factory for {excel_file} COMPLETE")
    
    
    except Exception as e:
        print(f"Error processing {excel_file}: {e}")
        logger.error(f"Error processing {excel_file}: {e}")

#################################################################################################################################################################################
if __name__ == '__main__':
    
    # First, run the Tkinter dialog to get excel files and log path from user
    # This must happen BEFORE setting up logging so the user can choose where logs go
    print("Main: Running 'Create job excel files'")
    excel_files, custom_log_path = create_job_excel_files()
    print(f"List of excel file paths is: {excel_files}")
    print(f"Custom log path: {custom_log_path if custom_log_path else 'Default location'}")
    
    # Now set up logging with the custom path provided by the user
    logger = setup_logging(custom_log_path)
    logger.info(f"List of excel file paths is: {excel_files}")
    logger.info(f"Custom log path: {custom_log_path if custom_log_path else 'Default location'}")

    # Load the default environment
    load_dotenv()

    # Call the import_ast function to import the AST toolbox
    template = import_ast(logger)
    
    current_path = os.path.dirname(os.path.realpath(__file__))

    # Call the setup_bcgw function to set up the database connection
    # secrets = setup_bcgw(logger)
    secrets, sde_connection, sde_path = setup_bcgw(logger)
    # username, password = secrets[0], secrets[1]
    
    # Set the SDE path environment variable for easy access by workers
    os.environ["SDE_FILE_PATH"] = sde_path
    logger.info(f"SDE Connection established at: {sde_path}")          
    
    
    # Process each Excel file
    if not excel_files or len(excel_files) == 0:
        error_msg = "No Excel files were created. This usually means no shapefiles were found in subfolders of the selected directory."
        print(f"ERROR: {error_msg}")
        logger.error(error_msg)
        logger.error("Make sure your directory structure has shapefiles in subfolders, not directly in the main folder.")
        print("\nScript completed with no files to process.")
    else:
        print(f"\nProcessing {len(excel_files)} Excel file(s)...")
        logger.info(f"Processing {len(excel_files)} Excel file(s)...")
        
        for excel_file in excel_files:
            process_excel_file(excel_file, secrets, logger, current_path)
        
        print("\nAll Excel files processed successfully!")
        logger.info("All Excel files processed successfully!")

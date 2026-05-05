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
from failed_job_tracker import FailedJobTracker  # Added: creates a summary spreadsheet of all failed jobs across all workbooks

# This script is the entry point when the client has already filled out their own spreadsheet(s).
# For the auto-setup path (spreadsheet created from a folder of shapefiles), use main_auto_setup.py instead.
# Both scripts share identical orchestration logic - any changes to shared logic must be mirrored in both files.
# The only intentional differences are:
#   - excel_files and delete_transitory are configured here as module-level variables (no GUI)
#   - setup_logging() is called with no argument (no custom log path dialog)


# *** INPUT YOUR EXCEL FILE NAME(S) HERE ***
# List each pre-filled client spreadsheet by filename (must be in the same directory as this script).
excel_files = [
    "gr_2026_147.xlsx"
]

# Set to True to delete transitory GDB/log files immediately after each job is verified COMPLETE.
# Failed/FAILED_OUTPUTS jobs are never cleaned -- their GDBs are required for restart analysis.
delete_transitory = False


# Mandatory function that feeds the list of excel files into the Toaster
def process_excel_file(excel_file, secrets, logger, current_path, failed_job_tracker, delete_transitory=False):
    '''
    This function takes a list of excel files and iterates over that list, applying the Batch AST Class (and hence the ast tool)
    to each row in each excel file. This is a workaround for multiprocessing issue with the BCGW sees too many db connections
    in batches of 8.
    failed_job_tracker: FailedJobTracker instance passed through to AST_FACTORY so all failures are recorded in the summary spreadsheet.
    delete_transitory: if True, transitory GDB/log files are deleted immediately after each job is verified COMPLETE (not in a
    bulk sweep at the end). Failed/FAILED_OUTPUTS jobs are never cleaned so their GDBs remain available for restart analysis.
    '''
    try:
        print(f"Main: Creating queuefile path for {excel_file}")
        logger.info(f"Main: Creating queuefile path for {excel_file}")

        qf = os.path.join(current_path, excel_file)

        # Create an instance of the AST_FACTORY class
        # Pass failed_job_tracker so AST_FACTORY logs failures to the summary spreadsheet
        # Pass delete_transitory so per-job cleanup is triggered only after verified COMPLETE
        ast = AST_FACTORY(qf, secrets[0], secrets[1], logger, current_path, failed_job_tracker, delete_transitory)

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

        print(f"\n{'='*100}")
        print(f"  >>> COMPLETED SPREADSHEET: {excel_file} <<<")
        print(f"{'='*100}\n")
        logger.info(f"\n{'='*100}")
        logger.info(f"  >>> COMPLETED SPREADSHEET: {excel_file} <<<")
        logger.info(f"{'='*100}\n")


    except Exception as e:
        print(f"Error processing {excel_file}: {e}")
        logger.error(f"Error processing {excel_file}: {e}")

#################################################################################################################################################################################
if __name__ == '__main__':

    print("Main: Starting AutoAST Batch Processor")

    # Set up logging -- no custom path dialog here; uses the default log location
    logger = setup_logging()
    logger.info(f"Excel files to process: {excel_files}")
    # Log the cleanup preference so it appears in the batch log for auditability
    logger.info(f"Delete Transitory Data on Completion: {delete_transitory}")

    # Load the default environment
    load_dotenv()

    # Call the import_ast function to import the AST toolbox
    template = import_ast(logger)

    current_path = os.path.dirname(os.path.realpath(__file__))
    # Initialize the failed job tracker; writes a timestamped summary spreadsheet to the script's directory
    failed_job_tracker = FailedJobTracker(current_path, logger)

    # Call the setup_bcgw function to set up the database connection
    secrets, sde_connection, sde_path = setup_bcgw(logger)

    # Set the SDE path environment variable for easy access by workers
    os.environ["SDE_FILE_PATH"] = sde_path
    logger.info(f"SDE Connection established at: {sde_path}")

    # Process each Excel file listed at the top of this script
    if not excel_files or len(excel_files) == 0:
        error_msg = "No Excel files are listed. Add at least one filename to the excel_files list at the top of this script."
        print(f"ERROR: {error_msg}")
        logger.error(error_msg)
        print("\nScript completed with no files to process.")
    else:
        print(f"\nProcessing {len(excel_files)} Excel file(s)...")
        logger.info(f"Processing {len(excel_files)} Excel file(s)...")

        for excel_file in excel_files:
            print(f"\n{'='*100}")
            print(f"{'='*100}")
            print(f"  >>> PROCESSING SPREADSHEET: {excel_file} <<<")
            print(f"{'='*100}")
            print(f"{'='*100}\n")
            logger.info(f"\n{'='*100}")
            logger.info(f"{'='*100}")
            logger.info(f"  >>> PROCESSING SPREADSHEET: {excel_file} <<<")
            logger.info(f"{'='*100}")
            logger.info(f"{'='*100}\n")
            # Pass failed_job_tracker so failures in each workbook are captured in the summary spreadsheet
            # Pass delete_transitory so AST_FACTORY can clean up each job immediately after COMPLETE verification.
            # Failed/FAILED_OUTPUTS jobs are never cleaned -- their GDBs are required for restart analysis.
            process_excel_file(excel_file, secrets, logger, current_path, failed_job_tracker, delete_transitory)

        print("\nAll Excel files processed successfully!")
        logger.info("All Excel files processed successfully!")
        # NOTE: Transitory data deletion is handled per-job inside AST_FACTORY.batch_ast(),
        # immediately after each job is verified COMPLETE. Failed/FAILED_OUTPUTS jobs are
        # never cleaned -- their GDBs are required for restart analysis on retries.
    
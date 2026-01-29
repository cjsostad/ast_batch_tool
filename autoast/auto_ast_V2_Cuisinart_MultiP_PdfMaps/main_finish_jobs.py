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
from openpyxl import load_workbook


# Jan 27, 2026: Helper function to check if workbook has any incomplete jobs
def has_incomplete_jobs(excel_file_path, logger):
    '''
    Checks if the Excel workbook has any jobs that are not marked as COMPLETE.
    Returns (has_incomplete, total_jobs, incomplete_count)
    '''
    try:
        if not os.path.exists(excel_file_path):
            return (True, 0, 0)  # File doesn't exist, treat as incomplete
        
        wb = load_workbook(filename=excel_file_path, read_only=True)
        ws = wb['Sheet1']  # Assuming standard sheet name
        
        # Get headers from first row
        headers = [cell.value for cell in ws[1]]
        
        # Find the ast_condition column index
        if 'ast_condition' not in headers:
            wb.close()
            return (True, 0, 0)  # No condition column, needs processing
        
        condition_col_idx = headers.index('ast_condition') + 1  # 1-indexed
        
        total_jobs = 0
        incomplete_jobs = 0
        
        # Check each row (skip header)
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if row_idx > ws.max_row:
                break
            
            # Skip empty rows
            if not any(row):
                continue
                
            total_jobs += 1
            condition = row[condition_col_idx - 1]  # 0-indexed for tuple
            
            if condition != 'COMPLETE':
                incomplete_jobs += 1
        
        wb.close()
        
        excel_file_name = os.path.basename(excel_file_path)
        if incomplete_jobs > 0:
            print(f"Main: [{excel_file_name}] Found {incomplete_jobs} incomplete jobs out of {total_jobs} total")
            logger.info(f"Main: [{excel_file_name}] Found {incomplete_jobs} incomplete jobs out of {total_jobs} total")
        else:
            print(f"Main: [{excel_file_name}] All {total_jobs} jobs already COMPLETE - skipping")
            logger.info(f"Main: [{excel_file_name}] All {total_jobs} jobs already COMPLETE - skipping")
        
        return (incomplete_jobs > 0, total_jobs, incomplete_jobs)
        
    except Exception as e:
        excel_file_name = os.path.basename(excel_file_path)
        print(f"Main: [{excel_file_name}] Error checking job status: {e} - will process anyway")
        logger.warning(f"Main: [{excel_file_name}] Error checking job status: {e} - will process anyway")
        return (True, 0, 0)  # Error checking, process to be safe


###################################################################################
#
# Jan 27, 2026 Update: Enhanced to intelligently reprocess failed/incomplete jobs
# - Timeout reduced to 3 hours per job (was 24 hours)
# - Checks for existing job status before processing
# - Skips fully completed workbooks
# - Reprocesses failed jobs up to 2 additional times
#
# This script is designed to finish processing jobs_9.xlsx through jobs_25.xlsx
# The spreadsheets and output directories have already been created
#
###################################################################################


# Mandatory function that feeds the list of excel files into the Toaster
def process_excel_file(excel_file_path, secrets, logger, current_path, max_retry_attempts=2):
    '''
    This function takes a list of excel files and iterates over that list, applying the Batch AST Class (and hence the ast tool)
    to each row in each excel file. This is a workaround for multiprocessing issue with the BCGW sees too many db connections
    in batches of 8
    
    Jan 27, 2026: Enhanced with retry logic for failed jobs (up to 2 additional attempts)
    '''
    try:
        excel_file_name = os.path.basename(excel_file_path)
        
        # Jan 27, 2026: Check if this workbook needs processing
        has_incomplete, total_jobs, incomplete_count = has_incomplete_jobs(excel_file_path, logger)
        
        if not has_incomplete:
            print(f"Main: [{excel_file_name}] Skipping - all jobs complete")
            logger.info(f"Main: [{excel_file_name}] Skipping - all jobs complete")
            return
        
        print(f"Main: [{excel_file_name}] Creating queuefile path for {excel_file_path}")
        logger.info(f"Main: [{excel_file_name}] Creating queuefile path for {excel_file_path}")

        # Create an instance of the AST_FACTORY class
        ast = AST_FACTORY(excel_file_path, secrets[0], secrets[1], logger, current_path)

        if not os.path.exists(excel_file_path):
            print(f"Main: [{excel_file_name}] ERROR - Queuefile {excel_file_path} not found!")
            logger.error(f"Main: [{excel_file_name}] ERROR - Queuefile {excel_file_path} not found!")
            return

        # Load jobs from the Excel file
        print(f"Main: [{excel_file_name}] Loading jobs from {excel_file_path}")
        logger.info(f"Main: [{excel_file_name}] Loading jobs from {excel_file_path}")
        jobs = ast.load_jobs()

        # Batch jobs (initial attempt)
        print(f"Main: [{excel_file_name}] Batching jobs for {excel_file_path} (initial attempt)")
        logger.info(f"Main: [{excel_file_name}] Batching jobs for {excel_file_path} (initial attempt)")
        ast.batch_ast()
        print(f"Main: [{excel_file_name}] Initial batch processing complete")
        logger.info(f"Main: [{excel_file_name}] Initial batch processing complete")

        # Jan 27, 2026: Retry failed jobs up to max_retry_attempts times
        for retry_num in range(1, max_retry_attempts + 1):
            print(f"Main: [{excel_file_name}] Checking for failed jobs (retry attempt {retry_num}/{max_retry_attempts})")
            logger.info(f"Main: [{excel_file_name}] Checking for failed jobs (retry attempt {retry_num}/{max_retry_attempts})")
            
            # Reload failed jobs
            print(f"Main: [{excel_file_name}] Reloading failed jobs for {excel_file_path}")
            logger.info(f"Main: [{excel_file_name}] Reloading failed jobs for {excel_file_path}")
            ast.re_load_failed_jobs_V2()
            print(f"Main: [{excel_file_name}] Failed jobs reload complete")
            logger.info(f"Main: [{excel_file_name}] Failed jobs reload complete")

            # Check if there are any failed jobs to retry
            if not ast.jobs or len(ast.jobs) == 0:
                print(f"Main: [{excel_file_name}] No failed jobs found - all complete!")
                logger.info(f"Main: [{excel_file_name}] No failed jobs found - all complete!")
                break
            
            # Re-batch failed jobs
            print(f"Main: [{excel_file_name}] Re-batching {len(ast.jobs)} failed jobs (attempt {retry_num})")
            logger.info(f"Main: [{excel_file_name}] Re-batching {len(ast.jobs)} failed jobs (attempt {retry_num})")
            ast.batch_ast()
            print(f"Main: [{excel_file_name}] Retry attempt {retry_num} complete")
            logger.info(f"Main: [{excel_file_name}] Retry attempt {retry_num} complete")

        print(f"Main: [{excel_file_name}] AST Factory for {excel_file_path} COMPLETE")
        logger.info(f"Main: [{excel_file_name}] AST Factory for {excel_file_path} COMPLETE")
    
    
    except Exception as e:
        excel_file_name = os.path.basename(excel_file_path)
        print(f"Main: [{excel_file_name}] ERROR processing {excel_file_path}: {e}")
        logger.error(f"Main: [{excel_file_name}] ERROR processing {excel_file_path}: {e}")

#################################################################################################################################################################################
if __name__ == '__main__':
    
    print("="*80)
    print("Main: Starting AutoAST V2 - Finish Remaining Jobs")
    print("Jan 27, 2026: Updated with 3-hour timeout and smart reprocessing")
    print("="*80)
    
    # Call the setup_logging function to log the messages
    logger = setup_logging()

    # Load the default environment
    load_dotenv()
    print("Main: Environment loaded")
    
    # Call the import_ast function to import the AST toolbox
    template = import_ast(logger)
    
    current_path = os.path.dirname(os.path.realpath(__file__))
    print(f"Main: Current path is {current_path}")
    
    # Call the setup_bcgw function to set up the database connection
    secrets, sde_connection, sde_path = setup_bcgw(logger)
    print("Main: BCGW Connection established")
    
    # Set the SDE path environment variable for easy access by workers
    os.environ["SDE_FILE_PATH"] = sde_path
    logger.info(f"SDE Connection established at: {sde_path}")
    
    
    # Define the directory where the excel files are located
    excel_directory = r"\\spatialfiles.bcgov\srm\gss\sandbox\csostad\Skeena 2026-2028 Shapefiles\shapefile\outputs"
    
    # List of Excel files to process (jobs_9.xlsx through jobs_25.xlsx)
    # Jan 27, 2026: Script will check each file and skip those already complete
    excel_files = [
        'jobs_9.xlsx',
        'jobs_10.xlsx',
        'jobs_11.xlsx',
        'jobs_12.xlsx',
        'jobs_13.xlsx',
        'jobs_14.xlsx',
        'jobs_15.xlsx',
        'jobs_16.xlsx',
        'jobs_17.xlsx',
        'jobs_18.xlsx',
        'jobs_19.xlsx',
        'jobs_20.xlsx',
        'jobs_21.xlsx',
        'jobs_22.xlsx',
        'jobs_23.xlsx',
        'jobs_24.xlsx',
        'jobs_25.xlsx',
    ]
    
    print(f"Main: Excel files directory: {excel_directory}")
    logger.info(f"Main: Excel files directory: {excel_directory}")
    
    # Jan 27, 2026: Track processing statistics
    processed_count = 0
    skipped_count = 0
    
    # Process each Excel file
    for excel_file in excel_files:
        excel_file_path = os.path.join(excel_directory, excel_file)
        print(f"\n{'='*80}")
        print(f"Main: Processing {excel_file}")
        print(f"{'='*80}\n")
        logger.info(f"\n{'='*80}")
        logger.info(f"Main: Processing {excel_file}")
        logger.info(f"{'='*80}\n")
        
        # Check if file needs processing before creating AST instance
        has_incomplete, total, incomplete = has_incomplete_jobs(excel_file_path, logger)
        
        if has_incomplete:
            process_excel_file(excel_file_path, secrets, logger, current_path)
            processed_count += 1
        else:
            skipped_count += 1
    
    print("\n" + "="*80)
    print(f"Main: All remaining jobs processing COMPLETE")
    print(f"Main: Processed {processed_count} workbooks, Skipped {skipped_count} complete workbooks")
    print("="*80)
    logger.info(f"Main: All remaining jobs processing COMPLETE - Processed {processed_count}, Skipped {skipped_count}")

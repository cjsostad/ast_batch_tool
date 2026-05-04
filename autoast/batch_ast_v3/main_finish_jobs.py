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
from failed_job_tracker import FailedJobTracker
from output_validator import verify_job_outputs


# Jan 27, 2026: Helper function to check if workbook has any incomplete jobs
# Feb 4, 2026: Enhanced to also verify output files exist and are valid
def has_incomplete_jobs(excel_file_path, logger, failed_job_tracker=None):
    '''
    Checks if the Excel workbook has any jobs that are not marked as COMPLETE
    AND verifies that all output files/folders exist and are not empty.
    Returns (has_incomplete, total_jobs, incomplete_count)
    '''
    try:
        if not os.path.exists(excel_file_path):
            return (True, 0, 0)  # File doesn't exist, treat as incomplete
        
        wb = load_workbook(filename=excel_file_path, read_only=True)
        ws = wb['ast_config']  # Standard sheet name for AST batch files
        
        # Get headers from first row
        headers = [cell.value for cell in ws[1]]
        
        # Find required column indices
        if 'ast_condition' not in headers:
            wb.close()
            return (True, 0, 0)  # No condition column, needs processing
        
        if 'output_directory' not in headers:
            wb.close()
            return (True, 0, 0)  # No output_directory column, needs processing
        
        condition_col_idx = headers.index('ast_condition') + 1  # 1-indexed
        output_dir_col_idx = headers.index('output_directory') + 1  # 1-indexed
        
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
            output_directory = row[output_dir_col_idx - 1]  # 0-indexed for tuple
            
            job_incomplete = False
            
            # Check spreadsheet status
            if condition != 'COMPLETE':
                job_incomplete = True
            else:
                # If marked COMPLETE, verify outputs actually exist
                is_valid, missing_items = verify_job_outputs(output_directory, logger)
                if not is_valid:
                    job_incomplete = True
                    excel_file_name = os.path.basename(excel_file_path)
                    missing_str = ', '.join(missing_items)
                    logger.warning(f"Main: [{excel_file_name}] Job {row_idx - 1} marked COMPLETE but missing outputs: {missing_str}")
                    
                    # Log to failed_job_tracker if provided
                    if failed_job_tracker:
                        # Build job_data dictionary from row
                        job_data = {header: value for header, value in zip(headers, row)}
                        failed_job_tracker.log_failed_job(
                            spreadsheet_name=excel_file_name,
                            job_index=row_idx - 1,  # 0-indexed job number
                            job_data=job_data,
                            failure_type='missing_outputs',
                            notes=f"Missing/invalid outputs: {missing_str}"
                        )
            
            if job_incomplete:
                incomplete_jobs += 1
        
        wb.close()
        
        excel_file_name = os.path.basename(excel_file_path)
        if incomplete_jobs > 0:
            logger.info(f"Main: [{excel_file_name}] Found {incomplete_jobs} incomplete jobs out of {total_jobs} total")
        else:
            logger.info(f"Main: [{excel_file_name}] All {total_jobs} jobs COMPLETE with verified outputs - skipping")
        
        return (incomplete_jobs > 0, total_jobs, incomplete_jobs)
        
    except Exception as e:
        excel_file_name = os.path.basename(excel_file_path)
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
def process_excel_file(excel_file_path, secrets, logger, current_path, failed_job_tracker, max_retry_attempts=0):
    '''
    This function takes a list of excel files and iterates over that list, applying the Batch AST Class (and hence the ast tool)
    to each row in each excel file. This is a workaround for multiprocessing issue with the BCGW sees too many db connections
    in batches of 8
    
    Jan 27, 2026: Enhanced with retry logic for failed jobs (up to 2 additional attempts)
    Jan 29, 2026: Changed max_retry_attempts default to 0 (no retries) - jobs run once only
    Feb 2, 2026: Added failed_job_tracker parameter to track failures
    '''
    try:
        excel_file_name = os.path.basename(excel_file_path)
        
        # Jan 27, 2026: Check if this workbook needs processing
        has_incomplete, total_jobs, incomplete_count = has_incomplete_jobs(excel_file_path, logger)
        
        if not has_incomplete:
            logger.info(f"Main: [{excel_file_name}] Skipping - all jobs complete")
            return
        
        logger.info(f"Main: [{excel_file_name}] Creating queuefile path for {excel_file_path}")

        # Create an instance of the AST_FACTORY class with failed job tracker
        ast = AST_FACTORY(excel_file_path, secrets[0], secrets[1], logger, current_path, failed_job_tracker)

        if not os.path.exists(excel_file_path):
            logger.error(f"Main: [{excel_file_name}] ERROR - Queuefile {excel_file_path} not found!")
            return

        # Load jobs from the Excel file
        logger.info(f"Main: [{excel_file_name}] Loading jobs from {excel_file_path}")
        jobs = ast.load_jobs()

        # Batch jobs (initial attempt)
        logger.info(f"Main: [{excel_file_name}] Batching jobs for {excel_file_path} (initial attempt)")
        ast.batch_ast()
        logger.info(f"Main: [{excel_file_name}] Initial batch processing complete")

        # Jan 27, 2026: Retry failed jobs up to max_retry_attempts times
        for retry_num in range(1, max_retry_attempts + 1):
            logger.info(f"Main: [{excel_file_name}] Checking for failed jobs (retry attempt {retry_num}/{max_retry_attempts})")
            
            # Reload failed jobs
            logger.info(f"Main: [{excel_file_name}] Reloading failed jobs for {excel_file_path}")
            ast.re_load_failed_jobs_V2()
            logger.info(f"Main: [{excel_file_name}] Failed jobs reload complete")

            # Check if there are any failed jobs to retry
            if not ast.jobs or len(ast.jobs) == 0:
                logger.info(f"Main: [{excel_file_name}] No failed jobs found - all complete!")
                break
            
            # Re-batch failed jobs
            logger.info(f"Main: [{excel_file_name}] Re-batching {len(ast.jobs)} failed jobs (attempt {retry_num})")
            ast.batch_ast()
            logger.info(f"Main: [{excel_file_name}] Retry attempt {retry_num} complete")

        logger.info(f"\n{'='*100}")
        logger.info(f"  >>> COMPLETED SPREADSHEET: {excel_file_name} <<<")
        logger.info(f"{'='*100}\n")
    
    
    except Exception as e:
        excel_file_name = os.path.basename(excel_file_path)
        logger.error(f"Main: [{excel_file_name}] ERROR processing {excel_file_path}: {e}")

#################################################################################################################################################################################
if __name__ == '__main__':
    
    print("="*80)
    print("Main: Starting AutoAST V2 - Finish Remaining Jobs")
    print("Jan 27, 2026: Updated with 3-hour timeout and smart reprocessing")
    print("Feb 2, 2026: Updated with 6-hour timeout, process health checks, and failed job tracking")
    print("="*80)
    
    # Call the setup_logging function to log the messages
    logger = setup_logging()

    # Load the default environment
    load_dotenv()
    logger.info("Main: Environment loaded")
    
    # Call the import_ast function to import the AST toolbox
    template = import_ast(logger)
    
    current_path = os.path.dirname(os.path.realpath(__file__))
    logger.info(f"Main: Current path is {current_path}")
    
    # Call the setup_bcgw function to set up the database connection
    secrets, sde_connection, sde_path = setup_bcgw(logger)
    
    # Set the SDE path environment variable for easy access by workers
    os.environ["SDE_FILE_PATH"] = sde_path
    logger.info(f"SDE Connection established at: {sde_path}")
    
    
    # Define the directory where the excel files are located
    excel_directory = r"\\spatialfiles.bcgov\srm\gss\sandbox\csostad\Skeena 2026-2028 Shapefiles\shapefile\outputs"
    
    # Feb 2, 2026: Initialize Failed Job Tracker
    failed_job_tracker = FailedJobTracker(excel_directory, logger)
    logger.info(f"Main: Failed Job Tracker initialized")
    
    # Feb 2, 2026: Dynamically discover all jobs_*.xlsx files in the directory
    # This replaces the hardcoded list to be more flexible
    logger.info(f"Main: Excel files directory: {excel_directory}")
    
    logger.info(f"Main: Scanning for jobs_*.xlsx files...")
    
    all_files = os.listdir(excel_directory)
    excel_files = []
    
    for filename in all_files:
        if filename.startswith('jobs_') and filename.endswith('.xlsx'):
            excel_files.append(filename)
    
    # Sort the files numerically (jobs_9.xlsx, jobs_10.xlsx, etc.)
    excel_files.sort(key=lambda x: int(x.replace('jobs_', '').replace('.xlsx', '')))
    
    logger.info(f"Main: Found {len(excel_files)} Excel files: {', '.join(excel_files)}")
    
    # Jan 27, 2026: Track processing statistics
    processed_count = 0
    skipped_count = 0
    
    # Process each Excel file
    for excel_file in excel_files:
        excel_file_path = os.path.join(excel_directory, excel_file)
        logger.info(f"\n{'='*100}")
        logger.info(f"{'='*100}")
        logger.info(f"  >>> PROCESSING SPREADSHEET: {excel_file} <<<")
        logger.info(f"{'='*100}")
        logger.info(f"{'='*100}\n")
        
        # Check if file needs processing before creating AST instance
        has_incomplete, total, incomplete = has_incomplete_jobs(excel_file_path, logger, failed_job_tracker)
        
        if has_incomplete:
            process_excel_file(excel_file_path, secrets, logger, current_path, failed_job_tracker)
            processed_count += 1
        else:
            skipped_count += 1
    
    # Feb 2, 2026: Print failed job tracker summary
    print("\n" + "="*80)
    failed_job_tracker.print_summary()
    print("="*80)
    
    logger.info(f"Main: All remaining jobs processing COMPLETE - Processed {processed_count}, Skipped {skipped_count}")

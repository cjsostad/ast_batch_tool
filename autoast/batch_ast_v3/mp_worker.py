import sys







def process_job_mp(ast_instance, job, job_index, current_path, sde_path, return_dict):
    import os
    import arcpy
    import datetime
    import logging
    import multiprocessing as mp
    import traceback
    from pathlib import Path  # Added: used to pass sde_path as a Path object to the runner

    logger = logging.getLogger(f"Process Job Mp: worker_{job_index}")

    logger.info("##########################################################################################################################")
    logger.info("#")
    logger.info("Running Multiprocessing Worker Function.....")
    logger.info("#")
    logger.info("##########################################################################################################################")

    print(f"Process Job Mp: Processing job {job_index}: {job}")
    arcpy.AddMessage("Inside Process Job MP")
    # version2 - old: injected shared SDE into arcpy workspace so workers could reach BCGW
    # The new fcbc_auto_status_tool manages its own per-job SDE connection internally
    # arcpy.env.workspace = sde_path
    
    # Set up logging folder in the worker process
    logger.info(f"Process Job Mp: Worker process {mp.current_process().pid} started for job {job_index}")
    log_folder = os.path.join(current_path, f'autoast_worker_logs_{datetime.datetime.now().strftime("%Y%m%d")}')
    if not os.path.exists(log_folder):
        os.mkdir(log_folder)
        logger.info(f"Process Job Mp: Created log folder {log_folder}")

    # Generate a unique log file name per process
    log_file = os.path.join(
        log_folder,
        f'ast_worker_log_{datetime.datetime.now().strftime("%Y_%m_%d_%H%M%S")}_{mp.current_process().pid}_job_{job_index}.log'
    )
    logger.info(f"Process Job Mp: Log file for worker process is: {log_file}")
    
    # Set up logging config in the worker process
    logging.basicConfig(
        filename=log_file,
        level=logging.DEBUG,  # Set level to DEBUG to capture all messages
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    try:
        # version2 - ImportToolbox + positional params list (replaced by runner approach below)
        # ast_toolbox = os.getenv('TOOLBOX')
        # ast_toolbox_alias = os.getenv('TOOLBOXALIAS')
        # if ast_toolbox:
        #     arcpy.ImportToolbox(ast_toolbox, ast_toolbox_alias)
        # else:
        #     raise ImportError("Process Job Mp: AST Toolbox path not found.")
        # params = []
        # for param in ast_instance.AST_PARAMETERS.values():
        #     value = job.get(param)
        #     if isinstance(value, str) and value.lower() in ['true', 'false']:
        #         value = True if value.lower() == 'true' else False
        #     params.append(value)

        # Derive the auto_status package root from the TOOLBOX path in .env
        # TOOLBOX = ...\Tools\src\auto_status\tools\fcbc_auto_status_tool.pyt
        # parents[2] resolves to ...\Tools\src\ which contains the auto_status package
        ast_toolbox_path = os.getenv('TOOLBOX')
        if not ast_toolbox_path:
            raise ImportError("Process Job Mp: TOOLBOX path not set in environment variables.")
        auto_status_src = str(Path(ast_toolbox_path.strip()).parents[2])
        if auto_status_src not in sys.path:
            sys.path.insert(0, auto_status_src)
        logger.info(f"Process Job Mp: auto_status package root resolved to: {auto_status_src}")

        # Import runner directly so the pre-made SDE can be passed in (avoids per-worker keyring lookups)
        # auto_status is resolved at runtime via sys.path.insert above; type: ignore suppresses Pylance static analysis warning
        from auto_status.analysis_tool import run as run_auto_status  # type: ignore

        # Build raw params dict matching AnalysisToolParams.from_mapping() keys
        raw = {
            'region':                                   job.get('region'),
            'feature_layer':                            job.get('feature_layer'),
            'crown_file_number':                        job.get('crown_file_number'),
            'disposition_number':                       job.get('disposition_number'),
            'parcel_number':                            job.get('parcel_number'),
            'output_directory':                         job.get('output_directory'),
            'output_dir_same_as_input':                 job.get('output_dir_same_as_input'),
            'dont_overwrite_outputs':                   job.get('dont_overwrite_outputs'),
            'dont_run_conflicts_and_constraints_tab3':  job.get('dont_run_conflicts_and_constraints_tab3'),
            'suppress_map_creation_tab3':               job.get('suppress_map_creation_tab3'),
            'open_output_directory':                    False,  # Batch workers cannot open Explorer windows
            'full_path_hyperlinks':                     job.get('full_path_hyperlinks'),
            'fcbc_spreadsheet_formatting':              True,   # AST always True
            'debug':                                    False,  # No shortened debug spreadsheets in batch
        }
        
        #NOTE: This is where the output directory is set
        # Get the output directory from the job
        output_directory = job.get('output_directory')

        # # If output_directory is not provided
        # if not output_directory:
        #     # Check if 'output directory is same as input directory' is set to True
        #     output_same_as_input = job.get('output_directory_is_same_as_input_directory')
        #     if output_same_as_input == True or str(output_same_as_input).lower() == 'true':
        #         # Use the input_directory as output_directory
        #         #NOTE This handling is already present in the AST Tool
        #         output_directory = job.get('input_directory')
        #         if not output_directory:
        #             raise ValueError(f"Process Job Mp: 'Input Directory' is required when 'Output Directory is same as Input Directory' is True for job {job_index}.")
        #         job['output_directory'] = output_directory
        #         logger.info(f"Process Job Mp: Output directory is same as input directory for job {job_index}. Using: {output_directory}")
        #     else:
        #         # If there was no output directory provided and 'output directory is same as input directory' is False
        #         # Set the default output directory to a default location (This can be changed later) This will prevent the job from failing due to a user error
                
        #         #DELETE This was put in for testing so that it's easy to delete all outputs from one place at once. 
        #         DEFAULT_DIR = os.getenv('DIR')
        #         output_directory = os.path.join("T:", f'job{job_index}')
        #         job['output_directory'] = output_directory
        #         logger.warning(f"Process Job Mp: Output directory not provided for job {job_index}. Using default path: {output_directory}")
        # else:
        #     # Output directory is provided
        #     job['output_directory'] = output_directory

        # Create the output directory if the user put in a path but failed to create the output directory in Windows explorer
        if output_directory and not os.path.exists(output_directory):
            try:
                os.makedirs(output_directory)
                print(f"Output directory '{output_directory}' created.")
                logger.warning(f"Process Job Mp: Output directory doesn't exist for job ({job_index}).")
                logger.warning(f"\n")
                logger.warning(f"'{output_directory}' created.")
            except OSError as e:
                raise RuntimeError(f"Failed to create the output directory '{output_directory}'. Check your permissions: {e}")


        # Ensure that region has been entered otherwise job will fail
        if not job.get('region'):
            raise ValueError("Process Job Mp: Region is required and was not provided. Job Failed")

        # version2 - logger.debug params list (replaced by raw dict logging below)
        # logger.debug(f"Process Job Mp: Job Parameters: {params}")
        logger.debug(f"Process Job Mp: Job raw params: {raw}")

        # version2 - arcpy GP tool calls (replaced by direct runner call below)
        # arcpy.alphaast.MakeAutomatedStatusSpreadsheet(*params)
        # arcpy.fcbc_auto_status_tool.AutomatedStatusTool(*params)

        # Call runner directly with the pre-made shared SDE — connection resolved once in main process
        # NOTE: Do not update Excel here - main process handles it after output validation
        logger.info(f"Process Job Mp: Calling auto_status runner with SDE: {sde_path}")
        run_auto_status(raw, sde=Path(sde_path))
        logger.info("Process Job Mp: auto_status runner completed successfully.")

        # version2 - arcpy GP message capture (not applicable when calling runner directly)
        # arcpy_messages = arcpy.GetMessages(0) ...
        
        # Indicate success
        return_dict[job_index] = 'Success'  

    except Exception as e:
        # Log the error first (before attempting Manager communication)
        logger.error(f"Process Job Mp: Job {job_index} failed with error: {e}")
        logger.debug(traceback.format_exc())
        exc_type, exc_value, exc_traceback = sys.exc_info()
        traceback_str = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        logger.error(f"Process Job Mp: Traceback:\n{traceback_str}")
        
        # Try to indicate failure via return_dict, but don't crash if Manager is dead
        try:
            return_dict[job_index] = 'Failed'
        except Exception as manager_error:
            logger.error(f"Process Job Mp: Failed to update return_dict (Manager connection may be dead): {manager_error}")
            logger.error(f"Process Job Mp: Original error that caused job failure was: {e}")

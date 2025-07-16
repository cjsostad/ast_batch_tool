import sys







def process_job_mp(batch_factory_instance, job, job_index, current_path, return_dict):
    import os
    import arcpy
    import datetime
    import logging
    import multiprocessing as mp
    import traceback
   

    logger = logging.getLogger(f"Process Job Mp: worker_{job_index}")

    logger.info("##########################################################################################################################")
    logger.info("#")
    logger.info("Running Multiprocessing Worker Function.....")
    logger.info("#")
    logger.info("##########################################################################################################################")

    print(f"Process Job Mp: Processing job {job_index}: {job}")

    # Set up logging folder in the worker process
    logger.info(f"Process Job Mp: Worker process {mp.current_process().pid} started for job {job_index}")
    log_folder = os.path.join(current_path, f'autoast_logs_{datetime.datetime.now().strftime("%Y%m%d")}')
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
    
    # Set up the parameters for the job from the list
    params = []
    for param_name in batch_factory_instance.parameter_names:
        value = job.get(param_name)
        
    logger.info(f"Parameters for job {job_index}: {params}")
    print(f"Parameters for job {job_index}: {params}")
    
    
    try:
        # Re-import the toolbox in each process
        any_toolbox = os.getenv('TOOLBOX')  # Get the toolbox path from environment variables
        
        if any_toolbox:
            arcpy.ImportToolbox(any_toolbox)
            print(f"Process Job Mp: {any_toolbox} imported successfully in worker.")
            logger.info(f"Process Job Mp: {any_toolbox} imported successfully in worker.")
        else:
            raise ImportError("Process Job Mp: Toolbox path not found. Ensure TOOLBOX path is set correctly in environment variables.")

        any_tool = os.getenv('TOOL') # Get the tool name from environment variables
        if not any_tool:
            raise ImportError("Tool name not found in .env")

        tool_func = getattr(arcpy, any_tool, None)
        if not tool_func:
            raise AttributeError(f"Tool '{any_tool}' not found in the toolbox.")
    
        
        # Prepare parameters
        params = []

        # Convert 'true'/'false' strings to booleans
        for param in batch_factory_instance.BATCH_PARAMETERS.values(): # use the batch_factory_instance that is passed into the function to access the ast factory parameters
            value = job.get(param)
            if isinstance(value, str) and value.lower() in ['true', 'false']:
                value = True if value.lower() == 'true' else False
            params.append(value)
        
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


    

        # Log the parameters being used
        logger.debug(f"Process Job Mp: Job Parameters: {params}")

        # Run the tool
        logger.info("Process Job Mp: Running your tool in Multiprocessing Batch Mode...Hold on!!!! ...")
        logger.info(f"Running tool {any_tool} with params={params}")
        arcpy.tool_func(*params)
        logger.info("Process Job Mp: Your Tool completed successfully.")
        batch_factory_instance.add_job_result(job_index, 'COMPLETE')

        # Capture and log arcpy messages
        logger.info("Process Job Mp: Capturing arcpy messages...")
        arcpy_messages = arcpy.GetMessages(0)
        arcpy_warnings = arcpy.GetMessages(1)
        arcpy_errors = arcpy.GetMessages(2)

        if arcpy_messages:
            logger.info(f'arcpy messages: {arcpy_messages}')
        if arcpy_warnings:
            logger.warning(f'arcpy warnings: {arcpy_warnings}')
        if arcpy_errors:
            logger.error(f'arcpy errors: {arcpy_errors}')
        
        # Indicate success
        return_dict[job_index] = 'Success'  

    except Exception as e:
        # Indicate failure
        return_dict[job_index] = 'Failed'
        logger.error(f"Process Job Mp: Job {job_index} failed with error: {e}")
        logger.debug(traceback.format_exc())
        exc_type, exc_value, exc_traceback = sys.exc_info()
        traceback_str = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        logger.error(f"Process Job Mp: Job {job_index} failed with error: {e}")
        logger.error(f"Process Job Mp: Traceback:\n{traceback_str}")

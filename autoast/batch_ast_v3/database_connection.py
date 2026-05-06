###############################################################################################################################################################################
#
# Set up the database connection
#
###############################################################################################################################################################################
import os
from dotenv import load_dotenv
import arcpy

def setup_bcgw(logger):
    # Get the secret file containing the database credentials
    SECRET_FILE = os.getenv('SECRET_FILE')

    # If secret file found, load the secret file and display a print message, if not found display an error message
    if SECRET_FILE:
        load_dotenv(SECRET_FILE)
        print(f"Secret file {SECRET_FILE} found")
        logger.info(f"Secret file {SECRET_FILE} found")
    else:
        print("Secret file not found")
        logger.error("Secret file not found")

    # Assign secret file data to variables    
    DB_USER = os.getenv('BCGW_USER')
    DB_PASS = os.getenv('BCGW_PASS')
    

    # If DB_USER and DB_PASS found display a print message, if not found display an error message
    if DB_USER and DB_PASS:
        print(f"Database user {DB_USER} and password found")
        logger.info(f"Database user {DB_USER} and password found")
    else:
        print("Database user and password not found")
        logger.error("Database user and password not found")

    # Define current path of the executing script
    current_path = os.path.dirname(os.path.realpath(__file__))

    # SDE_CONNECTION_DIR can be injected by Jenkins as a build environment variable to
    # redirect bcgw.sde to a path writable by the service account (e.g., the objectstore).
    # When running from VSCode, SDE_CONNECTION_DIR is not set so os.getenv() returns None
    # and the 'or' falls back to the original connection/ folder next to this script —
    # no change in behaviour for interactive runs.
    connection_folder = os.getenv('SDE_CONNECTION_DIR') or os.path.join(current_path, 'connection')

    # Check for the existance of the connection folder and if it doesn't exist, print an error and create a new connection folder
    if not os.path.exists(connection_folder):
        print("Connection folder not found, creating new connection folder")
        logger.info("Connection folder not found, creating new connection folder")
        # makedirs handles nested paths (e.g., a new subdirectory on the objectstore)
        os.makedirs(connection_folder, exist_ok=True)

    # Check for an existing bcgw connection, if there is one, remove it
    if os.path.exists(os.path.join(connection_folder, 'bcgw.sde')):
        os.remove(os.path.join(connection_folder, 'bcgw.sde'))

    # Create a bcgw connection
    bcgw_con = arcpy.management.CreateDatabaseConnection(connection_folder,
                                                        'bcgw.sde',
                                                        'ORACLE',
                                                        'bcgw.bcgov/idwprod1.bcgov',
                                                        'DATABASE_AUTH',
                                                        DB_USER,
                                                        DB_PASS,
                                                        'SAVE_USERNAME')
    
    #NOTE - Changed save username from do not save to save username

    print("new db connection created")
    logger.info("new db connection created")


    arcpy.env.workspace = bcgw_con.getOutput(0)
    
    sde_path = os.path.join(connection_folder, 'bcgw.sde')

    print("workspace set to bcgw connection")
    logger.info("workspace set to bcgw connection")
    
    secrets = [DB_USER, DB_PASS]
    sde_connection = arcpy.env.workspace
    
    return secrets, sde_connection, sde_path

###############################################################################################################################################################################

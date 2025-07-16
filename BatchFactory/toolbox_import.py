import os
import arcpy

#NOTE - Need to remove the template portion in the future

def import_any_toolbox(logger):
    print("Attempting to import the toolbox...")
    # Get the toolbox path from environment variables
    any_toolbox = os.getenv('TOOLBOX') # File path 

    if any_toolbox is None:
        print("Unable to find the toolbox. Check the path in .env file")
        logger.error("Unable to find the toolbox. Check the path in .env file")
        exit() 

    # Import the toolbox
    try:
        arcpy.ImportToolbox(any_toolbox)
        print(f"Toolbox imported successfully.")
        logger.info(f"Toolbox imported successfully.")
    except Exception as e:
        print(f"Error importing toolbox: {e}")
        logger.error(f"Error importing toolbox: {e}")
        exit()

    # Assign the shapefile template for FW Setup to a variable
    template = os.getenv('TEMPLATE') # File path in .env
    if template is None:
        print("Unable to find the template. Check the path in .env file")
        logger.error("Unable to find the template. Check the path in .env file")
        
    return template
import os
import arcpy

#NOTE - Need to remove the template portion in the future

def import_ast(logger):
    # Get the toolbox path from environment variables
    ast_toolbox = os.getenv('TOOLBOX') # File path 
    #NOTE transfer this to modular version
    ast_tool_alias = os.getenv("TOOLBOXALIAS") # Alias name

    if ast_toolbox is None:
        print("Unable to find the toolbox. Check the path in .env file")
        logger.error("Unable to find the toolbox. Check the path in .env file")
        exit() 

    # Import the toolbox
    try:
        # print(arcpy.ListTools("*"))
        arcpy.ImportToolbox(ast_toolbox, ast_tool_alias)
 
 
        logger.info(f"AST Toolbox imported successfully.")
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
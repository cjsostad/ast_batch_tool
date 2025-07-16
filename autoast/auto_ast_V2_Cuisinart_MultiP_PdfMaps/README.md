# autoast
This will be a tool to manage and automate status processes from a queue. 

# requirements
geopandas  
openpyxl  
arcpy  
automated status tool

You need a test excel spreadsheet. 
If file_number is filled out, the script will run the FW Setup tool on the feature layer. Enter the file number of the permit and it
will create shapefiles and .kml files in the appropriate directory the way the old FW Setup did. 

If file number is left blank, the script will pass the raw shapefile or .kml into the Ast Toolbox.
Be sure to update the output directory to the output where you want the results of the AST Toolbox to be placed. 








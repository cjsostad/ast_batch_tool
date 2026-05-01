
import arcpy
import os
import datetime
import geopandas
import shutil

# Assign the shapefile template for FW Setup to a variable
template = os.getenv('TEMPLATE') # File path in .env

# def classify_input_type(job, logger):
#     '''Classify the input type and process accordingly.'''

#     if job.get('feature_layer'):
#         print(f'Feature layer found: {job["feature_layer"]}')
#         logger.info(f'Classifying Input Type - Feature layer found: {job["feature_layer"]}')
#         feature_layer_path = job['feature_layer']
#         print(f"Processing feature layer: {feature_layer_path}")
#         logger.info(f"Classifying Input Type - Processing feature layer: {feature_layer_path}")

#         if feature_layer_path.lower().endswith('.kml'):
#             print('KML found, building AOI from KML')
#             logger.info('Classifying Input Type - KML found, building AOI from KML')
#             job['feature_layer'] = build_aoi_from_kml(job, feature_layer_path)

#         elif feature_layer_path.lower().endswith('.shp'):
#             if job.get('file_number'):
#                 print(f"File number found, running FW setup on shapefile: {feature_layer_path}")
#                 logger.info(f"Classifying Input Type - File number found, running FW setup on shapefile: {feature_layer_path}")
#                 new_feature_layer_path = build_aoi_from_shp(job, feature_layer_path)
#                 job['feature_layer'] = new_feature_layer_path
#             else:
#                 print('No FW File Number provided for the shapefile, using original shapefile path')
#                 logger.info('Classifying Input Type - No FW File Number provided, using original shapefile path')
#         else:
#             print(f"Unsupported feature layer format: {feature_layer_path}")
#             logger.warning(f"Classifying Input Type - Unsupported feature layer format: {feature_layer_path} - Marking job as Failed")
#             #add_job_result(job, 'Failed')
#     else:
#         print('No feature layer provided in job')
#         logger.warning('Classifying Input Type - No feature layer provided in job')

def build_aoi_from_kml(aoi, logger):
        "Write shp file for temporary use"

        # Ensure the KML file exists
        if not os.path.exists(aoi):
            raise FileNotFoundError(f"The KML file '{aoi}' does not exist.")

        print("Building AOI from KML")
        logger.info("Building AOI from KML")
        from fiona.drvsupport import supported_drivers
        supported_drivers['LIBKML'] = 'rw'
        tmp = os.getenv('TEMP')
        if not tmp:
            raise EnvironmentError("TEMP environment variable is not set.")
        bname = os.path.basename(aoi).split('.')[0]
        fc = bname.replace(' ', '_')
        out_name = os.path.join(tmp, bname + '.gdb')
        if os.path.exists(out_name):
            shutil.rmtree(out_name, ignore_errors=True)
        df = geopandas.read_file(aoi)
        df.to_file(out_name, layer=fc, driver='OpenFileGDB')

        #DELETE?
        print(f' kml ouput is {out_name} / {fc}')
        logger.info(f' kml ouput is {out_name} / {fc}')
        return out_name + '/' + fc



def build_aoi_from_shp(job, feature_layer_path, template, logger):
        """This is snippets of Mike Eastwoods FW Setup Script, if run FW Setup is set to true **Not sure if we need this
        as an option or just make it standard.
        This function will take the raw un-appended shapefile and run it through the FW Setup Script"""
        
        if template is None:
            print("Unable to find the template. Check the path in .env file")
            logger.error("Unable to find the template. Check the path in .env file")

        # Mike Eastwoods FW Setup Script
        print("Processing shapefile using FW Setup Script")
        logger.info("Processing shapefile using FW Setup Script")
        
        fsj_workspace = os.getenv('FSJ_WORKSPACE')
        arcpy.env.workspace = fsj_workspace
        arcpy.env.overwriteOutput = False

        # Check if there is a file path in Feature Layer
        if feature_layer_path:
            print(f"Processing feature layer: {feature_layer_path}")
            logger.info(f"Processing feature layer: {feature_layer_path}")

        # Check to see if a file number was entered in the excel sheet, if so, use it to name the output directory and authorize the build_aoi_from_shp function to run
        file_number = job.get('file_number')

        if not file_number:
            raise ValueError("Error: File Number is required if you are putting in a shapefile that has not been processed in the FW Setup Tool.")
        else:
            print(f"Running FW Setup on File Number: {file_number}")
            logger.info(f"Running FW Setup on File Number: {file_number}")

        # Convert file_number to string and make it uppercase
        file_number_str = str(file_number).upper()

        # Calculate date variables
        date = datetime.date.today()
        year = str(date.year)

        # Set variables
        base = arcpy.env.workspace
        baseYear = os.path.join(base, year)
        outName = file_number_str
        geometry = "POLYGON"
    
        m = "SAME_AS_TEMPLATE"
        z = "SAME_AS_TEMPLATE"
        spatialReference = arcpy.Describe(template).spatialReference

        # ===========================================================================
        # Create Folders
        # ===========================================================================

        print("Creating FW Setup folders . . .")
        logger.info("Creating FW Setup folders . . .")
        outName = file_number_str

        # Create path to folder location
        fileFolder = os.path.join(baseYear, outName)
        shapeFolder = fileFolder
        outPath = shapeFolder
        if os.path.exists(fileFolder):
            print(outName + " folder already exists.")
            logger.info(outName + " folder already exists.")
        else:
            os.mkdir(fileFolder)

        # ===========================================================================
        # Create Shapefile(s) and add them to the current map
        # ===========================================================================

        print("Creating Shapefiles using FW Setup . . .")
        logger.info("Creating Shapefiles using FW Setup . . .")
        if os.path.isfile(os.path.join(outPath, outName + ".shp")):
            print(os.path.join(outPath, outName + ".shp") + " already exists")
            logger.info(os.path.join(outPath, outName + ".shp") + " already exists")
            print("Exiting without creating files")
            logger.info("Exiting without creating files")
            return os.path.join(outPath, outName + ".shp")
        else:
            # Creating template shapefile
            create_shp = arcpy.management.CreateFeatureclass(outPath, outName, geometry, template, m, z, spatialReference)
            # Append the newly created shapefile with area of interest
            append_shp = arcpy.management.Append(feature_layer_path, create_shp, "NO_TEST")
            print("Append Successful")
            logger.info("Append Successful")
            # Making filename for kml
            create_kml = os.path.join(outPath, outName + ".kml")
            # Make layer for kml to be converted from 
            layer_shp = arcpy.management.MakeFeatureLayer(append_shp, outName)
            # Populate the shapefile                          
            arcpy.conversion.LayerToKML(layer_shp, create_kml)
            # Send message to user that kml has been created
            print("kml created: " + create_kml)
            logger.info("kml created: " + create_kml)

            print(f"FW Setup complete, returned shapefile is {os.path.join(outPath, outName + '.shp')}")
            logger.info(f"FW Setup complete, returned shapefile is {os.path.join(outPath, outName + '.shp')}")

            return os.path.join(outPath, outName + ".shp")
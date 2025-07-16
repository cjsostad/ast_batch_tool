# ArcGIS Pro 3 Toolbox Template

#===========================================================================
# Script name: Arcgis Pro 3 Toolbox Template
# Author: https://pro.arcgis.com/en/pro-app/latest/arcpy/geoprocessing_and_python/a-template-for-python-toolboxes.htm

# Created on: 10/21/2024

#New sample changes
# 
#

# 
#
# 
#============================================================================

import arcpy


class Toolbox:
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Descriptive Name of your Toolbox"
        self.alias = "toolbox"

        # List of tool classes associated with this toolbox
        self.tools = [NameOfYourTool, MySecondTool]  
        # Insert the name of each tool in your toolbox if you have more than one. 
        # i.e. self.tools = [FullSiteOverviewMaps, ExportSiteAndImageryLayout, Amendment]


class NameOfYourTool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Tool"
        self.description = ""

    def getParameterInfo(self):
        """This function assigns parameter information for tool""" 
        
        
        #This parameter is the file number of the application
        region = arcpy.Parameter(
            displayName = "Region",
            name="region",
            datatype="String",
            parameterType="Required",
            direction="Input")
        
        
        # This parameter is the proponent name for the cutting permit
        shape_file = arcpy.Parameter(
            displayName="Shapefile",
            name="shape_file",
            datatype="GPFeatureLayer",
            parameterType="Optional",
            direction="Input"
            )

        # This parameter is the cutting permit or TSL number
        crown_file_num = arcpy.Parameter(
            displayName="Crown File Number",
            name="crown_file_nuum",
            datatype="String",
            parameterType="Optional",
            direction="Input"
            ) 
        #variable name
        disp_num = arcpy.Parameter( #variable name
            displayName="Disposition Number",
            name="disp_num",
            datatype="String",
            parameterType="Optional",
            direction="Input"
            ) 


        parameters = [region, shape_file, crown_file_num, disp_num]  # Each parameter name needs to be in here, separated by a comma

        return parameters


    
    def isLicensed(self):
        """Set whether the tool is licensed to execute."""
        return True
    

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter. This method is called after internal validation."""
        return

    
    #NOTE This is where you cut a past your code
    def execute(self, parameters, messages):
        
        """The source code of the tool."""
        
        
        # Assign your parameters to variables
        region = parameters[0].valueAsText  # Take the first parameter **0 indexed!!** and assign in the the variable file_num
        shape_file = parameters[1].valueAsText
        crown_file_num = parameters[2].valueAsText
        disp_num = parameters[3].valueAsText
        # rp_ID = parameters[4].valueAsText
        # rp_amendment = parameters[5].valueAsText
        # rp_sections = parameters[6].valueAsText
        # sup_ID = parameters[7].valueAsText
        
        # Now write your script
        arcpy.AddMessage(f"Hello World, the file number is {crown_file_num}, the region is {region}, and the shapefile path  is {shape_file}")
        
        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return
    
#NOTE add another tool to the toolbox if you want
class MySecondTool(object):
    def __init__(self):
        
        """This tool will prep all required data for an individual crown tenure - to be used to add/subtract amendment - it will create the 
        amendment polygon, centroid, and splitline, and calculate geometries for centroid and splitline"""
        
        self.label = "Descriptive Name of your Second Tool"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """This function assigns parameter information for tool""" 
        
        
        #This parameter is the file number of the application
        amend_file_num = arcpy.Parameter(
            displayName = "Lands Amendment File Number",
            name="file_num",
            datatype="String",
            parameterType="Required",
            direction="Input")
        
        parameters = [amend_file_num] #Each parameter name needs to be in here, separated by a comma

        return parameters

    
    def execute(self,parameters,messages):
        
        # Bring in parameters to the function to be used as variables 
        amend_file_num = parameters[0].valueAsText
        
        arcpy.addMessage(f"Hello World, the file number is {amend_file_num}")
        
        ############################################################################################################################################################################
        #
        # Create the shapefile polygon layer to be used for the Amendment.
        #
        ############################################################################################################################################################################

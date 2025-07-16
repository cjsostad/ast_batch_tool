'''
Author: Mark McGirrX
Purpose:    This is the "CALL ROUTINE" for the universal overlap tool.  It takes all the input 
            arguments from the tool interface and builds a list out of them.  It then invokes the
            Universal_Overlap_Tool (revolt) and passes in that list. 

            The universal_overlap_tool (revolt) takes an input featureclass, from either a geodatabase or a shape file
            which can be a point line or polygon and overlays it with all the files specified 
            in an analysis_type spreadsheet.
            It generate an XLS file of all the overlaps and creates a small overlap map for each
            of the overlap types.
            
Date: Feb 2014

Arguments: argv[1] = featureclass to analyze
           argv[2] = create subreports on this field inthe above featureclass
           argv[3] = the type of analysis to run.  This is a pick list of xls spreadsheets stored on the P: drive
           argv[4] = an xls file the you can use for inputs if you don't want one of the ones from the pick list
           argv[5] = header information to print on the output xls file
           argv[6] = header information to print on the output xls file
           argv[7] = header information to print on the output xls file
           argv[8] = header information to print on the output xls file
           arcv[9] = boolean if you want the individual subreports to be on the same page, or different xls tabs
           arcv[10] = path to create the output xls in if not in same directory as the input featureclass
           arcv[11] = boolean if you want report fields to be split onto multile lines of the xls



Outputs:    The universal_overlap_tool creates an xls file of all the overlaps,
            listing details of each type of overlap, and producing
            maps to view each of the overlaps in.

Dependencies: MUST BE RUN IN ArcGIS PRO


History:
----------------------------------------------------------------------------------------------
Date: 
Author: 
Modification: 
'----------------------------------------------------------------------------------'

''' 


#===============================================================================
# Import system modules, corporate tools library, universal_overlap_tool.
#===============================================================================
#import sys, string, os, time,win32com.client,datetime,win32api,arcpy, csv
import sys, string, os, time, datetime, arcpy, csv, subprocess, keyring


#stop
message = "Now Running " + str(sys.argv[0])
#EDIT changed the path to the statusing tool to the new location

# import the statusing tool
# sys.path.append(r'\\GISWHSE.ENV.GOV.BC.CA\WHSE_NP\corp\script_whse\python\Utility_Misc\Ready\statusing_tools_arcpro\beta')
sys.path.append(r'\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\csostad\GitHubAutoAST\gss_authorizations\autoast\auto_ast_V2_Cuisinart_MultiP_PdfMaps')
import universal_overlap_tool_arcpro as revolt
# import create_bcgw_sde_connection as connect_bcgw
import config

print("Getting SDE File Path from os.getenv")

#EDIT - get the SDE file path from the environment variable
sde = os.getenv("SDE_FILE_PATH")

if not sde or not os.path.exists(sde):
    arcpy.AddError("SDE connection file not found or not accessible. Check environment variable 'SDE_FILE_PATH'.")
    sys.exit()
# Verify it works
print("Database User Found")




#------------------------------------------------------------------------------ 
arcpy.AddMessage("======================================================================")
arcpy.AddMessage("Checking for ArcGIS Pro Advanced license")

#Check to ensure Advanced licencing has been applied.
advStatus = ["Available", "AlreadyInitialized"]
if arcpy.CheckProduct("ArcInfo") not in advStatus:
    msg = 'ArcGIS Pro Advanced license not available. Set license to "Advanced" and try again'
    arcpy.AddError(msg)
    sys.exit()

arcpy.AddMessage("======================================================================")

#===============================================================================
# Set empty variables in case the passed in arguments are no longer needed.
# Not all of these are used in the PRO version, but some blanks will be
# passed into the universal overlap tool as place holders. 
#===============================================================================
''' This sets up some empty variables that the system will need
    if the reading_arguments section further down fails.
'''
input_feature_class = ""                            
sub_reports_on_this_field = ""                      
xls_file1_overlap_to_run_from_dropdown_list = ""    
xls_file1_overlap_to_run_from_user_specified = ""   
report_header_line1 = ""                            
report_header_line2 = ""                            
report_header_line3 = ""                            
report_header_line4 = ""                            
sub_reports_on_seperate_sheets = ""                 
directory_to_store_output = ""                      
summary_fields_on_seperate_lines = ""               
dont_overwrite_existing_data = ""               
xls_file2_overlap_to_add_for_autostatus_region = "" 
red_report_header_disclaimer = ""                   
suppress_map_creation = ""                                               
region = ""
crown_file_number = ""
disposition_number = ""
parcel_number = ""
run_as_fcbc = ""
aprx_path = ""
add_maps_to_current = ""



#===============================================================================
# Read arguments passed from tool
#===============================================================================
''' This reads the arguments passed in by the tool and sets variables based on their names. 
'''
try:
    arcpy.AddMessage("Assigning arguments to variables")
    if sys.argv[1]:
        input_feature_class = sys.argv[1]
        arcpy.AddMessage("input_feature_class " + input_feature_class)
    if sys.argv[2]:
        sub_reports_on_this_field = sys.argv[2]
        if sub_reports_on_this_field == '#':
            sub_reports_on_this_field = ''
        arcpy.AddMessage("sub_reports_on_this_field " + sub_reports_on_this_field)
    if sys.argv[3]:
        if sys.argv[3] == "":
            xls_file1_overlap_to_run_from_dropdown_list = ""
        elif sys.argv[3] == "#":
            xls_file1_overlap_to_run_from_dropdown_list = ""
        else:
            xls_file1_overlap_to_run_from_dropdown_list = os.path.join(r'\\Giswhse.env.gov.bc.ca\whse_np\corp\script_whse\python\Utility_Misc\Ready\statusing_tools_arcpro\statusing_input_spreadsheets',str(sys.argv[3]))
        arcpy.AddMessage("xls_file1_overlap_to_run_from_dropdown_list" + xls_file1_overlap_to_run_from_dropdown_list)
    if sys.argv[4]:
        xls_file1_overlap_to_run_from_user_specified = sys.argv[4]
        arcpy.AddMessage("xls_file1_overlap_to_run_from_user_specified " + xls_file1_overlap_to_run_from_user_specified)
    if sys.argv[5]:
        directory_to_store_output = sys.argv[5]
        arcpy.AddMessage("directory_to_store_output " + directory_to_store_output)
    if sys.argv[6]:
        output_dir_same_as_input = sys.argv[6]
        arcpy.AddMessage("output_dir_same_as_input " + output_dir_same_as_input)
    if sys.argv[7]:
        dont_overwrite_existing_data = sys.argv[7]
        arcpy.AddMessage("dont_overwrite_existing_data " + dont_overwrite_existing_data)
    if sys.argv[8]:
        suppress_map_creation = sys.argv[8]
        arcpy.AddMessage("suppress_map_creation " + suppress_map_creation)
    if sys.argv[9]:
        add_maps_to_current = sys.argv[9]
        arcpy.AddMessage("add_maps_to_current " + add_maps_to_current)
    if sys.argv[10]:
        red_report_header_disclaimer = sys.argv[10]
        if red_report_header_disclaimer == '#': 
            red_report_header_disclaimer = ''
        arcpy.AddMessage("red_report_header_disclaimer " + red_report_header_disclaimer)
    if sys.argv[11]:
        report_header_line1 = sys.argv[11]
        if report_header_line1 == '#':
            report_header_line1 = ''
        arcpy.AddMessage("report_header_line1 " + report_header_line1)
    if sys.argv[12]:
        report_header_line2 = sys.argv[12]
        if report_header_line2 == '#':
            report_header_line2 = ''
        arcpy.AddMessage("report_header_line2 " + report_header_line2)
    if sys.argv[13]:
        report_header_line3 = sys.argv[13]
        if report_header_line3 == '#':
            report_header_line3 = ''
        arcpy.AddMessage("report_header_line3 " + report_header_line3)
    if sys.argv[14]:
        report_header_line4 = sys.argv[14]
        if report_header_line4 == '#':
            report_header_line4 = ''
        arcpy.AddMessage("report_header_line4 " + report_header_line4)
except:
    pass

#update directory_to_store_output if not explicitly provided by user in tool
try:
    if directory_to_store_output != "#" and directory_to_store_output != "":
        if not os.path.exists(directory_to_store_output):
            try:
                os.makedirs(directory_to_store_output)
            except:
                arcpy.AddError("The Output Folder Directory Does Not Exist and Could Not Be Created")
                sys.exit()
    # Output directory set to where the input shapefile/feature class resides
    else:
        desc = arcpy.Describe(input_feature_class)
        gis_data_types = ["ShapeFile", "FeatureLayer", "FeatureClass"]
        if desc.dataType in gis_data_types:
            analyize_this_featureclass = desc.catalogPath
        else:
            arcpy.AddError(desc.dataType)
        directory_to_store_output = revolt.get_fc_directory_name(str(analyize_this_featureclass))
except Exception as e:
    arcpy.AddWarning(e)
    sys.exit()

# end of Read arguments passed from tool
#------------------------------------------------------------------------------

arcpy.AddMessage("======================================================================")
arcpy.AddMessage("Checking BCGW Credentials - may take a minute to process...")

#set the key name that will be used for storing credentials in keyring
key_name = config.CONNNAME
# try:
#     oracleCreds = connect_bcgw.ManageCredentials(key_name, directory_to_store_output)
#     #get sde path location
#     if not oracleCreds.check_credentials():
#         arcpy.AddError("BCGW credentials could not be established.")
#         sys.exit()
#     sde = os.getenv("SDE_FILE_PATH")
#     arcpy.AddMessage(f"sde file location: {sde}")

# except Exception as e:
#     arcpy.AddError(f"Failure occurred when establishing BCGW connection - {e}. Please try again.")
#     sys.exit()

#Check RAAD connection
raad = os.path.join(sde, "WHSE_ARCHAEOLOGY.RAAD_TFM_SITE")
try:
    arcpy.MakeFeatureLayer_management(raad, "RAAD_lyr")
except arcpy.ExecuteError as e:
    arcpy.AddWarning(f"Unable to connect to RAAD data")

arcpy.AddMessage("======================================================================")


#===========================================================================
# Populate the overlay criteria list.  This list will be passed in by the
# tool interface if your are running it that way.
#===========================================================================
revolt_criteria_to_pass = []
revolt_criteria_to_pass.append(input_feature_class)  
revolt_criteria_to_pass.append(sub_reports_on_this_field)
revolt_criteria_to_pass.append(xls_file1_overlap_to_run_from_dropdown_list) 
revolt_criteria_to_pass.append(xls_file1_overlap_to_run_from_user_specified)   
revolt_criteria_to_pass.append(report_header_line1)
revolt_criteria_to_pass.append(report_header_line2)     
revolt_criteria_to_pass.append(report_header_line3)     
revolt_criteria_to_pass.append(report_header_line4)      
revolt_criteria_to_pass.append(sub_reports_on_seperate_sheets)      
revolt_criteria_to_pass.append(directory_to_store_output)        
revolt_criteria_to_pass.append(summary_fields_on_seperate_lines)      
revolt_criteria_to_pass.append(dont_overwrite_existing_data)
revolt_criteria_to_pass.append(xls_file2_overlap_to_add_for_autostatus_region)
revolt_criteria_to_pass.append(red_report_header_disclaimer)
revolt_criteria_to_pass.append(suppress_map_creation)
revolt_criteria_to_pass.append(region)
revolt_criteria_to_pass.append(crown_file_number)
revolt_criteria_to_pass.append(disposition_number)
revolt_criteria_to_pass.append(parcel_number)
revolt_criteria_to_pass.append(run_as_fcbc)
revolt_criteria_to_pass.append(add_maps_to_current)
#------------------------------------------------------------------------------   


arcpy.AddMessage('About to launch the revolt tool')
#===============================================================================
# Call the universal overlap tool 
#===============================================================================
# Call the Revolt_Universal_Overlap_Tool 
arcpy.AddMessage("---------------------------------------------------")
arcpy.AddMessage("Passing Control on to Revolt Universal Overlap Tool")
arcpy.AddMessage("---------------------------------------------------")

overlapObj = revolt.revolt_tool()
overlapObj.run_revolt_tool(revolt_criteria_to_pass)
#------------------------------------------------------------------------------ 

# try:
#     if directory_to_store_output != "#" and directory_to_store_output != "":
#         project_path = 'explorer  ' + directory_to_store_output
#     elif output_dir_same_as_input != "false":
#         project_path = 'explorer  ' + os.path.dirname(input_feature_class)
#     else:
#         current_aprx = arcpy.mp.ArcGISProject("CURRENT")
#         home_folder = current_aprx.homeFolder
#         project_path = 'explorer  ' + home_folder
#     subprocess.Popen(project_path)
# except:
#     pass
'''
DeleteASTFiles_arcpro.py

Description:  Delete Automated Status Tool generated geodatabases and
              associated Excel and ILRR files (if present), from a
              directory and, optionally, all subdirectories.
              
Author:  Mike Eastwood
         Ministry of Forests, Lands and Natural Resource Operations
         250-774-5502
         michael.eastwood@gov.bc.ca
'''

# ================================================================================
# Preamble stuff - modules, environment, parameters and variables
# ================================================================================
# Import python modules
import arcpy
import os
import os.path

# Script arguments
folder = arcpy.GetParameterAsText(0)
subFolders = arcpy.GetParameter(1)

# Variables - Explicity name files and folders so there's no accidents
gdb1 = "one_status_common_datasets_aoi.gdb"
gdb2 = "one_status_tabs_1_and_2_datasets.gdb"
gdb3 = "aoi_boundary.gdb"
mapx = "mapx_files"
csv1 = "IlrrBusinessKeys.csv"
csv2 = "IlrrInterestHolders.csv"
csv3 = "IlrrInterests.csv"
csv4 = "IlrrLocations.csv"
csv5 = "Summary Report.csv"
txt = "readme.txt"
xls1 = "one_status_common_datasets_aoi.xlsx"
xls2 = "one_status_tabs_1_and_2.xlsx"
# ================================================================================

# ================================================================================
# Loop through and get the filenames and delete as necessary
# ================================================================================
# Generate a list of folders and files
arcpy.AddMessage(u"  \u200B \u200B  ")
arcpy.AddMessage("Deleting geodatabases and files . . . ")
# --------------------------------------------------------------------------------
if subFolders == True:
    # Cycle through the folders and delete as necessary
    #   Windows treats GDBs as folders
    walk = os.walk(folder)
    for (dirPath, dirNames, fileNames) in walk:
        for dirName in dirNames:
            if dirName in (gdb1, gdb2, gdb3, mapx):
                run = os.walk(os.path.join(dirPath, dirName))
                for path, dirs, files in run:
                    for file in files:
                        os.remove(os.path.join(path, file))
                os.rmdir(os.path.join(dirPath, dirName))
        for fileName in fileNames:
            if fileName in (csv1, csv2, csv3, csv4, csv5, txt, xls1, xls2):
                os.remove(os.path.join(dirPath, fileName))
# --------------------------------------------------------------------------------
else:
    # Only do the current folder
    dirList = os.listdir(folder)
    for item in dirList:
        if item in (gdb1, gdb2, gdb3, mapx):
            walk = os.walk(os.path.join(folder, item))
            for path, dirs, files in walk:
                for file in files:
                    os.remove(os.path.join(path, file))
            os.rmdir(os.path.join(folder, item))
        if item in (csv1, csv2, csv3, csv4, csv5, txt, xls1, xls2):
            os.remove(os.path.join(folder, item))
# --------------------------------------------------------------------------------
#arcpy.AddMessage(u"\u200B")
arcpy.AddMessage(u"\u200B\u0009\u0009\u0009\u0009. . . That was easy!")
arcpy.AddMessage(u"\u200B")
# ================================================================================
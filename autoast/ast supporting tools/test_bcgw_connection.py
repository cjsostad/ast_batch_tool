import arcpy
from database_connection import setup_bcgw
import os

# Get the sde path from database_connection.py
sde = setup_bcgw()
print("Database Connection (.sde) established at: ", sde)

# Set workspace to the created SDE path
arcpy.env.workspace = sde
print("Workspace explicitly set to: ", arcpy.env.workspace)

# List RAAD layers for quick check (optional)
arch_layers = arcpy.ListFeatureClasses("WHSE_ARCHAEOLOGY.*")
print("RAAD Layers found: ", arch_layers)

# Explicitly test connecting to the FWA Streams feature class
fwa_streams_fc = os.path.join(sde, "WHSE_BASEMAPPING.FWA_STREAM_NETWORKS_SP")
print("Attempting to access feature class: ", fwa_streams_fc)

# Test creating a feature layer and getting a count
try:
    arcpy.MakeFeatureLayer_management(fwa_streams_fc, "fwa_streams_lyr")
    print("Successfully created feature layer.")
    count = int(arcpy.GetCount_management("fwa_streams_lyr").getOutput(0))
    print(f"Feature layer contains {count} features.")
except Exception as e:
    print(f"Failed to create feature layer or get count: {e}")

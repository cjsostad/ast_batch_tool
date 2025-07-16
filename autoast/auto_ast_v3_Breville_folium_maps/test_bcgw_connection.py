import arcpy
from create_test_bcgw_connection import setup_bcgw
import os
from dotenv import load_dotenv


    # Load the default environment
load_dotenv()

# Call the setup_bcgw function to set up the database connection
# secrets = setup_bcgw(logger)
print("Attempting to set up BCGW connection")
secrets, sde_connection, sde_path = setup_bcgw()
print(f"Secrets: {secrets}")
print(f"SDE Connection: {sde_connection}")
print(f"SDE Path: {sde_path}")

# Get the sde path from database_connection.py
sde = sde_path
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

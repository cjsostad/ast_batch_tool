print("Testing arcpy import...")

try:
    import arcpy
    print("arcpy imported successfully")

    # Try a simple arcpy function
    print("Available licenses:")
    print(arcpy.CheckProduct("arcinfo"))  # Checks for ArcGIS Pro Advanced license

except ImportError as e:
    print("Failed to import arcpy:", e)

except Exception as e:
    print("arcpy is installed but an error occurred:", e)


import arcpy

print("Testing CreateDatabaseConnection tool...")

try:
    arcpy.management.CreateDatabaseConnection(
        out_folder_path=r"T:\Temp",  # Make sure this folder exists
        out_name="test_bcgw.sde",
        database_platform="ORACLE",
        instance="bcgw.bcgov",
        account_authentication="DATABASE_AUTH",
        username="csostad",  # Replace with a test username
        password="Larry2025",  # Replace with a test password
        save_user_pass="DO_NOT_SAVE_USERNAME"
    )
    print("CreateDatabaseConnection succeeded.")
except Exception as e:
    print("CreateDatabaseConnection failed:", e)
    
    print("Script Complete")



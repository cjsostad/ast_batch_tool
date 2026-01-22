"""
Script to convert KML files to shapefiles organized in individual folders.

This script:
1. Creates a 'shapefile' folder in the specified directory
2. For each KML file, creates a subfolder named after the file
3. Converts the KML to a shapefile with NAD 1983 BC Environment Albers projection
4. Saves the shapefile in its corresponding folder
"""

import arcpy
import os
import tempfile
import shutil


def convert_kml_to_shapefiles(source_folder):
    """
    Convert all KML files in a folder to shapefiles in organized subfolders.
    
    Args:
        source_folder (str): Path to folder containing KML files
    """
    # Enable overwrite
    arcpy.env.overwriteOutput = True
    
    # Create temporary directory for intermediate processing
    temp_base_dir = tempfile.mkdtemp(prefix="kml_conversion_")
    
    try:
        # Validate source folder exists
        if not os.path.exists(source_folder):
            print(f"Error: Source folder does not exist: {source_folder}")
            return
        
        # Create shapefile directory
        shapefile_dir = os.path.join(source_folder, "shapefile")
        if not os.path.exists(shapefile_dir):
            os.makedirs(shapefile_dir)
            print(f"Created shapefile directory: {shapefile_dir}")
        else:
            print(f"Shapefile directory already exists: {shapefile_dir}")
        
        # Define output coordinate system - NAD 1983 BC Environment Albers
        out_cs = arcpy.SpatialReference(3005)  # EPSG:3005
        
        # Find all KML files
        kml_files = [f for f in os.listdir(source_folder) 
                     if f.lower().endswith('.kml')]
        
        if not kml_files:
            print(f"No KML files found in {source_folder}")
            return
        
        print(f"\nFound {len(kml_files)} KML file(s) to process")
        print(f"Using temporary directory: {temp_base_dir}\n")
        
        # Process each KML file
        for kml_file in kml_files:
            temp_gdb = None
            try:
                # Get file name without extension
                file_name = os.path.splitext(kml_file)[0]
                
                # Create subfolder for this KML in final location
                output_folder = os.path.join(shapefile_dir, file_name)
                if not os.path.exists(output_folder):
                    os.makedirs(output_folder)
                    print(f"Created folder: {file_name}")
                
                # Full paths
                kml_path = os.path.join(source_folder, kml_file)
                
                print(f"  Converting {kml_file}...")
                
                # KMLToLayer will create a geodatabase in the source folder
                # This avoids temp directory issues
                temp_gdb_name = f"temp_{file_name}"
                arcpy.conversion.KMLToLayer(
                    kml_path,
                    source_folder,  # Create GDB in source folder (network path)
                    temp_gdb_name,
                    include_groundoverlay='NO_GROUNDOVERLAY'
                )
                
                temp_gdb = os.path.join(source_folder, f"{temp_gdb_name}.gdb")
                
                # Find the feature classes in the geodatabase
                arcpy.env.workspace = temp_gdb
                datasets = arcpy.ListDatasets(feature_type='feature')
                feature_classes = []
                
                # Get feature classes from datasets
                for dataset in datasets:
                    fcs = arcpy.ListFeatureClasses(feature_dataset=dataset)
                    for fc in fcs:
                        feature_classes.append(os.path.join(temp_gdb, dataset, fc))
                
                # Also get standalone feature classes
                standalone_fcs = arcpy.ListFeatureClasses()
                for fc in standalone_fcs:
                    feature_classes.append(os.path.join(temp_gdb, fc))
                
                if not feature_classes:
                    print(f"  Warning: No features found in {kml_file}")
                    continue
                
                # Create temp shapefile in local temp directory
                temp_shapefile = os.path.join(temp_base_dir, f"{file_name}.shp")
                
                # Merge all feature classes if there are multiple, otherwise just use the one
                if len(feature_classes) > 1:
                    arcpy.Merge_management(feature_classes, temp_shapefile)
                    input_for_project = temp_shapefile
                else:
                    input_for_project = feature_classes[0]
                
                # Project to NAD 1983 BC Environment Albers
                final_shapefile = os.path.join(temp_base_dir, f"{file_name}_final.shp")
                arcpy.Project_management(
                    input_for_project,
                    final_shapefile,
                    out_cs
                )
                
                # Copy all shapefile components to final location
                shapefile_base = os.path.splitext(final_shapefile)[0]
                for ext in ['.shp', '.shx', '.dbf', '.prj', '.cpg', '.sbn', '.sbx', '.shp.xml']:
                    src_file = shapefile_base + ext
                    if os.path.exists(src_file):
                        dst_file = os.path.join(output_folder, file_name + ext)
                        shutil.copy2(src_file, dst_file)
                
                print(f"  ✓ Created: {file_name}.shp")
                
            except Exception as e:
                print(f"  ✗ Error processing {kml_file}: {str(e)}")
                continue
            finally:
                # Clean up temp geodatabase from source folder
                if temp_gdb and os.path.exists(temp_gdb):
                    try:
                        arcpy.Delete_management(temp_gdb)
                    except:
                        pass
        
        print(f"\n✓ Conversion complete! Output location: {shapefile_dir}")
        
    except Exception as e:
        print(f"Error in conversion process: {str(e)}")
    
    finally:
        # Clean up temporary directory and all contents
        try:
            if os.path.exists(temp_base_dir):
                # Force delete - remove read-only attributes if needed
                for root, dirs, files in os.walk(temp_base_dir, topdown=False):
                    for name in files:
                        filepath = os.path.join(root, name)
                        try:
                            os.chmod(filepath, 0o777)
                            os.remove(filepath)
                        except:
                            pass
                    for name in dirs:
                        dirpath = os.path.join(root, name)
                        try:
                            os.chmod(dirpath, 0o777)
                            os.rmdir(dirpath)
                        except:
                            pass
                try:
                    shutil.rmtree(temp_base_dir, ignore_errors=True)
                except:
                    pass
                print(f"Cleaned up temporary directory")
        except Exception as e:
            print(f"Warning: Could not fully clean up temp directory: {str(e)}")


if __name__ == "__main__":
    # Folder containing KML files
    source_folder = r"\\spatialfiles.bcgov\srm\gss\projects\gr_2026_xx_skeena_replacements\source_data\Skeena 2026-2028 Shapefiles"
    
    print("=" * 70)
    print("KML to Shapefile Converter")
    print("=" * 70)
    print(f"Source folder: {source_folder}")
    print(f"Output coordinate system: NAD 1983 BC Environment Albers (EPSG:3005)")
    print("=" * 70 + "\n")
    
    convert_kml_to_shapefiles(source_folder)

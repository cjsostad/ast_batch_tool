"""
Test script to verify inactive_dispositions changes don't affect output

This script tests that the new SDE-based query method returns identical results 
to the old pyodbc connection method.

Usage:
    Run this from ArcGIS Pro Python environment:
    python test_inactive_dispositions.py
"""

import os
import sys
from dotenv import load_dotenv
import arcpy
import pandas as pd

# Add the script directory to path
current_path = os.path.dirname(os.path.realpath(__file__))
sys.path.append(current_path)

import inactive_dispositions as inactives
from database_connection import setup_bcgw
from logging_setup import setup_logging

def compare_dataframes(df1, df2, method1_name, method2_name):
    """Compare two dataframes and report differences"""
    
    print(f"\n{'='*80}")
    print(f"COMPARISON: {method1_name} vs {method2_name}")
    print(f"{'='*80}")
    
    # Shape comparison
    print(f"\n{method1_name} shape: {df1.shape}")
    print(f"{method2_name} shape: {df2.shape}")
    
    if df1.shape != df2.shape:
        print("❌ DIFFERENT SHAPES - Results do not match!")
        return False
    
    # Column comparison
    cols1 = set(df1.columns)
    cols2 = set(df2.columns)
    
    if cols1 != cols2:
        print(f"\n❌ DIFFERENT COLUMNS!")
        print(f"Only in {method1_name}: {cols1 - cols2}")
        print(f"Only in {method2_name}: {cols2 - cols1}")
        return False
    
    print(f"✓ Columns match: {list(df1.columns)}")
    
    # Sort both dataframes to ensure consistent comparison
    df1_sorted = df1.sort_values(by=df1.columns.tolist()).reset_index(drop=True)
    df2_sorted = df2.sort_values(by=df2.columns.tolist()).reset_index(drop=True)
    
    # Value comparison
    try:
        pd.testing.assert_frame_equal(df1_sorted, df2_sorted, check_dtype=False)
        print("\n✅ SUCCESS: DataFrames are identical!")
        print(f"   {len(df1)} rows match perfectly")
        return True
    except AssertionError as e:
        print(f"\n❌ FAILED: DataFrames differ!")
        print(f"Error: {str(e)}")
        
        # Show sample differences
        diff_mask = (df1_sorted != df2_sorted).any(axis=1)
        if diff_mask.any():
            print(f"\nFirst 5 differing rows:")
            print(df1_sorted[diff_mask].head())
            print("\nvs")
            print(df2_sorted[diff_mask].head())
        
        return False


def test_inactive_dispositions_methods():
    """Test both query methods with a small sample"""
    
    print("\n" + "="*80)
    print("TESTING INACTIVE DISPOSITIONS - OLD vs NEW METHOD")
    print("="*80)
    
    # Setup
    logger = setup_logging()
    load_dotenv()
    secrets, sde_connection, sde_path = setup_bcgw(logger)
    username, password = secrets[0], secrets[1]
    
    print(f"\nSDE Connection: {sde_path}")
    print(f"Username: {username}")
    
    # Get test data - clip small area to get a few parcels
    test_aoi = r"memory\test_aoi"
    
    # Create a small test polygon in Skeena region (adjust coordinates as needed)
    print("\nCreating test AOI...")
    sr = arcpy.SpatialReference(3005)  # BC Albers
    
    # Small box in Skeena region - adjust if needed
    point_list = [
        arcpy.Point(1163000, 1011000),
        arcpy.Point(1165000, 1011000),
        arcpy.Point(1165000, 1013000),
        arcpy.Point(1163000, 1013000),
    ]
    polygon = arcpy.Polygon(arcpy.Array(point_list), sr)
    arcpy.management.CopyFeatures(polygon, test_aoi)
    
    # Get parcel list
    print("Clipping parcels to test AOI...")
    parcel_fc = os.path.join(sde_path, r'WHSE_TANTALIS.TA_INTEREST_PARCEL_SHAPES')
    clip_parcel = arcpy.Clip_analysis(parcel_fc, test_aoi, r"memory\parcel_clip")
    result = int(arcpy.GetCount_management(clip_parcel).getOutput(0))
    
    print(f"Found {result} parcels in test area")
    
    if result == 0:
        print("\n⚠ WARNING: No parcels found in test area. Adjust test AOI coordinates.")
        print("Test cannot proceed without sample data.")
        return False
    
    # Limit to first 50 parcels to keep test quick
    parcel_list = [row[0] for row in arcpy.da.SearchCursor(clip_parcel, ['INTRID_SID'])][:50]
    print(f"Testing with {len(parcel_list)} parcels")
    
    # Get Oracle driver
    oracle_driver = inactives.get_oracle_driver()
    if not oracle_driver:
        print("❌ Oracle driver not found")
        return False
    
    print(f"Oracle driver: {oracle_driver}")
    
    # METHOD 1: Old method - pyodbc connection (no SDE parameter)
    print("\n" + "-"*80)
    print("METHOD 1: OLD (pyodbc connection)")
    print("-"*80)
    try:
        result_old = inactives.execute_process(parcel_list, username, password, oracle_driver, sde_connection=None)
        print(f"✓ Old method completed")
        print(f"  Returned {len(result_old.get('interest_status', []))} records")
    except Exception as e:
        print(f"❌ Old method failed: {e}")
        result_old = None
    
    # METHOD 2: New method - SDE connection
    print("\n" + "-"*80)
    print("METHOD 2: NEW (SDE connection)")
    print("-"*80)
    try:
        result_new = inactives.execute_process(parcel_list, username, password, oracle_driver, sde_connection=sde_path)
        print(f"✓ New method completed")
        print(f"  Returned {len(result_new.get('interest_status', []))} records")
    except Exception as e:
        print(f"❌ New method failed: {e}")
        result_new = None
    
    # Compare results
    if result_old is None or result_new is None:
        print("\n❌ TEST FAILED: One or both methods failed to execute")
        return False
    
    # Convert to DataFrames for comparison
    df_old = pd.DataFrame(result_old)
    df_new = pd.DataFrame(result_new)
    
    # Compare
    success = compare_dataframes(df_old, df_new, "OLD METHOD", "NEW METHOD")
    
    # Cleanup
    arcpy.Delete_management(test_aoi)
    arcpy.Delete_management(clip_parcel)
    
    return success


if __name__ == "__main__":
    try:
        success = test_inactive_dispositions_methods()
        
        if success:
            print("\n" + "="*80)
            print("✅ ALL TESTS PASSED - New method produces identical results!")
            print("="*80)
            sys.exit(0)
        else:
            print("\n" + "="*80)
            print("❌ TESTS FAILED - Results differ between methods")
            print("="*80)
            sys.exit(1)
            
    except Exception as e:
        print(f"\n❌ TEST ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

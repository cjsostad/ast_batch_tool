# region_detector.py
# Batch NR Region detection for the AutoAST batch tool.
#
# Given a list of shapefile paths and a BCGW SDE connection path, runs a single
# SpatialJoin against the NR Regions layer and returns the detected region name
# for each shapefile.
#
# Using a single Merge + SpatialJoin (rather than one join per shapefile) loads
# the BCGW layer only once, which is significantly faster for large batches.
#
# IMPORTANT: NR_REGION_FIELD must match the actual field name in the BCGW layer.
# If region values are coming back empty, check this constant first.
# Common candidates: "REGION_NAME", "ORG_UNIT_NAME", "REGION_NAM" (shapefile truncation).

import os
import collections

try:
    import arcpy  # type: ignore
except ImportError:
    arcpy = None  # Allow module import in non-ArcPy environments for testing

# BCGW layer containing NR Region boundaries
NR_REGION_LAYER = "WHSE_ADMIN_BOUNDARIES.ADM_NR_REGIONS_SPG"

# Field in NR_REGION_LAYER that holds the human-readable region name.
NR_REGION_FIELD = "REGION_NAME"

# Maps the full BCGW REGION_NAME values to the short names the AST tool expects.
# The AST tool constructs its region-specific input spreadsheet filename from this
# value (e.g. "West_Coast" -> one_status_west_coast_specific.xlsx).
# If the BCGW returns a full name not in this map, the raw value is passed through
# and a warning is logged so the mismatch is visible.
REGION_NAME_MAP = {
    "Skeena Natural Resource Region":            "Skeena",
    "Northeast Natural Resource Region":         "Northeast",
    "Thompson-Okanagan Natural Resource Region": "Thompson_Okanagan",
    "Cariboo Natural Resource Region":           "Cariboo",
    "West Coast Natural Resource Region":        "West_Coast",
    "South Coast Natural Resource Region":       "South_Coast",
    "Omineca Natural Resource Region":           "Omineca",
    "Kootenay-Boundary Natural Resource Region": "Kootenay_Boundary",
}


def detect_regions(shp_paths, sde_path, logger):
    """
    Batch-detect the NR Region for each shapefile using a single SpatialJoin.

    All shapefiles are merged into one in_memory feature class with a SRC_IDX
    integer field tracking which shapefile each feature came from. A single
    SpatialJoin against the BCGW NR Regions layer is then run with
    match_option="LARGEST_OVERLAP" so that AOIs straddling a region boundary
    are assigned to the region with the greatest overlap area.

    Results are grouped by SRC_IDX and the most-common matched region name is
    returned for each shapefile. Multi-feature shapefiles (several polygons per
    file) are handled by majority vote across all their features.

    Parameters:
        shp_paths : list of Path
            The shapefiles to detect regions for, in order.
        sde_path : str
            Path to the bcgw.sde connection file.
        logger : logging.Logger
            For progress and warning messages.

    Returns:
        dict {int: str}
            Maps index (position in shp_paths) to the detected region name.
            Missing entries mean detection failed for that shapefile; the
            caller should treat a missing entry as an empty string.
    """
    if arcpy is None:
        raise RuntimeError("arcpy is not available — region detection requires ArcGIS Pro.")

    # Define temp FC names up front so the finally block can always reference them.
    # Use scratchGDB (a local file GDB ArcPy manages automatically) instead of in_memory:
    # on Jenkins, CopyFeatures from a .shp source to in_memory appended '.shp' to the output
    # name, producing 'in_memory\\ast_region_temp_0.shp' which in_memory rejects (ERROR 000354).
    # File GDB feature classes have no extension, so the issue cannot recur.
    scratch_gdb = arcpy.env.scratchGDB
    temp_fcs = []
    merged_fc = os.path.join(scratch_gdb, "ast_region_merged")
    join_fc = os.path.join(scratch_gdb, "ast_region_join")
    region_map = {}

    try:
        # Step 1: Copy each shapefile to in_memory and stamp it with its list index.
        # The SRC_IDX field lets us map join results back to the originating shapefile
        # after the merge collapses all features into a single FC.
        for i, shp_path in enumerate(shp_paths):
            # Write to scratchGDB — see comment above merged_fc/join_fc for rationale.
            temp_fc = os.path.join(scratch_gdb, f"ast_region_temp_{i}")
            arcpy.management.CopyFeatures(str(shp_path), temp_fc)
            arcpy.management.AddField(temp_fc, "SRC_IDX", "LONG")
            arcpy.management.CalculateField(temp_fc, "SRC_IDX", str(i), "PYTHON3")
            temp_fcs.append(temp_fc)
            logger.info(f"Region Detector: Prepared '{shp_path.name}' as temp FC index {i}")

        if not temp_fcs:
            logger.warning("Region Detector: No shapefiles to process — returning empty result.")
            return region_map

        # Step 2: Merge all per-shapefile temp FCs into one.
        # This lets us call SpatialJoin exactly once regardless of how many shapefiles there are.
        arcpy.management.Merge(temp_fcs, merged_fc)
        logger.info(f"Region Detector: Merged {len(temp_fcs)} temp FC(s) for batch spatial join")

        # Step 3: Single SpatialJoin against the BCGW NR Regions layer.
        # LARGEST_OVERLAP: when an AOI polygon straddles two regions, the region
        # covering the greater area is selected as the join match.
        # KEEP_ALL (outer join): preserves unmatched features so we can warn when
        # a shapefile falls entirely outside all NR Regions.
        nr_layer_path = os.path.join(sde_path, NR_REGION_LAYER)
        arcpy.analysis.SpatialJoin(
            target_features=merged_fc,
            join_features=nr_layer_path,
            out_feature_class=join_fc,
            join_operation="JOIN_ONE_TO_ONE",
            join_type="KEEP_ALL",
            match_option="LARGEST_OVERLAP"
        )
        logger.info(f"Region Detector: SpatialJoin complete against {NR_REGION_LAYER}")

        # Step 4: Read results and group by SRC_IDX.
        # Join_Count is the number of region polygons that overlapped the target feature
        # before the JOIN_ONE_TO_ONE reduction — > 1 indicates a straddle.
        raw_results = collections.defaultdict(list)
        with arcpy.da.SearchCursor(join_fc, ["SRC_IDX", NR_REGION_FIELD, "Join_Count"]) as cursor:
            for src_idx, region_name, join_count in cursor:
                raw_results[int(src_idx)].append((region_name, join_count or 0))

        # Step 5: Derive one region name per shapefile; warn on straddles and misses.
        for i, shp_path in enumerate(shp_paths):
            rows = raw_results.get(i, [])

            if not rows:
                logger.warning(
                    f"Region Detector: No spatial join result for '{shp_path.name}' — "
                    f"shapefile may lie outside all NR Regions. Leaving region blank."
                )
                region_map[i] = ""
                continue

            # Majority vote across all features in this shapefile (handles multi-part AOIs)
            region_counts = collections.Counter(r[0] for r in rows if r[0])
            if not region_counts:
                logger.warning(
                    f"Region Detector: '{NR_REGION_FIELD}' was empty for all features in "
                    f"'{shp_path.name}'. Verify NR_REGION_FIELD constant in region_detector.py."
                )
                region_map[i] = ""
                continue

            best_region_raw = region_counts.most_common(1)[0][0]

            # Straddle check: Join_Count > 1 means at least one feature overlapped multiple regions
            max_join_count = max((r[1] for r in rows), default=1)
            unique_regions = len(region_counts)
            if max_join_count > 1 or unique_regions > 1:
                all_regions = ", ".join(sorted(region_counts.keys()))
                logger.warning(
                    f"Region Detector: '{shp_path.name}' straddles multiple NR regions "
                    f"({all_regions}) — largest overlap region '{best_region_raw}' was used."
                )

            # Translate the full BCGW name to the short name the AST tool expects.
            # e.g. "West Coast Natural Resource Region" -> "West_Coast"
            if best_region_raw in REGION_NAME_MAP:
                best_region = REGION_NAME_MAP[best_region_raw]
            else:
                # Unknown BCGW value — pass raw through and warn so the mismatch is visible
                best_region = best_region_raw
                logger.warning(
                    f"Region Detector: '{shp_path.name}' — BCGW region name '{best_region_raw}' "
                    f"is not in REGION_NAME_MAP. Passing raw value through; AST tool may fail. "
                    f"Update REGION_NAME_MAP in region_detector.py if this is a new region."
                )

            region_map[i] = best_region
            logger.info(f"Region Detector: '{shp_path.name}' -> '{best_region}' (BCGW: '{best_region_raw}')")

        return region_map

    finally:
        # Best-effort cleanup of all scratchGDB feature classes created during detection.
        # Using a finally block ensures cleanup runs even if an exception is raised above.
        for fc in temp_fcs + [merged_fc, join_fc]:
            try:
                if arcpy.Exists(fc):
                    arcpy.management.Delete(fc)
            except Exception:
                pass  # Don't mask the real error with a cleanup failure

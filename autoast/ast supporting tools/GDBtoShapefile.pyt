import arcpy
import re
from openpyxl import Workbook
from pathlib import Path

def to_posix(path):
    """Convert Path or str to POSIX forward-slash form."""
    return path.as_posix() if isinstance(path, Path) else str(path).replace("\\", "/")

def sanitize_filename(name) -> str:
    """Replace invalid characters with underscore and strip extra spaces."""
    name = re.sub(r'[\\/*?:"<>|]', '_', str(name))
    name = re.sub(r'\s+', '_', name)  # replace spaces with underscores
    return name[:12]  # shapefile name limit

# Hardcoded region layer path
REGION_LAYER = Path(r"\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\orahimi\Projects\Overlay Analysis\Natural Resource Regions.shp")
REGION_FIELD = "REGION_NAM"

class Toolbox(object):
    def __init__(self):
        self.label = "Unified Toolbox"
        self.alias = "Unified"
        self.tools = [UnifiedProcessor]

class UnifiedProcessor(object):
    def __init__(self):
        self.label = "Unified Region Processor"
        self.description = "Either process a folder of shapefiles (Case 1) or export features by field (Case 2)"
        self.canRunInBackground = False

    def getParameterInfo(self):
        mode_param = arcpy.Parameter(
            displayName="Mode",
            name="mode",
            datatype="String",
            parameterType="Required",
            direction="Input"
        )
        mode_param.filter.type = "ValueList"
        mode_param.filter.list = ["Case 1 - Folder", "Case 2 - By Field"]

        input_folder = arcpy.Parameter(
            displayName="Input Folder (for Case 1)",
            name="input_folder",
            datatype="DEWorkspace",
            parameterType="Optional",
            direction="Input"
        )

        input_layer = arcpy.Parameter(
            displayName="Input Feature Layer (for Case 2)",
            name="input_layer",
            datatype="DEFeatureClass",
            parameterType="Optional",
            direction="Input"
        )

        query_field = arcpy.Parameter(
            displayName="Query Field (for Case 2)",
            name="query_field",
            datatype="Field",
            parameterType="Optional",
            direction="Input"
        )
        query_field.parameterDependencies = [input_layer.name]

        return [mode_param, input_folder, input_layer, query_field]

    def execute(self, parameters, messages):
        mode = parameters[0].valueAsText
        input_folder = parameters[1].valueAsText
        input_layer = parameters[2].valueAsText
        query_field = parameters[3].valueAsText if parameters[3].value else None

        if mode == "Case 1 - Folder" and input_folder:
            self.run_case_1(Path(input_folder))
        elif mode == "Case 2 - By Field" and input_layer and query_field:
            self.run_case_2(Path(input_layer), query_field)
        else:
            arcpy.AddError("Invalid parameters for selected mode.")

    def run_case_1(self, input_path):
        outputs_root = input_path / "outputs"
        outputs_root.mkdir(parents=True, exist_ok=True)

        headers = [
            'region', 'feature_layer', 'crown_file_number', 'disposition_number',
            'parcel_number', 'output_directory', 'output_directory_same_as_input',
            'dont_overwrite_outputs', 'skip_conflicts_and_constraints',
            'suppress_map_creation', 'add_maps_to_current', 'run_as_fcbc',
            'ast_condition', 'file_number'
        ]

        wb = Workbook()
        ws = wb.active
        ws.title = "ast_config"
        ws.append(headers)

        shapefile_count = 0
        skipped_count = 0
        file_index = 1
        record_in_file = 0

        for shp_path in input_path.rglob("*.shp"):
            shp_name = shp_path.stem
            temp_join = Path("in_memory") / f"join_{shp_name}"
            arcpy.AddMessage(f"📂 Processing: {shp_name}")
            try:
                arcpy.analysis.SpatialJoin(
                    target_features=to_posix(shp_path),
                    join_features=to_posix(REGION_LAYER),
                    out_feature_class=to_posix(temp_join),
                    join_type="KEEP_COMMON",
                    match_option="INTERSECT"
                )

                result_count = int(arcpy.management.GetCount(to_posix(temp_join))[0])
                if result_count == 0:
                    arcpy.AddWarning(f"⚠️ Skipping {shp_name}: no region found")
                    skipped_count += 1
                    continue

                with arcpy.da.SearchCursor(to_posix(temp_join), [REGION_FIELD]) as cursor:
                    region = next(cursor)[0]

                output_subfolder = outputs_root / shp_name
                output_subfolder.mkdir(parents=True, exist_ok=True)

                ws.append([
                    region, to_posix(shp_path), None, None, None, to_posix(output_subfolder),
                    str(False), str(False), str(False), str(False),
                    str(False), str(True), None, None
                ])
                shapefile_count += 1
                record_in_file += 1

                if record_in_file == 8:
                    excel_file = outputs_root / f"Batch_Jobs_part{file_index}.xlsx"
                    wb.save(to_posix(excel_file))
                    arcpy.AddMessage(f"📝 Saved {excel_file}")
                    file_index += 1
                    record_in_file = 0
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "ast_config"
                    ws.append(headers)

            except Exception as e:
                arcpy.AddError(f"❌ Error with {shp_name}: {e}")
                continue

        if record_in_file > 0:
            excel_file = outputs_root / f"Batch_Jobs_part{file_index}.xlsx"
            wb.save(to_posix(excel_file))
            arcpy.AddMessage(f"📝 Saved {excel_file}")

        arcpy.AddMessage(f"✅ Done! {shapefile_count} shapefile(s) processed, {skipped_count} skipped.")

    def run_case_2(self, input_layer, query_field):
        arcpy.env.overwriteOutput = True
        field_names = [f.name for f in arcpy.ListFields(to_posix(input_layer))]
        if query_field not in field_names:
            arcpy.AddError(f"Field '{query_field}' not found.")
            return

        unique_values = set()
        with arcpy.da.SearchCursor(to_posix(input_layer), [query_field]) as cursor:
            for row in cursor:
                if row[0] not in (None, ''):
                    unique_values.add(str(row[0]))

        if not unique_values:
            arcpy.AddWarning(f"No valid values found in field '{query_field}'.")
            return

        # Determine if input is inside a .gdb
        if ".gdb" in to_posix(input_layer).lower():
            outputs_dir = Path(input_layer).parents[1] / "outputs"
            arcpy.AddMessage(f"⚠️ Input is inside a GDB, outputs placed outside at: {outputs_dir}")
        else:
            outputs_dir = input_layer.parent / "outputs"

        outputs_dir.mkdir(parents=True, exist_ok=True)

        wb = Workbook()
        ws = wb.active
        ws.title = "ast_config"

        headers = [
            'region', 'feature_layer', 'crown_file_number', 'disposition_number',
            'parcel_number', 'output_directory', 'output_directory_same_as_input',
            'dont_overwrite_outputs', 'skip_conflicts_and_constraints',
            'suppress_map_creation', 'add_maps_to_current', 'run_as_fcbc',
            'ast_condition', 'file_number'
        ]
        ws.append(headers)

        file_index = 1
        record_in_file = 0

        for value in sorted(unique_values):
            safe_value = sanitize_filename(value)
            query = f"{arcpy.AddFieldDelimiters(to_posix(input_layer), query_field)} = '{value}'"
            arcpy.AddMessage(f"🔹 Processing: {value}")

            feature_output_dir = outputs_dir / safe_value
            feature_output_dir.mkdir(parents=True, exist_ok=True)

            shp_out = feature_output_dir / f"{safe_value}.shp"
            kml_out = feature_output_dir / f"{safe_value}.kmz"
            lyr_name = f"lyr_{safe_value}"

            if arcpy.Exists(lyr_name):
                arcpy.management.Delete(lyr_name)

            arcpy.management.MakeFeatureLayer(to_posix(input_layer), lyr_name, query)
            arcpy.management.CopyFeatures(lyr_name, to_posix(shp_out))
            arcpy.conversion.LayerToKML(lyr_name, to_posix(kml_out))

            region_value = None
            try:
                join_output = Path("in_memory") / f"join_{safe_value}"
                arcpy.analysis.SpatialJoin(
                    target_features=to_posix(shp_out),
                    join_features=to_posix(REGION_LAYER),
                    out_feature_class=to_posix(join_output),
                    join_type="KEEP_COMMON",
                    match_option="INTERSECT"
                )
                with arcpy.da.SearchCursor(to_posix(join_output), [REGION_FIELD]) as cursor:
                    region_value = next(cursor)[0]

                arcpy.management.Delete(to_posix(join_output))
            except Exception as e:
                arcpy.AddWarning(f"⚠️ Spatial join failed for {value}: {e}")

            ws.append([
                region_value, to_posix(shp_out), None, None, None, to_posix(feature_output_dir),
                str(False), str(False), str(False), str(False),
                str(False), str(True), None, None
            ])
            arcpy.management.Delete(lyr_name)

            record_in_file += 1
            if record_in_file == 8:
                excel_file = outputs_dir / f"roads_status_sheet_part{file_index}.xlsx"
                wb.save(to_posix(excel_file))
                arcpy.AddMessage(f"📝 Saved {excel_file}")
                file_index += 1
                record_in_file = 0
                wb = Workbook()
                ws = wb.active
                ws.title = "ast_config"
                ws.append(headers)

        if record_in_file > 0:
            excel_file = outputs_dir / f"roads_status_sheet_part{file_index}.xlsx"
            wb.save(to_posix(excel_file))
            arcpy.AddMessage(f"📝 Saved {excel_file}")

        arcpy.AddMessage(f"✅ Excel saved with {file_index} part(s) at: {to_posix(outputs_dir)}")

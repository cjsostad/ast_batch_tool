'''
This script generates interavtive HTML maps.

Update: May 14, 2024.
'''
import os
import re
import warnings
warnings.simplefilter(action='ignore')

import os
import time
import logging
import base64
import numpy as np
import pandas as pd
import geopandas as gpd
import fiona
import shapely.wkt as wkt
from pathlib import Path
import folium
from folium.plugins import MeasureControl, MousePosition,FloatImage, MiniMap, Search, GroupedLayerControl
from branca.element import Template, MacroElement
import mapstyle
import psutil

LOGO = (r"\\spatialfiles.bcgov\work\lwbc\visr\Workarea\moez_labiadh\MAPS\Logos\BCID_V_key_pms_pos_small_150px.JPG")

class HTMLGenerator:
    def __init__(self, common_xls, region_xls, status_gdb, out_location):
        self.common_xls = common_xls
        self.region_xls = region_xls
        self.status_gdb = status_gdb
        self.out_loc = out_location
        self._setup_logger()

    def _setup_logger(self):
        log_name = "html_time.log"
        log_path = Path(self.out_loc) / log_name
        log_path.parent.mkdir(parents=True, exist_ok=True)

        # Create a class-specific logger
        self.logger = logging.getLogger(self.__class__.__name__)
        self.logger.setLevel(logging.INFO)

        # Avoid adding duplicate handlers
        if not self.logger.handlers:
            file_handler = logging.FileHandler(log_path)
            file_handler.setFormatter(logging.Formatter(
                "%(asctime)s - %(levelname)s - %(message)s"
            ))

            stream_handler = logging.StreamHandler()
            stream_handler.setFormatter(logging.Formatter(
                "%(asctime)s - %(levelname)s - %(message)s"
            ))

            self.logger.addHandler(file_handler)
            self.logger.addHandler(stream_handler)

    @staticmethod
    def parse_args():
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--xls1", required=True, help="Path to the common xls file")
        parser.add_argument("--xls2", required=True, help="Path to the region xls file")
        parser.add_argument("--gdb", required=True, help="Path to the status gdb")
        parser.add_argument("--out", required=True, help="Path to the output location")
        return parser.parse_args()

    @classmethod
    def run_from_cli(cls):
        args = cls.parse_args()
        cls.run_from_args(vars(args))

    @classmethod
    def run_from_args(cls, args_dict):
        instance = cls(
            common_xls=args_dict["xls1"],
            region_xls=args_dict["xls2"],
            status_gdb=args_dict["gdb"],
            out_location=args_dict["out"]
        )
        instance.generate_html_maps()

    def timed_step(self, label):
        process = psutil.Process(os.getpid())
        self.logger.info(f"[TIMER] Starting: {label}")

        start = time.perf_counter()
        start_mem = process.memory_info().rss / 1024 ** 2
        start_cpu = process.cpu_percent(interval=None) 

        def end():
            elapsed = time.perf_counter() - start
            end_mem = process.memory_info().rss / 1024 ** 2
            mem_diff = end_mem - start_mem
            cpu_usage = process.cpu_percent(interval=0.1)

            self.logger.info(
                f"[TIMER] Finished: {label} in {elapsed:.2f} sec | "
                f"Mem: {end_mem:.2f} MB ({mem_diff:+.2f} MB) | "
                f"CPU: {cpu_usage:.2f}%"
            )

        return end

    def get_input_xlsx(self):
        """returns a dataframe of status input xlsxs """
        df_stat_c = pd.read_excel(self.common_xls)
        df_stat_r = pd.read_excel(self.region_xls)
        
        df_stat = pd.concat([df_stat_c, df_stat_r])
        df_stat.dropna(how='all', inplace=True)
    
        df_stat['Category'].fillna(method='ffill', inplace=True)
        
        df_stat = df_stat.reset_index(drop=True)
        
        return df_stat


    def create_map_template(self, Ymin=0, Xmin=0, Ymax=0, Xmax=0):
        """Returns an empty folium map object"""
        # Create a map object
        map_obj = folium.Map(control_scale=True)
        
        # Add GeoBC basemap to the map
        wms_url = 'https://maps.gov.bc.ca/arcgis/rest/services/province/web_mercator_cache/MapServer/tile/{z}/{y}/{x}'
        wms_attribution = 'GeoBC, DataBC, TomTom, © OpenStreetMap Contributors'
        folium.TileLayer(
            tiles=wms_url,
            name='GeoBC Basemap',
            attr=wms_attribution,
            overlay=False,
            control=True,
            transparent=True).add_to(map_obj)
        
        # Add a satellite basemap to the map
        satellite_url = 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}'
        satellite_attribution = 'Tiles &copy; Esri'
        folium.TileLayer(
            tiles=satellite_url,
            name='Imagery Basemap',
            attr=satellite_attribution,
            overlay=False,
            control=True).add_to(map_obj)
        
        # create a title and button to refresh the map to the initial view
        map_var_nme = map_obj.get_name()
        
        title_refresh = """
        <div style="position: fixed; 
            top: 88px; left: 70px; width: 150px; height: 70px; 
            background-color:transparent; border:0px solid grey;z-index: 900;">
            <button style="font-weight:bold; color:#DE1610" 
                    onclick="{}.fitBounds([[{}, {}], [{}, {}]])">Refresh View</button>
        </div>
    
        """.format(map_var_nme, Ymin, Xmin, Ymax, Xmax)
        
        map_obj.get_root().html.add_child(folium.Element(title_refresh))
        
        # Add measure controls to the map
        map_obj.add_child(MeasureControl(primary_length_unit='meters', 
                                        secondary_length_unit='kilometers',
                                        primary_area_unit='hectares'))
        
        # Add mouse position to the map
        MousePosition().add_to(map_obj)
    
        # Add Lat/Long Popup to the map
        map_obj.add_child(folium.features.LatLngPopup())
    
        # b64_content = base64.b64encode(open(LOGO, 'rb').read()).decode('utf-8')
        # float_image = FloatImage('data:image/png;base64,{}'.format(b64_content), bottom=3, left=2)
        # float_image.add_to(map_obj)
        
        #Add a Mini Map
        minimap = MiniMap(position="bottomleft")
        map_obj.add_child(minimap)

        # Offset Leaflet controls from the bottom to avoid the footer
        minimap_offset_css = """
        <style>
            .leaflet-bottom.leaflet-left {
                margin-bottom: 80px;  /* Adjust this based on your footer height */
            }
        </style>
        """
        map_obj.get_root().html.add_child(folium.Element(minimap_offset_css))
    
        # Add custom css style to the map
        app_css = mapstyle.map_css
        style = MacroElement()
        style._template = Template(app_css)
        map_obj.get_root().add_child(style)
        
        return map_obj
    
    
    def inject_branding(self, map_path, title="Feature Class"):
        """
        Injects BC Sans font, a header with a logo, and a footer into a saved Folium HTML map.

        Parameters:
        - map_path (Path or str): Path to the Folium-generated HTML map file.
        - logo_base64 (str): Base64 string of the logo image (e.g., from image_to_base64()).
        - title (str): Title text to display in the header.
        """

        def image_to_base64(image_path):
            with open(image_path, "rb") as img_file:
                encoded = base64.b64encode(img_file.read()).decode("utf-8")
            return f"data:image/png;base64,{encoded}"
        

        map_path = Path(map_path)
        if not map_path.exists():
            raise FileNotFoundError(f"Map file not found: {map_path}")

        # Read original HTML
        with map_path.open('r', encoding='utf-8') as f:
            html = f.read()

        # Inject BC Sans font
        bc_sans_link = '<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/bcgov/bc-sans@main/css/BCSans.css" />'
        html = html.replace('<head>', f'<head>\n    {bc_sans_link}')

        # Get the LOGO and format
        logo_base64 = image_to_base64(LOGO)

        # Define header and footer HTML
        header_html = f"""
        <div style="background:#003366;color:white;padding:10px 20px;
                    display:flex;align-items:center;font-family:'BC Sans', sans-serif;">
            <img src="{logo_base64}" alt="Logo" style="height:50px;margin-right:15px;">
            <div style="display:flex; flex-direction:column;">
                <div style="font-size:18px; font-weight:bold; color:white;">
                    Water, Land, and Resource Stewardship | Front Counter BC
                </div>
                <div style="font-size:22px; font-weight:bold; color:#e3a82b;">
                    {title}
                </div>
            </div>
        </div>
        """

        footer_html = """
        <div style="background:#003366;color:white;padding:10px;text-align:center;
                    font-size:14px;font-family:'BC Sans', sans-serif;position:absolute;
                    bottom:0;width:100%;z-index:9999;">
            © 2025 Government of British Columbia | INTERNAL GOVERNMENT USE ONLY
        </div>
        </body>"""

        # Inject into <body>
        html = html.replace('<body>', f'<body>\n{header_html}')
        html = html.replace('</body>', footer_html)

        # Write modified HTML back to file
        with map_path.open('w', encoding='utf-8') as f:
            f.write(html)

        print(f"Branding injected into: {map_path}")

    def add_multi_popup_handler(self, map_object, geojson_gdf, popup_fields):
        """
        Adds a dynamic Turf.js-based popup handler to a Folium map.
        Handles Polygons, Lines, and Points with scrollable, dynamic-width popups.

        Parameters:
        - map_object: folium.Map
        - geojson_gdf: GeoDataFrame with geometry + attributes
        - popup_fields: list of attribute fields to include in popup (required)
        """
        import json

        if not isinstance(popup_fields, list) or not popup_fields:
            raise ValueError("popup_fields must be a non-empty list of field names.")

        geojson_data = json.loads(geojson_gdf.to_json())
        fields_js = json.dumps(popup_fields)

        # Single popup: table format, scrollable, dynamic width
        single_popup_html = f"""
            var html = "<div style='max-height:350px; overflow-y:auto; max-width:100vw; width:auto; box-sizing:border-box;'>";
            html += "<div style='font-weight:bold; font-size:14px; margin-bottom:6px;'>" + f['map_title'] + " (1 overlap)</div>";
            html += "<table style='width:100%;'>";
            {fields_js}.forEach(function(key) {{
                html += "<tr><td style='padding-right: 6px; vertical-align: top;'><b>" + key + "</b></td><td>" + f[key] + "</td></tr>";
            }});
            html += "</table></div>";
        """

        # Multi popup: alternating color, line separators, scrollable, dynamic width
        multi_popup_html = f"""
            if (overlapping.length > 0) {{
                var headerTitle = overlapping[0]['map_title'] || 'Feature';
                html += "<div style='font-weight:bold; font-size:14px; margin-bottom:6px;'>" + headerTitle + " (" + overlapping.length + " overlaps)</div>";
            }}
            html += "<div style='max-height:350px; overflow-y:auto; max-width:100vw; width:auto; box-sizing:border-box;'>";
            overlapping.forEach(function(f, index) {{
                if (index > 0) html += "<hr style='border-top: 2px solid #003366; margin: 8px 0;'>";
                var color = (index % 2 === 0) ? '#000000' : '#003366';
                html += "<table style='width:100%; color:" + color + "'>";
                {fields_js}.forEach(function(key) {{
                    html += "<tr><td style='padding-right: 6px; vertical-align: top;'><b>" + key + "</b></td><td>" + f[key] + "</td></tr>";
                }});
                html += "</table>";
            }});
            html += "</div>";
        """

        # Inject the Turf.js logic
        js_code = f"""
        <script>
        {{% raw %}}
        var geojson = {json.dumps(geojson_data)};
        document.addEventListener("DOMContentLoaded", function() {{
            var map = {map_object.get_name()};
            map.on('click', function(e) {{
                var pt = turf.point([e.latlng.lng, e.latlng.lat]);
                var overlapping = [];

                geojson.features.forEach(function(feature) {{
                    var turfGeom = turf.feature(feature.geometry);
                    var type = feature.geometry.type;

                    var isMatch =
                        (type === 'Polygon' || type === 'MultiPolygon') &&
                            turf.booleanPointInPolygon(pt, turfGeom) ||

                        (type === 'LineString' || type === 'MultiLineString') &&
                            turf.booleanPointInPolygon(pt, turf.buffer(turfGeom, 5, {{ units: 'meters' }})) ||

                        (type === 'Point') &&
                            turf.distance(pt, turfGeom, {{ units: 'meters' }}) < 5;

                    if (isMatch) {{
                        overlapping.push(feature.properties);
                    }}
                }});

                if (overlapping.length === 1) {{
                    var f = overlapping[0];
                    {single_popup_html}
                    L.popup().setLatLng(e.latlng).setContent(html).openOn(map);
                }} else if (overlapping.length > 1) {{
                    var html = "";
                    {multi_popup_html}
                    L.popup().setLatLng(e.latlng).setContent(html).openOn(map);
                }}
            }});
        }});
        {{% endraw %}}
        </script>
        """
        return js_code


    def get_safe_popup_fields(self, gdf, max_fields=5, exclude_columns=None):
        """
        Returns a list of safe popup fields from a GeoDataFrame.

        Parameters:
        - gdf: GeoDataFrame to extract fields from
        - max_fields: Max number of fields to return
        - exclude_columns: Optional list of column names to exclude (case-insensitive)

        Returns:
        - List of column names safe for use in popups
        """
        if exclude_columns is None:
            exclude_columns = []

        exclude_columns = [col.lower() for col in exclude_columns]

        safe_cols = []
        for col in gdf.columns:
            if col.lower() == 'geometry' or col.lower() in exclude_columns:
                continue

            # Check if the field has simple types
            sample = gdf[col].dropna().head(5)
            if sample.empty:
                safe_cols.append(col)
            elif all(isinstance(val, (str, int, float, bool)) for val in sample):
                safe_cols.append(col)

            if len(safe_cols) >= max_fields:
                break

        return safe_cols

    def is_pathlike(self, s):
        """
        Returns True if the input string looks like a file system path.
        """
        s = str(s).strip()
        return bool(re.match(r"^([a-zA-Z]:[\\\\/]|[\\\\/]{2}|/)", s))


    def generate_html_maps(self):
        """Creates a HTML map for each feature class in gdb"""
        self.logger.info("[INFO] Starting HTML map generation process.")

        print('\nReading input xlsxs')

        end_timer = self.timed_step("Reading input xlsxs")
        df_st= self.get_input_xlsx()
        end_timer()

        print ('\nPreparing Layers for mapping')
        # Read the AOI feature class into a gdf
        end_timer = self.timed_step("Reading AOI layer")
        gdf_aoi = gpd.read_file(filename=self.status_gdb, layer='aoi')
        end_timer()

        end_timer = self.timed_step("Creating all Buffers")
        aoi_buffers = (
            pd.to_numeric(df_st['Buffer_Distance'].astype(str).str.strip(), errors='coerce')
            .dropna()
            .astype(int)
            .unique()
        )
        aoi_buffers = np.append(aoi_buffers, 1000) if 1000 not in aoi_buffers else aoi_buffers
        aoi_buffers = np.sort(aoi_buffers)
        bf_gdfs = {}

        for buffer_distance in aoi_buffers:
            key = f'aoi_{buffer_distance}'
            if key not in bf_gdfs:
                print(f'  Creating buffer: {buffer_distance}')
                bf_gdfs[key] = gpd.GeoDataFrame(
                    geometry=gdf_aoi.buffer(buffer_distance),
                    crs=gdf_aoi.crs
                )
        end_timer()

        end_timer = self.timed_step("Getting total bounds")
        # Zoom the all-layers map to the AOI extent
        buf_df = bf_gdfs.get('aoi_1000')
        xmin,ymin,xmax,ymax = buf_df.to_crs(epsg=4326).total_bounds
        end_timer()

        end_timer = self.timed_step("Create Map Template and fit to bounding box")

        map_all = self.create_map_template(Ymin=ymin, Xmin=xmin, Ymax=ymax, Xmax=xmax)
        
        map_all.fit_bounds([[ymin, xmin], [ymax, xmax]])
        end_timer()

        end_timer = self.timed_step("Add AOI to All-layers Map")
        # Add the AOI layer to the all-layers map
        grp_aoi= folium.FeatureGroup(name='AOI')  
        lyr_aoi= folium.GeoJson(data=gdf_aoi, name='AOI',
                    style_function=lambda x:{'color': 'red', 
                                                'fillColor': 'none',
                                                'weight': 3})
        lyr_aoi.add_to(grp_aoi)
        grp_aoi.add_to(map_all)
        
        aoi_grps= [grp_aoi]
        end_timer()

        end_timer = self.timed_step("Add Buffers to All-layers Map")
        # Add buffered areas to the all-layers map
        for k,v in bf_gdfs.items():
            grp_aoi_b= folium.FeatureGroup(name= k.upper()+' m', show=False)  
            lyr_aoi_b= folium.GeoJson(data=v, name=k,
                            style_function=lambda x:{'color': 'orange',
                                                     'fillColor': 'none',
                                                     'weight': 3})
            lyr_aoi_b.add_to(grp_aoi_b)
            grp_aoi_b.add_to(map_all)
        
            aoi_grps.append(grp_aoi_b)

        end_timer()

        print ('\nGenerating Individual Maps')
        end_timer = self.timed_step("Acquire Categories and list of feature classes")
        ctg_list= list(df_st['Category'].unique())
        ctg_list.insert(0, 'Area of Interest')
        
        ctg_grps=[aoi_grps]

        fc_list= fiona.listlayers(self.status_gdb)
        end_timer()
        
        for ctg in ctg_list:
            print (f'\nGenerating Maps for {ctg}')
            
            df = df_st.loc[df_st['Category'] == ctg] # Fil
            fc_grps= []
            counter= 1
            for i, row in df.iterrows():
                # Check if dealing with UNC path to feature class or table name in BCGW.
                raw_source = str(row['Datasource']).strip()
                source = Path(raw_source) if self.is_pathlike(raw_source) else raw_source

                fc = row['Featureclass_Name(valid characters only)']
                fc = fc.replace(" ", "_")
        
                print (f"..creating Map {counter} of {len(df)}: {fc}")
                if fc in fc_list:
                    end_timer = self.timed_step(f"Acquire geodataframe for {fc}")
                    gdf_fc = gpd.read_file(filename=self.status_gdb, layer=fc)
                    end_timer()
            
                    if gdf_fc.shape[0] > 0:
                        #Flatten 3D geometries to 2D (Folium doesn't like 3D)
                        end_timer = self.timed_step(f"Acquire geometry for {fc}") 
                        if gdf_fc['geometry'].has_z.any():
                            gdf_fc['geometry'] = gdf_fc['geometry'].apply(
                                        lambda geom: wkt.loads(
                                            wkt.dumps(geom, output_dimension=2)))
                        end_timer()

                        end_timer = self.timed_step(f"Convert Columns to strings") 
                        #convert all cols to str except geometry
                        for col in gdf_fc.columns:
                            if col != 'geometry':
                                gdf_fc[col] = gdf_fc[col].astype(str)
                        end_timer()
                            
                        # Set label column. Will be used for tooltip and legend.
                        map_title = fc.replace('_', ' ')
                        print(f"   generating html map for {map_title}")

                        end_timer = self.timed_step(f"Acquire field data from spreadsheet; build out fields to summarize for popups") 
                        #Get the row for the current feature class
                        df_item = df_st.loc[df_st['Featureclass_Name(valid characters only)'] == map_title]

                        if df_item.empty:
                            print(f"Warning: No rows match the filter condition for {map_title}.")
                            label_col = None
                            map_label_is_null = True
                            fields_to_summarize_is_null = True
                            continue

                        # Check the fields explicitly
                        map_label_field = df_item['map_label_field'].iloc[0] if 'map_label_field' in df_item else None
                        fields_to_summarize = df_item['Fields_to_Summarize'].iloc[0] if 'Fields_to_Summarize' in df_item else None

                        # Flags for later reference
                        map_label_is_null = pd.isnull(map_label_field)
                        fields_to_summarize_is_null = pd.isnull(fields_to_summarize)

                        # Decide on label_col
                        if not map_label_is_null:
                            label_col = map_label_field
                        elif not fields_to_summarize_is_null:
                            label_col = fields_to_summarize
                        else:
                            label_col = gdf_fc.columns[0]  # fallback
                            
                        # Set pop up columns
                        popup_cols = []
                        
                        # first_field = df_item['Fields_to_Summarize'].iloc[0]
                        first_field = df_item['Fields_to_Summarize'].iloc[0] if not df_item.empty and 'Fields_to_Summarize' in df_item.columns else None
                        if pd.notnull(first_field):
                            popup_cols.append(str(first_field.strip()))
                            
                        for f in range (2,7):
                            for i in df_item['Fields_to_Summarize' + str(f)].tolist():
                                if pd.notnull(i):
                                    popup_cols.append(str(i.strip()))
                        end_timer()

                        end_timer = self.timed_step(f"Get all popup fields and format") 
                        if len(popup_cols) == 0:
                            popup_cols = self.get_safe_popup_fields(gdf_fc, max_fields=5)

                        # Format the popup columns for better visualization
                        for col in popup_cols:
                            gdf_fc[col] = gdf_fc[col].astype(str)
                            gdf_fc[col] = gdf_fc[col].str.wrap(width=20).str.replace('\n','<br>')
                        end_timer()

                        end_timer = self.timed_step(f"Determine all values and symbology to show in legend and map.") 
                        unique_values = gdf_fc[label_col].unique()
                        label_count = len(unique_values)

                        if label_count > 20 or (map_label_is_null and fields_to_summarize_is_null):
                            # Lump all into one color
                            lumped_color = "#000000"
                            gdf_fc['__group__'] = map_title
                            gdf_fc['color'] = lumped_color
                            legend_labels = [(lumped_color, map_title)]
                            legend_header = "Legend"

                        else:
                            # Group and assign unique colors
                            gdf_fc['__group__'] = gdf_fc[label_col].fillna("").astype(str)
                            unique_values = gdf_fc['__group__'].unique()

                            color_mapping = {
                                value: f"#{np.random.randint(0, 255):02X}"
                                    f"{np.random.randint(0, 255):02X}"
                                    f"{np.random.randint(0, 255):02X}"
                                for value in unique_values
                            }

                            gdf_fc['color'] = gdf_fc['__group__'].map(color_mapping)
                            legend_labels = [(color, value) for value, color in color_mapping.items()]
                            legend_header = label_col
                        end_timer()

                        gdf_fc['color2'] = gdf_fc['color'].iloc[-1]

                        end_timer = self.timed_step(f"Get total bounds for the layer and set up map template.") 
                        # Zoom the map to the layer extent
                        gdf_fc = gdf_fc.to_crs(4326)
                        xmin, ymin, xmax, ymax = gdf_fc['geometry'].total_bounds
                            
                        # Create an individual map
                        map_one = self.create_map_template(Ymin=ymin, Xmin=xmin, Ymax=ymax, Xmax=xmax)

                        map_one.fit_bounds([[ymin, xmin], [ymax, xmax]])
                        end_timer()
                        
                        end_timer = self.timed_step(f"Create feature class layer and add to individual group and map") 
                        # Add the AOI layer to individual maps
                        grp_aoi_o= folium.FeatureGroup(name= 'AOI', show=True)  
                        lyr_aoi_o= folium.GeoJson(data=gdf_aoi, name='AOI',
                                    style_function=lambda x:{'color': 'red', 
                                                                'fillColor': 'none',
                                                                'weight': 3})
                        lyr_aoi_o.add_to(grp_aoi_o)
                        grp_aoi_o.add_to(map_one)
            
                        aoi_grps_o = [grp_aoi_o]
                        end_timer()
            
                        end_timer = self.timed_step(f"Create AOI layer and add to map") 
                        # Add buffered areas to individual maps
                        for k,v in bf_gdfs.items():
                            grp_aoi_b_o= folium.FeatureGroup(name= k.upper()+' m', show=False)  
                            lyr_aoi_b_o= folium.GeoJson(data=v, name=k,
                                            style_function=lambda x:{'color': 'orange',
                                                                    'fillColor': 'none',
                                                                    'weight': 3})
                            lyr_aoi_b_o.add_to(grp_aoi_b_o)
                            grp_aoi_b_o.add_to(map_one)
            
                            aoi_grps_o.append(grp_aoi_b_o)
                        end_timer()

                        end_timer = self.timed_step(f"Generate tooltip fields") 
                        # Create a list of columns for the tooltip
                        gdf_fc['map_title'] = map_title
                        gdf_fc['category'] = ctg
                        gdf_fc['data_source'] = [
                                                str(source.parent / source.name) if isinstance(source, Path) else str(source)
                                            ] * len(gdf_fc)
                        tooltip_cols = ['map_title', 'category', 'data_source']
                        end_timer()

                        end_timer = self.timed_step(f"Create a feature group in the layers list to house all unique values") 
                        # For each unique group value, create a separate layer
                        group_layers = []
                        for group_value in gdf_fc['__group__'].unique():
                            # Filter GeoDataFrame to this group
                            gdf_subset = gdf_fc[gdf_fc['__group__'] == group_value]

                            # Create a FeatureGroup for this group
                            grp_fc_o = folium.FeatureGroup(name=group_value, show=True)

                            # Add GeoJson layer with the color from the 'color' column
                            lyr_fc_o = folium.GeoJson(
                                data=gdf_subset,
                                name=group_value,
                                style_function=lambda x: {
                                    'fillColor': x['properties']['color'],
                                    'color': x['properties']['color'],
                                    'weight': 2
                                },
                                tooltip=folium.GeoJsonTooltip(fields=tooltip_cols,
                                                                aliases=['LAYER', 'CATEGORY', 'SOURCE'],
                                                                labels=True))

                            lyr_fc_o.add_to(grp_fc_o)
                            grp_fc_o.add_to(map_one)  # Add each group to the map
                            group_layers.append(grp_fc_o)
                        end_timer()
            
                        # # Add the layer to the individual map
                        # grp_fc_o= folium.FeatureGroup(name=map_title, show=True)  
                        # lyr_fc_o= folium.GeoJson(data=gdf_fc, name=map_title,
                        #             marker=folium.Circle(radius=5),
                        #             style_function= lambda x: {'fillColor': x['properties']['color'],
                        #                                         'color': x['properties']['color'],
                        #                                         'weight': 2},
                        #             tooltip=folium.features.GeoJsonTooltip(fields=tooltip_cols,
                        #                                                     aliases=['LAYER', label_col, 'CATEGORY', 'SOURCE'],
                        #                                                     labels=True))
                        # # Add layer to group and group to map
                        # lyr_fc_o.add_to(grp_fc_o)
                        # grp_fc_o.add_to(map_one)

                        # Add Turf.js to the map
                        end_timer = self.timed_step(f"Create popup for each layer with overlap reporting") 
                        turf_script = '<script src="https://unpkg.com/@turf/turf@6/turf.min.js"></script>'
                        map_one.get_root().html.add_child(folium.Element(turf_script))

                        js_code_one = self.add_multi_popup_handler(map_one, gdf_fc, popup_fields=popup_cols)
                        map_one.get_root().html.add_child(folium.Element(js_code_one))
   
                        # Add the layer to the all-Layers map
                        grp_fc_a= folium.FeatureGroup(name=map_title, show=False)
                        lyr_fc_a= folium.GeoJson(data=gdf_fc, name=map_title,
                                    marker=folium.Circle(radius=5),
                                    style_function= lambda x: {'fillColor': x['properties']['color2'],
                                                                'color': x['properties']['color2'],
                                                                'weight': 2},
                                    tooltip=folium.GeoJsonTooltip(fields=tooltip_cols,
                                                                            aliases=['LAYER', 'CATEGORY', 'SOURCE'],
                                                                            labels=True))
                        lyr_fc_a.add_to(grp_fc_a)
                        grp_fc_a.add_to(map_all)
                        fc_grps.append(grp_fc_a)
                        end_timer()

                        # map_all.get_root().html.add_child(folium.Element(turf_script))
                        # js_code_all = self.add_multi_popup_handler(map_all, gdf_fc, popup_fields=popup_cols)
                        # map_all.get_root().html.add_child(folium.Element(js_code_all))

                        end_timer = self.timed_step(f"Create legend HTML for individual maps") 
                        legend_html = '''
                            <div id="legend" style="position: fixed; 
                                bottom: 60px; right: 30px; z-index: 1000; 
                                background-color: #fff; padding: 10px; 
                                border-radius: 5px; border: 1px solid grey;
                                max-height: 500px; overflow-y: auto;">
                            <div style="font-weight: bold; margin-bottom: 5px;">{}</div>
                        '''.format(legend_header)

                        for color, name in legend_labels:
                            legend_html += f'''
                                <div style="display: inline-block; 
                                margin-right: 10px; background-color: {color}; 
                                width: 15px; height: 15px;"></div>{name}<br>
                            '''

                        # AOI + buffer
                        legend_html += '''
                            <div style="display: inline-block; 
                            margin-right: 10px;
                            background-color: transparent;
                            border: 2px solid red;
                            width: 15px; height: 15px;"></div>AOI<br>

                            <div style="display: inline-block; 
                            margin-right: 10px;background-color: transparent; 
                            border: 2px solid orange;
                            width: 15px; height: 15px;"></div>AOI buffers<br>
                        </div>
                        '''
            
                        #add the legend to the individual maps
                        map_one.get_root().html.add_child(folium.Element(legend_html))
                        end_timer()

                        end_timer = self.timed_step(f"Add layer control for individual maps") 
                        # Add layer controls to the individual map
                        lyr_cont_one = folium.LayerControl()
                        lyr_cont_one.add_to(map_one)
                        end_timer()

                        end_timer = self.timed_step(f"Add grouped layer control for individual maps; AOI, feature class") 
                        #Add AOI and buffer groups to the layer controls of the individual maps
                        GroupedLayerControl(
                        groups={
                        "AREA OF INTEREST": aoi_grps_o,
                            },
                        exclusive_groups=False,
                        collapsed=True
                            ).add_to(map_one)

                        #Add featured layer groups to the layer controls of the individual maps
                        GroupedLayerControl(
                            groups={
                                f"{map_title}": group_layers  # Correct placement of the key and colon
                            },
                            exclusive_groups=False,
                            collapsed=True
                        ).add_to(map_one)
                        end_timer()
                    
                        end_timer = self.timed_step(f"Save ind. map and cleanup") 
                        # Save the indivdiual map to html file
                        fc_map_path = Path(self.out_loc) / f"{fc}.html"
                        map_one.save(str(fc_map_path))
                        self.inject_branding(str(fc_map_path), title=map_title)

                        del legend_labels, legend_header
                        end_timer()
        
                    del gdf_fc
                counter += 1
        
                # Create a Legend for all-layers map
                legend_html_all = '''
                        <div id="legend" style="position: fixed; 
                        bottom: 200px; right: 30px; z-index: 1000; 
                        background-color: #fff; padding: 10px; 
                        border-radius: 5px; border: 1px solid grey;">
        
                        <div style="display: inline-block; 
                        margin-right: 10px;
                        background-color: transparent;
                        border: 2px solid red;
                        width: 15px; height: 15px;"></div>AOI<br>
                        
                        <div style="display: inline-block; 
                        margin-right: 10px;background-color: transparent; 
                        border: 2px solid orange;
                        width: 15px; height: 15px;"></div>AOI buffers<br>
                        
                        </div>
                        '''  
        
            if len(fc_grps) > 0:            
                ctg_grps.append(fc_grps) 
        
        #add the legend to the all-layers map
        end_timer = self.timed_step(f"Add legend to all-layers map")
        map_all.get_root().html.add_child(folium.Element(legend_html_all))
        end_timer()          
                
        # Add layer controls to the all-layers map
        end_timer = self.timed_step(f"Add layer controls to all-layers map")
        lyr_cont_all = folium.LayerControl()
        lyr_cont_all.add_to(map_all)
        end_timer()
        
        #Add status categories to the layer controls of  the all-layers map
        end_timer = self.timed_step(f"Add Grouped Layer Control to All-layers map")
        ctg_list = [x.upper() for x in ctg_list]
        GroupedLayerControl(
                    dict(zip(ctg_list, ctg_grps)),
                    exclusive_groups=False,
                    collapsed=False
                           ).add_to(map_all)
        end_timer()
        
        # Save the all-layers map to html file
        print('\nGenerating the all-layers map')
        end_timer = self.timed_step(f"Save All-layers map")
        all_path = Path(self.out_loc) / '00_all_layers.html'
        map_all.save(str(all_path))
        map_head = 'Overview Map - All Overlaps'
        self.inject_branding(str(all_path), title=map_head)
        end_timer()

        self.logger.info("[INFO] Complete.")
import arcpy

# Import your toolbox (no alias so tools register with their full name)
print("Importing toolbox...")
arcpy.ImportToolbox(r"\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\csostad\GitHub_Repositories\ast_batch_tool\autoast\auto_ast_v3_Breville_folium_maps\alpha_ast.atbx")

#Perform a check the toolbox was imported
if arcpy.Exists("alpha_ast"):
    print("Toolbox imported successfully.")



print("Searching for tools in the toolbox...")
# Filter tools to only show ones from your toolbox
search_text = "Statusing"  # or "Alpha", or whatever part you know
matching_tools = [t for t in arcpy.ListTools("*") if search_text.lower() in t.lower()]

print("Matching tools found:")
# Print matching tools
for tool in matching_tools:
    print(tool)

import bpy
import os
import json
import ifcopenshell
from bonsai.bim.ifc import IfcStore
import bonsai.tool as tool

# Make sure you put the right path where your files are stored
json_path = "OID_2.json"
out_path = "output2.ifc"
pset_name = "SBI_Custom"

# === Get current IFC file loaded via Bonsai ===
ifc_file = IfcStore.file
if not ifc_file:
    raise RuntimeError("No IFC file is loaded in Bonsai.")

# === Get the active Blender object and its IFC entity ===
obj = bpy.context.active_object
if not obj:
    raise RuntimeError("Please select a Blender object linked to an IFC entity.")

ifc_entity = tool.Ifc.get_entity(obj)
if not ifc_entity:
    raise RuntimeError("Selected object has no IFC entity linked to it.")


if not os.path.exists(json_path):
    raise FileNotFoundError(f"Could not find JSON file: {json_path}")

with open(json_path, "r") as f:
    data = json.load(f)

attributes = data.get("attributes", {})

# === Define which keys to include ===
allowed_keys = [
    "OBJECTID", "TO_PNT", "WIDTH", "HEIGHT",
    "SHAPE_DESC", "Shape_Leng", "GlobalID"
]
allowed_keys += [k for k in attributes.keys() if k.startswith("UUMS_")]

# === Filter relevant attributes ===
filtered_attrs = {k: v for k, v in attributes.items() if k in allowed_keys}

if not filtered_attrs:
    raise ValueError("No matching attributes found in the JSON file.")

# === Create IFC IfcPropertySingleValue entities ===
ifc_properties = []
for name, value in filtered_attrs.items():
    # Safely convert value type
    if isinstance(value, bool):
        nominal_value = ifc_file.create_entity("IfcBoolean", value)
    elif isinstance(value, (int, float)):
        nominal_value = ifc_file.create_entity("IfcReal", float(value))
    else:
        nominal_value = ifc_file.create_entity("IfcText", str(value))
    
    prop = ifc_file.create_entity("IfcPropertySingleValue", name, None, nominal_value, None)
    ifc_properties.append(prop)

# === Create Property Set ===
pset = ifc_file.create_entity(
    "IfcPropertySet",
    ifcopenshell.guid.new(),
    ifc_entity.OwnerHistory,
    pset_name,
    None,
    ifc_properties
)

# === Link Pset to the IFC element ===
ifc_file.create_entity(
    "IfcRelDefinesByProperties",
    ifcopenshell.guid.new(),
    ifc_entity.OwnerHistory,
    None,
    None,
    [ifc_entity],
    pset
)

print(f"Added custom Pset '{pset_name}' with {len(filtered_attrs)} properties to {ifc_entity.GlobalId}")

# === Optionally save a copy ===
ifc_file.write(out_path)
print(f"Saved modified IFC file: {out_path}")

# Blender / Bonsai / IfcOpenShell: Apply colours from Excel by "System Type"
# - Uses the currently open IFC (via BlenderBIM/Bonsai)
# - Immediately shows colours in Blender viewport by assigning Blender materials
# - ALSO assigns IfcSurfaceStyle in the IFC (so saving later preserves colours)
# - Does NOT rely on Blender object names (uses IFC definition id)

import os
import re
import traceback

# --------------------- CONFIG ---------------------

XLSX_PATH = r"SimpleBIM_Type_Filter.xlsx"     # <-- use the correct path where you store your excel file
SHEET_NAME     = "ModelView"
SAVE_NEW_IFC   = False                          # <-- keep False; set True only if user wants a new IFC
DEFAULT_OUT_BASENAME = "_COLORED.ifc"
# --------------------------------------------------

# ---- deps
try:
    import bpy
except Exception:
    bpy = None

import ifcopenshell
from ifcopenshell.api import run
from openpyxl import load_workbook

# BlenderBIM entry point
try:
    from bonsai.bim.ifc import IfcStore
except Exception:
    IfcStore = None

# ------------- Excel helpers (header-agnostic) -------------
REQ_HEADERS = ["Object or Group [+]", "Color", "Transparency %"]

def _find_header(ws):
    header_pos = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=False):
        texts = [str(c.value).strip() if c.value is not None else "" for c in row]
        lower = [t.lower() for t in texts]
        if all(h.lower() in lower for h in REQ_HEADERS):
            for idx, name in enumerate(texts):
                if name and name.strip().lower() in [h.lower() for h in REQ_HEADERS]:
                    header_pos[name.strip()] = idx
            return row[0].row, header_pos
    raise RuntimeError("Couldn't find a header row with: " + ", ".join(REQ_HEADERS))

def _cell_hex_or_fill(cell):
    # Prefer explicit #RRGGBB / RRGGBB (or 8-digit with alpha)
    if cell.value:
        s = str(cell.value).strip()
        m = re.fullmatch(r"#?[0-9A-Fa-f]{6,8}", s)
        if m:
            return "#" + s.lstrip("#")
    # Else try cell fill ARGB
    fill = cell.fill
    if getattr(fill, "fgColor", None):
        rgb = getattr(fill.fgColor, "rgb", None)
        if rgb and re.fullmatch(r"[0-9A-Fa-f]{8}", rgb):
            a = int(rgb[0:2], 16) / 255.0
            r = rgb[2:4]; g = rgb[4:6]; b = rgb[6:8]
            return "#" + r + g + b, a
    return None

def parse_mapping_from_excel(xlsx_path, sheet_name=SHEET_NAME):
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name]
    header_row, pos = _find_header(ws)

    name_col = pos[[h for h in REQ_HEADERS if h.startswith("Object")][0]]
    color_col = pos[[h for h in REQ_HEADERS if h == "Color"][0]]
    trans_col = pos[[h for h in REQ_HEADERS if h.startswith("Transparency")][0]]

    mapping = {}  # name -> (hex, transparency_0_1 or None)
    for r in range(header_row + 1, ws.max_row + 1):
        name_cell = ws.cell(r, name_col + 1)
        color_cell = ws.cell(r, color_col + 1)
        trans_cell = ws.cell(r, trans_col + 1)

        name = (str(name_cell.value).strip() if name_cell.value else "")
        if not name:
            continue

        hex_result = _cell_hex_or_fill(color_cell)
        hex_code, alpha_from_fill = (None, None)
        if isinstance(hex_result, tuple):
            hex_code, alpha_from_fill = hex_result
        elif isinstance(hex_result, str):
            hex_code = hex_result

        trans = None
        if trans_cell.value is not None:
            try:
                pct = float(trans_cell.value)
                trans = max(0.0, min(1.0, pct / 100.0))
            except Exception:
                pass

        if not hex_code:
            continue

        # If not specified, infer from ARGB alpha (openpyxl alpha=1 opaque)
        if trans is None and alpha_from_fill is not None:
            trans = 1.0 - alpha_from_fill  # IFC: 0 opaque, 1 fully transparent

        mapping[name.strip()] = (hex_code.upper(), trans)
    return mapping

# ------------- IFC helpers ----------------
def hex_to_rgb01(hex_code):
    s = hex_code.lstrip("#")
    r = int(s[0:2], 16) / 255.0
    g = int(s[2:4], 16) / 255.0
    b = int(s[4:6], 16) / 255.0
    return (r, g, b)

def _scalar(val):
    if val is None:
        return None
    return getattr(val, "wrappedValue", val)

def _collect_psets_on_instance(elem, f):
    psets = {}
    for rel in f.get_inverse(elem):
        if rel.is_a("IfcRelDefinesByProperties"):
            pdef = rel.RelatingPropertyDefinition
            if pdef and pdef.is_a("IfcPropertySet"):
                props = {}
                for p in (pdef.HasProperties or []):
                    if p.is_a("IfcPropertySingleValue"):
                        props[p.Name] = _scalar(p.NominalValue)
                psets[pdef.Name or ""] = props
    return psets

def _get_type(elem, f):
    for rel in f.get_inverse(elem):
        if rel.is_a("IfcRelDefinesByType"):
            return rel.RelatingType
    return None

def _collect_psets_on_type(typ):
    psets = {}
    for pdef in (getattr(typ, "HasPropertySets", []) or []):
        if pdef and pdef.is_a("IfcPropertySet"):
            props = {}
            for p in (pdef.HasProperties or []):
                if p.is_a("IfcPropertySingleValue"):
                    props[p.Name] = _scalar(p.NominalValue)
            psets[pdef.Name or ""] = props
    return psets

def _canon(s):
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())

def get_system_type(elem, f):
    # instance first
    psets = _collect_psets_on_instance(elem, f)
    for _pset, props in (psets or {}).items():
        for k, v in (props or {}).items():
            if _canon(k) == _canon("System Type") and v not in (None, ""):
                return str(v)
    # then type definition
    typ = _get_type(elem, f)
    if typ:
        tps = _collect_psets_on_type(typ)
        for _pset, props in (tps or {}).items():
            for k, v in (props or {}).items():
                if _canon(k) == _canon("System Type") and v not in (None, ""):
                    return str(v)
    return None

def get_body_reps(prod):
    reps = []
    rep = getattr(prod, "Representation", None)
    if not rep or not rep.Representations:
        return reps
    for r in rep.Representations:
        ident = getattr(r, "RepresentationIdentifier", None)
        ctx = getattr(r, "ContextOfItems", None)
        ctx_ident = getattr(ctx, "ContextIdentifier", None) if ctx else None
        if ident == "Body" or ctx_ident == "Body":
            reps.append(r)
    if not reps:
        reps = list(rep.Representations)
    return reps

# Cache (hex, trans) -> IfcSurfaceStyle
_style_cache = {}

def ensure_surface_style(f, hex_code, transparency=None):
    t = None if transparency is None else round(float(transparency), 3)
    key = (hex_code.upper(), t)
    if key in _style_cache:
        return _style_cache[key]

    r, g, b = hex_to_rgb01(hex_code)
    if t is None:
        t = 0.0

    style = run("style.add_style", f, name=f"SYS::{hex_code}")
    run("style.add_surface_style", f, style=style, ifc_class="IfcSurfaceStyleShading",
        attributes={
            "SurfaceColour": {"Name": None, "Red": r, "Green": g, "Blue": b},
            "Transparency": float(t),
        })
    _style_cache[key] = style
    return style

def assign_style_to_product(f, prod, style):
    for r in get_body_reps(prod):
        run("style.assign_representation_styles", f,
            shape_representation=r, styles=[style],
            replace_previous_same_type_style=True)

# ------------- Blender viewport material helpers ----------------
def get_or_create_material(hex_code, transparency):
    """Create/find a Blender material matching hex + transparency and set it up."""
    name = f"IFC_SYSCOLOR_{hex_code.upper()}_{0 if transparency is None else round(float(transparency),3)}"
    mat = bpy.data.materials.get(name)
    if mat is None:
        mat = bpy.data.materials.new(name)
        mat.use_nodes = True
        # Principled BSDF base color + alpha
        r, g, b = hex_to_rgb01(hex_code)
        alpha = 1.0 - (transparency or 0.0)
        bsdf = mat.node_tree.nodes.get("Principled BSDF")
        if bsdf:
            bsdf.inputs["Base Color"].default_value = (r, g, b, 1.0)
            bsdf.inputs["Alpha"].default_value = alpha
        # Viewport Display color (used by Solid shading when Color: Material)
        mat.diffuse_color = (r, g, b, 1.0)  # Solid ignores alpha anyway
        mat.blend_method = 'BLEND' if alpha < 1.0 else 'OPAQUE'
        mat.show_transparent_back = True
        mat.use_backface_culling = False
    return mat

def assign_material_to_ifc_product_objects(prod_id, mat):
    """Assign Blender material to all Blender objects that represent the IFC product."""
    assigned_obj_count = 0
    for obj in bpy.data.objects:
        props = getattr(obj, "BIMObjectProperties", None)
        if not props:
            continue
        if getattr(props, "ifc_definition_id", None) == prod_id:
            # ensure material on the mesh
            if obj.type == 'MESH' and obj.data:
                if len(obj.data.materials) == 0:
                    obj.data.materials.append(mat)
                else:
                    # Replace all slots for consistent look
                    for i in range(len(obj.data.materials)):
                        obj.data.materials[i] = mat
                assigned_obj_count += 1
    return assigned_obj_count

# Attempt to get the currently open IFC and its path
def get_open_ifc_and_path():
    f = None
    src_path = None
    if IfcStore:
        try:
            f = IfcStore.get_file()
        except Exception:
            f = None
        try:
            if hasattr(IfcStore, "get_file_path"):
                src_path = IfcStore.get_file_path()
            elif hasattr(IfcStore, "path"):
                src_path = IfcStore.path
        except Exception:
            src_path = None
    return f, src_path

def derive_out_path(src_path):
    if src_path and os.path.isfile(src_path):
        root, _ext = os.path.splitext(src_path)
        return root + DEFAULT_OUT_BASENAME
    # fallback to Desktop
    home = os.path.expanduser("~")
    desk = os.path.join(home, "Desktop")
    os.makedirs(desk, exist_ok=True)
    return os.path.join(desk, "colored" + DEFAULT_OUT_BASENAME)

def force_view_update():
    if bpy:
        try:
            for area in bpy.context.screen.areas:
                if area.type == 'VIEW_3D':
                    area.tag_redraw()
            bpy.context.view_layer.update()
        except Exception:
            pass

def main():
    try:
        # 1) Get currently open IFC
        f, src_path = get_open_ifc_and_path()
        if f is None:
            raise RuntimeError("No open IFC detected. Open your ORIGINAL IFC in Blender first, then run this script.")

        # 2) Parse Excel mapping
        mapping = parse_mapping_from_excel(XLSX_PATH, SHEET_NAME)
        print(f"[INFO] Loaded {len(mapping)} mapping rows from Excel.")
        norm = {k.strip().lower(): v for k, v in mapping.items()}

        # 3) Assign styles + viewport materials
        assigned_products = 0
        assigned_objects = 0
        missing = {}

        skip = {"IfcProject","IfcSite","IfcBuilding","IfcBuildingStorey","IfcSpace"}
        for prod in f.by_type("IfcProduct"):
            if prod.is_a() in skip:
                continue

            stype = get_system_type(prod, f)
            if not stype:
                continue

            pair = norm.get(stype.strip().lower())
            if not pair:
                missing[stype.strip().lower()] = missing.get(stype.strip().lower(), 0) + 1
                continue

            hex_code, transparency = pair
            # IFC data: ensure IfcSurfaceStyle + assign to Body reps
            style = ensure_surface_style(f, hex_code, transparency)
            assign_style_to_product(f, prod, style)
            assigned_products += 1

            # Blender viewport: apply a matching Blender material now
            if bpy is not None:
                mat = get_or_create_material(hex_code, transparency or 0.0)
                assigned_objects += assign_material_to_ifc_product_objects(prod.id(), mat)

        print(f"[INFO] Assigned IFC styles to {assigned_products} product(s).")
        print(f"[INFO] Assigned Blender materials to {assigned_objects} object(s).")
        if missing:
            print("[WARN] No colour mapping for these System Types (count → name):")
            for k, c in sorted(missing.items(), key=lambda x: (-x[1], x[0])):
                print(f"  {c:5d} → {k}")

        # 4) Immediate viewport refresh
        force_view_update()

        # 5) Optional: Save a NEW IFC (only if user wants)
        if SAVE_NEW_IFC:
            out_path = derive_out_path(src_path)
            try:
                f.write(out_path)
                print(f"[OK] Wrote new IFC with colours: {out_path}")
            except Exception as e:
                print("[ERROR] Couldn't write IFC file:", e)
                print("Styles are still applied in memory/viewport.")

    except Exception as e:
        print("[FATAL] Script aborted.")
        print(e)
        traceback.print_exc()

if __name__ == "__main__":
    main()
      

# -*- coding: utf-8 -*-



import os
import sys

try:
    from lib.core.branding import get_logo_path
except ImportError:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    while True:
        if os.path.basename(current_dir).lower() == "pymenvic.extension":
            lib_dir = os.path.join(current_dir, "lib")
            if lib_dir not in sys.path:
                sys.path.append(lib_dir)
            break
        parent_dir = os.path.dirname(current_dir)
        if parent_dir == current_dir:
            break
        current_dir = parent_dir
    from core.branding import get_logo_path

__title__ = "Workset Standardizer"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
WORKSET STANDARDIZER
_____________________________________________________

Description:

Analyze existing user worksets, suggest standard names by discipline,
allow manual editing, create missing target worksets, and apply changes
using safe rename or consolidate logic.

_____________________________________________________
What the tool does:

• detects likely standard targets by discipline
• allows manual review and editing before apply
• creates missing standard worksets when requested
• renames, consolidates and optionally deletes empty leftovers

_____________________________________________________
Output:

Preview report and apply summary in pyRevit output.

_____________________________________________________
Usage:

1. Open the tool
2. Review suggested targets and actions
3. Preview or apply changes

_____________________________________________________

Author: Ricardo J. Mendieta
"""

from pyrevit import revit, DB, forms, script
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from System import Uri, UriKind
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System.Windows.Data import CollectionViewSource
import os
import sys


doc = revit.doc




def is_workshared_document(current_doc):
    try:
        return current_doc is not None and current_doc.IsWorkshared
    except Exception:
        return False


def has_custom_user_worksets(current_doc):
    try:
        collector = DB.FilteredWorksetCollector(current_doc)
        worksets = collector.OfKind(DB.WorksetKind.UserWorkset).ToWorksets()
        custom_count = 0

        for ws in worksets:
            try:
                ws_name = ws.Name
                if ws_name and ws_name.strip().lower() != "workset1":
                    custom_count += 1
            except Exception:
                pass

        return custom_count > 0

    except Exception:
        return False
output = script.get_output()


# ==================================================
# CONFIG
# ==================================================

PLACEHOLDER_TARGET = "-- Select --"

DISCIPLINE_ORDER = [
    "ARCHITECTURE",
    "STRUCTURE",
    "MECHANICAL",
    "ELECTRICAL",
    "PLUMBING",
    "SITE",
    "COORDINATION"
]

BASE_STANDARD_WORKSETS = [
    "ARC_MODEL",
    "ARC_LEVELS_GRIDS",
    "STR_MODEL",
    "STR_LEVELS_GRIDS",
    "MECH_MODEL",
    "MECH_LEVELS_GRIDS",
    "ELE_MODEL",
    "ELE_LEVELS_GRIDS",
    "PLM_MODEL",
    "PLM_LEVELS_GRIDS",
    "SITE_MODEL",
    "SITE_LEVELS_GRIDS",
    "LINK_ARC",
    "LINK_STR",
    "LINK_MECH",
    "LINK_ELE",
    "LINK_PLM",
    "LINK_SITE",
    "LINK_CAD",
    "LINK_REF"
]

PROFILE_STANDARDS = {
    "ARCHITECTURE": [
        "ARC_LEVELS_GRIDS",
        "ARC_MODEL",
        "ARC_MODEL_CORE",
        "ARC_MODEL_FINISHES",
        "ARC_MODEL_INTERIORS"
    ],
    "STRUCTURE": [
        "STR_LEVELS_GRIDS",
        "STR_MODEL",
        "STR_MODEL_FOUNDATION",
        "STR_MODEL_SLABS",
        "STR_MODEL_COLUMNS",
        "STR_MODEL_BEAMS",
        "STR_MODEL_TRUSSES",
        "STR_MODEL_REBARS",
        "STR_MODEL_STEEL",
        "STR_MODEL_CONCRETE",
        "STR_MODEL_STAIRS"
    ],
    "MECHANICAL": [
        "MECH_LEVELS_GRIDS",
        "MECH_MODEL",
        "MECH_MODEL_EQUIPMENT",
        "MECH_MODEL_DUCTS",
        "MECH_MODEL_PIPING"
    ],
    "ELECTRICAL": [
        "ELE_LEVELS_GRIDS",
        "ELE_MODEL",
        "ELE_MODEL_LIGHTING",
        "ELE_MODEL_POWER",
        "ELE_MODEL_LOW_CURRENT"
    ],
    "PLUMBING": [
        "PLM_LEVELS_GRIDS",
        "PLM_MODEL",
        "PLM_MODEL_GENERAL",
        "PLM_MODEL_PLUVIAL",
        "PLM_MODEL_COLD_WATER",
        "PLM_MODEL_HOT_WATER",
        "PLM_MODEL_HOT_WATER_RETURN",
        "PLM_MODEL_SANITARY",
        "PLM_MODEL_FIRE_PROTECTION",
        "PLM_MODEL_TANKS",
        "PLM_MODEL_VENT"
    ],  
    "SITE": [
        "SITE_LEVELS_GRIDS",
        "SITE_MODEL",
        "SITE_MODEL_TOPO",
        "SITE_MODEL_EXTERNAL"
    ],
    "COORDINATION": [
        "SHARED_LEVELS_GRIDS",
        "COORD_MODEL",
        "COORD_SCOPE_BOXES",
        "COORD_REFERENCE"
    ]
}

LINK_STANDARDS = [
    "LINK_ARC",
    "LINK_STR",
    "LINK_MECH",
    "LINK_ELE",
    "LINK_PLM",
    "LINK_SITE",
    "LINK_CAD",
    "LINK_REF"
]

STATUS_STANDARD = "STANDARD"
STATUS_ASSIGNABLE = "ASSIGNABLE"
STATUS_REVIEW = "NEEDS REVIEW"
STATUS_KEEP_BY_USER = "KEEP BY USER"

ACTION_NO = "IGNORE"
ACTION_ASSIGN = "ASSIGN"
ACTION_REVIEW = "REVIEW"
ACTION_KEEP = "KEEP"

ACTION_VALUES = [ACTION_ASSIGN, ACTION_KEEP, ACTION_NO, ACTION_REVIEW]

DISCIPLINE_KEYWORDS = {
    "ARCHITECTURE": ["ARC", "ARCH", "ARCHITECT", "WALL", "DOOR", "ROOM", "CEILING", "FINISH", "MURO", "PISO"],
    "STRUCTURE": ["STR", "COLUMN", "COLUMNAS", "VIGA", "VIGAS", "BEAM", "SLAB", "LOSA", "LOSAS", "FOUND", "FUND", "FUNDACION", "STEEL", "HORMIGON", "FIERRO", "CERCHA", "ESCALERA", "DINTEL", "PERFIL", "PERFILES", "TRUSS", "REBAR"],
    "MECHANICAL": ["MECH", "HVAC", "DUCT", "AHU", "FAN", "DIFFUSER", "MECHANICAL"],
    "ELECTRICAL": ["ELE", "ELECT", "POWER", "LIGHT", "LIGHTING", "LOW CURRENT", "DATA", "CABLE", "TRAY"],
    "PLUMBING": ["PLM", "PLUMB", "SANIT", "DRAIN", "WATER", "PIPE", "PIPING"],
    "SITE": ["SITE", "TOPO", "ROAD", "EXTERNAL", "LANDSCAPE", "PLOT"],
    "COORDINATION": ["COORD", "SHARED", "GRID", "LEVEL", "REFERENCE"]
}


# ==================================================
# HELPERS
# ==================================================


def safe_str(value):
    try:
        return str(value)
    except Exception:
        return "Unknown error"


def safe_upper(value):
    if value is None:
        return ""
    try:
        return str(value).strip().upper()
    except Exception:
        try:
            return unicode(value).strip().upper()
        except Exception:
            return ""


def norm_name(value):
    text = safe_upper(value)
    text = text.replace("-", "_").replace(" ", "_")
    while "__" in text:
        text = text.replace("__", "_")
    return text


def contains_any(text, keywords):
    for item in keywords:
        if item in text:
            return True
    return False


def is_empty_target(name):
    return (not name) or safe_upper(name) == safe_upper(PLACEHOLDER_TARGET)


def make_structural_fallback_target(text):
    raw = norm_name(text)
    prefixes_to_remove = ["STR_MODEL_", "STR_", "MODEL_"]

    for prefix in prefixes_to_remove:
        if raw.startswith(prefix):
            raw = raw[len(prefix):]

    while raw.startswith("_"):
        raw = raw[1:]

    if not raw:
        return "STR_MODEL"

    return "STR_MODEL_{}".format(raw)


def get_existing_user_worksets():
    result = []
    collector = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
    for ws in collector:
        result.append(ws)
    result.sort(key=lambda x: x.Name)
    return result


def get_workset_map():
    result = {}
    for ws in get_existing_user_worksets():
        result[safe_upper(ws.Name)] = ws
    return result


def get_workset_by_name(name):
    return get_workset_map().get(safe_upper(name))


def get_existing_workset_names():
    return [safe_upper(ws.Name) for ws in get_existing_user_worksets()]


def build_workset_element_count_map():
    counts = {}
    try:
        elems = list(DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements())
    except Exception:
        elems = []

    for el in elems:
        try:
            wid = el.WorksetId.IntegerValue
            counts[wid] = counts.get(wid, 0) + 1
        except Exception:
            pass

    return counts


def get_true_workset_element_count(workset_id):
    try:
        return DB.FilteredElementCollector(doc) \
                 .WherePasses(DB.ElementWorksetFilter(workset_id, False)) \
                 .GetElementCount()
    except Exception:
        count = 0
        for el in DB.FilteredElementCollector(doc).WhereElementIsNotElementType():
            try:
                p = el.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                if p and p.HasValue and p.AsInteger() == workset_id.IntegerValue:
                    count += 1
            except Exception:
                pass
        return count


def get_reassignable_elements_by_workset_id(source_wsid):
    result = []
    fec = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()

    for el in fec:
        try:
            p = el.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
            if not p:
                continue
            if not p.HasValue:
                continue
            if p.AsInteger() != source_wsid:
                continue
            result.append(el)
        except Exception:
            pass

    return result


def create_workset_if_missing(name):
    name = safe_upper(name)
    existing = get_workset_by_name(name)
    if existing:
        return existing, False

    if hasattr(DB.WorksetTable, "IsWorksetNameUnique"):
        if not DB.WorksetTable.IsWorksetNameUnique(doc, name):
            return get_workset_by_name(name), False

    created = DB.Workset.Create(doc, name)
    return created, True


def rename_workset_safe(source_ws, target_name):
    if source_ws is None:
        return False, "Source workset not found."

    target_name = safe_upper(target_name)
    if not target_name:
        return False, "Empty target name."

    if safe_upper(source_ws.Name) == target_name:
        return True, "Already has target name."

    try:
        if hasattr(DB.WorksetTable, "IsWorksetNameUnique"):
            if not DB.WorksetTable.IsWorksetNameUnique(doc, target_name):
                return False, "Target workset already exists."
    except Exception:
        pass

    t = DB.Transaction(doc, "Rename Workset - {} -> {}".format(source_ws.Name, target_name))
    t.Start()
    try:
        if hasattr(DB.WorksetTable, "RenameWorkset"):
            DB.WorksetTable.RenameWorkset(doc, source_ws.Id, target_name)
        else:
            source_ws.Name = target_name
        t.Commit()
        return True, ""
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        return False, safe_str(ex)


def move_elements_to_workset(source_wsid, target_wsid):
    changed = 0
    skipped = 0
    failed = []

    for el in get_reassignable_elements_by_workset_id(source_wsid):
        try:
            p = el.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
            if not p:
                skipped += 1
                continue
            try:
                if p.IsReadOnly:
                    skipped += 1
                    continue
            except Exception:
                pass
            p.Set(target_wsid)
            changed += 1
        except Exception as ex:
            failed.append((el.Id.IntegerValue, safe_str(ex)))

    return changed, skipped, failed


def api_supports_delete_workset():
    return hasattr(DB, "WorksetTable") and hasattr(DB.WorksetTable, "DeleteWorkset")


def build_delete_move_settings(target_workset_id):
    if not hasattr(DB, "DeleteWorksetSettings"):
        return None
    if not hasattr(DB, "DeleteWorksetOption"):
        return None

    try:
        return DB.DeleteWorksetSettings(DB.DeleteWorksetOption.MoveElementsToWorkset, target_workset_id)
    except Exception:
        return None


def can_delete_workset_safe(workset_id, delete_settings=None):
    try:
        if not hasattr(DB.WorksetTable, "CanDeleteWorkset"):
            return True

        if delete_settings is not None:
            try:
                return DB.WorksetTable.CanDeleteWorkset(doc, workset_id, delete_settings)
            except Exception:
                pass

        try:
            return DB.WorksetTable.CanDeleteWorkset(doc, workset_id)
        except Exception:
            pass

        if hasattr(DB, "DeleteWorksetSettings"):
            try:
                settings = DB.DeleteWorksetSettings()
                return DB.WorksetTable.CanDeleteWorkset(doc, workset_id, settings)
            except Exception:
                pass

        return True
    except Exception:
        return False


def delete_workset_safe(workset_id, delete_settings=None):
    if not api_supports_delete_workset():
        raise Exception("DeleteWorkset API not available")

    if delete_settings is not None:
        try:
            return DB.WorksetTable.DeleteWorkset(doc, workset_id, delete_settings)
        except Exception:
            pass

    try:
        return DB.WorksetTable.DeleteWorkset(doc, workset_id)
    except Exception:
        pass

    if hasattr(DB, "DeleteWorksetSettings"):
        try:
            settings = DB.DeleteWorksetSettings()
            return DB.WorksetTable.DeleteWorkset(doc, workset_id, settings)
        except Exception:
            pass

    raise Exception("No compatible DeleteWorkset overload found")


def checkout_worksets_if_possible(workset_ids):
    checked_out = []
    failed = []

    if not workset_ids:
        return checked_out, failed

    ids = List[DB.WorksetId]()
    seen = set()

    for wsid in workset_ids:
        try:
            key = wsid.IntegerValue
        except Exception:
            key = str(wsid)

        if key in seen:
            continue
        seen.add(key)

        try:
            ids.Add(wsid)
        except Exception as ex:
            failed.append((wsid, "Could not add workset id: {}".format(safe_str(ex))))

    if ids.Count == 0:
        return checked_out, failed

    try:
        checked_collection = DB.WorksharingUtils.CheckoutWorksets(doc, ids)
        try:
            for wsid in checked_collection:
                checked_out.append(wsid)
        except Exception:
            pass
    except Exception as ex:
        reason = safe_str(ex)
        for wsid in ids:
            failed.append((wsid, reason))

    return checked_out, failed


def load_logo_if_available(window):
    try:
        logo_path = get_logo_path()

        if not os.path.exists(logo_path):
            return

        bmp = BitmapImage()
        bmp.BeginInit()
        bmp.CacheOption = BitmapCacheOption.OnLoad
        bmp.UriSource = Uri(logo_path, UriKind.Absolute)
        bmp.EndInit()
        window.logoImage.Source = bmp
    except Exception:
        pass


# ==================================================
# SCAN
# ==================================================


def is_base_standard_name(name):
    return safe_upper(name) in BASE_STANDARD_WORKSETS


def is_link_like_name(name):
    text = norm_name(name)
    if text.startswith("LINK_"):
        return True
    if contains_any(text, ["LINK", "RVT", "CAD", "XREF", "DWG"]):
        return True
    return False


def detect_discipline_from_names(names):
    scores = {}
    for disc in DISCIPLINE_ORDER:
        scores[disc] = 0

    for raw in names:
        text = norm_name(raw)

        if safe_upper(raw) in BASE_STANDARD_WORKSETS:
            if text.startswith("ARC_"):
                scores["ARCHITECTURE"] += 6
            elif text.startswith("STR_"):
                scores["STRUCTURE"] += 6
            elif text.startswith("MECH_"):
                scores["MECHANICAL"] += 6
            elif text.startswith("ELE_"):
                scores["ELECTRICAL"] += 6
            elif text.startswith("PLM_"):
                scores["PLUMBING"] += 6
            elif text.startswith("SITE_"):
                scores["SITE"] += 6

        for disc, keys in DISCIPLINE_KEYWORDS.items():
            for key in keys:
                if key in text:
                    scores[disc] = scores.get(disc, 0) + 1

        if text.startswith("ARC_"):
            scores["ARCHITECTURE"] += 4
        elif text.startswith("STR_"):
            scores["STRUCTURE"] += 4
        elif text.startswith("MECH_"):
            scores["MECHANICAL"] += 4
        elif text.startswith("ELE_"):
            scores["ELECTRICAL"] += 4
        elif text.startswith("PLM_"):
            scores["PLUMBING"] += 4
        elif text.startswith("SITE_"):
            scores["SITE"] += 4

    best_disc = "STRUCTURE"
    best_score = -1
    for disc in DISCIPLINE_ORDER:
        val = scores.get(disc, 0)
        if val > best_score:
            best_disc = disc
            best_score = val

    return best_disc


def get_profile_targets(discipline, include_links):
    targets = []
    for item in PROFILE_STANDARDS.get(discipline, []):
        targets.append(item)
    if include_links:
        for item in LINK_STANDARDS:
            if item not in targets:
                targets.append(item)
    return targets


def classify_and_suggest(name, discipline, include_links):
    text = norm_name(name)

    if is_base_standard_name(text):
        return STATUS_STANDARD, text, ACTION_NO, "Already matches base standard."

    if text in get_profile_targets(discipline, include_links):
        return STATUS_STANDARD, text, ACTION_NO, "Already matches active profile."

    if include_links and is_link_like_name(text):
        if contains_any(text, ["ARC", "ARCH"]):
            return STATUS_ASSIGNABLE, "LINK_ARC", ACTION_ASSIGN, "Detected architecture link."
        if contains_any(text, ["STR", "STRUCT", "STEEL"]):
            return STATUS_ASSIGNABLE, "LINK_STR", ACTION_ASSIGN, "Detected structure link."
        if contains_any(text, ["MECH", "HVAC", "DUCT"]):
            return STATUS_ASSIGNABLE, "LINK_MECH", ACTION_ASSIGN, "Detected mechanical link."
        if contains_any(text, ["ELE", "POWER", "LIGHT", "CABLE"]):
            return STATUS_ASSIGNABLE, "LINK_ELE", ACTION_ASSIGN, "Detected electrical link."
        if contains_any(text, ["PLM", "PLUMB", "SANIT", "WATER", "DRAIN", "PIPE"]):
            return STATUS_ASSIGNABLE, "LINK_PLM", ACTION_ASSIGN, "Detected plumbing link."
        if contains_any(text, ["SITE", "TOPO", "ROAD", "LAND"]):
            return STATUS_ASSIGNABLE, "LINK_SITE", ACTION_ASSIGN, "Detected site link."
        if contains_any(text, ["CAD", "DWG"]):
            return STATUS_ASSIGNABLE, "LINK_CAD", ACTION_ASSIGN, "Detected CAD link."
        return STATUS_ASSIGNABLE, "LINK_REF", ACTION_ASSIGN, "Detected generic link."

    if "LEVEL" in text or "GRID" in text:
        base = {
            "ARCHITECTURE": "ARC_LEVELS_GRIDS",
            "STRUCTURE": "STR_LEVELS_GRIDS",
            "MECHANICAL": "MECH_LEVELS_GRIDS",
            "ELECTRICAL": "ELE_LEVELS_GRIDS",
            "PLUMBING": "PLM_LEVELS_GRIDS",
            "SITE": "SITE_LEVELS_GRIDS",
            "COORDINATION": "SHARED_LEVELS_GRIDS"
        }.get(discipline, "")
        return STATUS_ASSIGNABLE, base, ACTION_ASSIGN, "Levels / grids pattern."

    if discipline == "STRUCTURE":
        if contains_any(text, ["FUND", "FOUND", "FOUNDATION", "FOUNDATIONS", "FUNDACION", "FUNDACIONES"]):
            return STATUS_ASSIGNABLE, "STR_MODEL_FOUNDATION", ACTION_ASSIGN, "Foundation related."
        if contains_any(text, ["LOSA", "LOSAS", "SLAB", "SLABS"]):
            return STATUS_ASSIGNABLE, "STR_MODEL_SLABS", ACTION_ASSIGN, "Slab related."
        if contains_any(text, ["COLUMN", "COLUMNS", "COLUMNA", "COLUMNAS"]):
            return STATUS_ASSIGNABLE, "STR_MODEL_COLUMNS", ACTION_ASSIGN, "Column related."
        if contains_any(text, ["VIGA", "VIGAS", "BEAM", "BEAMS", "DINTEL", "DINTELS"]):
            return STATUS_ASSIGNABLE, "STR_MODEL_BEAMS", ACTION_ASSIGN, "Beam related."
        if contains_any(text, ["CERCHA", "CERCHAS", "TRUSS", "TRUSSES"]):
            return STATUS_ASSIGNABLE, "STR_MODEL_TRUSSES", ACTION_ASSIGN, "Truss related."
        if contains_any(text, ["STAIR", "STAIRS", "ESCALERA", "ESCALERAS"]):
            return STATUS_ASSIGNABLE, "STR_MODEL_STAIRS", ACTION_ASSIGN, "Stair related."
        if contains_any(text, ["FIERRO", "FIERROS", "REBAR", "REBARS", "ARMADURA", "ARMADURAS"]):
            return STATUS_ASSIGNABLE, "STR_MODEL_REBARS", ACTION_ASSIGN, "Rebar related."
        if contains_any(text, ["STEEL", "METAL", "METALICO", "METALICOS", "METALICA", "METALICAS", "PERFIL", "PERFILES"]):
            return STATUS_ASSIGNABLE, "STR_MODEL_STEEL", ACTION_ASSIGN, "Steel related."
        if contains_any(text, ["HORMIGON", "CONCRETE"]):
            return STATUS_ASSIGNABLE, "STR_MODEL_CONCRETE", ACTION_ASSIGN, "Concrete related."
        fallback_target = make_structural_fallback_target(text)
        return STATUS_REVIEW, fallback_target, ACTION_REVIEW, "No exact rule matched. Generated fallback target from current workset name."

    if discipline == "ARCHITECTURE":
        if contains_any(text, ["FINISH", "FINISHES", "PISO", "MURO", "WALL", "DOOR", "WINDOW"]):
            return STATUS_ASSIGNABLE, "ARC_MODEL_FINISHES", ACTION_ASSIGN, "Finish / enclosure pattern."
        if contains_any(text, ["CORE", "SHAFT"]):
            return STATUS_ASSIGNABLE, "ARC_MODEL_CORE", ACTION_ASSIGN, "Core pattern."
        if contains_any(text, ["INTERIOR", "INTERIORS", "FURN", "CASEWORK"]):
            return STATUS_ASSIGNABLE, "ARC_MODEL_INTERIORS", ACTION_ASSIGN, "Interior pattern."
        return STATUS_REVIEW, "ARC_MODEL", ACTION_REVIEW, "General architecture guess."

    if discipline == "MECHANICAL":
        if contains_any(text, ["DUCT", "DIFFUSER", "FLEX"]):
            return STATUS_ASSIGNABLE, "MECH_MODEL_DUCTS", ACTION_ASSIGN, "Duct related."
        if contains_any(text, ["EQUIP", "AHU", "FCU", "FAN", "CHILLER"]):
            return STATUS_ASSIGNABLE, "MECH_MODEL_EQUIPMENT", ACTION_ASSIGN, "Equipment related."
        if contains_any(text, ["PIPE", "PIPING"]):
            return STATUS_ASSIGNABLE, "MECH_MODEL_PIPING", ACTION_ASSIGN, "Piping related."
        return STATUS_REVIEW, "MECH_MODEL", ACTION_REVIEW, "General mechanical guess."

    if discipline == "ELECTRICAL":
        if contains_any(text, ["LIGHT", "LIGHTING"]):
            return STATUS_ASSIGNABLE, "ELE_MODEL_LIGHTING", ACTION_ASSIGN, "Lighting related."
        if contains_any(text, ["POWER", "PANEL", "SWITCHBOARD"]):
            return STATUS_ASSIGNABLE, "ELE_MODEL_POWER", ACTION_ASSIGN, "Power related."
        if contains_any(text, ["DATA", "LOW", "COMM", "CABLE", "TRAY"]):
            return STATUS_ASSIGNABLE, "ELE_MODEL_LOW_CURRENT", ACTION_ASSIGN, "Low current related."
        return STATUS_REVIEW, "ELE_MODEL", ACTION_REVIEW, "General electrical guess."

    if discipline == "PLUMBING":
        if contains_any(text, ["SANIT", "SEWER"]):
            return STATUS_ASSIGNABLE, "PLM_MODEL_SANITARY", ACTION_ASSIGN, "Sanitary related."
        if contains_any(text, ["PLUVIAL", "PLUBIAL", "RAIN", "STORM"]):
            return STATUS_ASSIGNABLE, "PLM_MODEL_PLUVIAL", ACTION_ASSIGN, "Pluvial related."
        if contains_any(text, ["HOT_WATER_RETURN", "RETURN", "HWR"]):
            return STATUS_ASSIGNABLE, "PLM_MODEL_HOT_WATER_RETURN", ACTION_ASSIGN, "Hot water return related."
        if contains_any(text, ["HOT", "HW", "AGUA_CALIENTE"]):
            return STATUS_ASSIGNABLE, "PLM_MODEL_HOT_WATER", ACTION_ASSIGN, "Hot water related."
        if contains_any(text, ["COLD", "CW", "AGUA_FRIA"]):
            return STATUS_ASSIGNABLE, "PLM_MODEL_COLD_WATER", ACTION_ASSIGN, "Cold water related."
        if contains_any(text, ["FIRE", "CONTRAINCENDIOS", "PROTECCION_FUEGO"]):
            return STATUS_ASSIGNABLE, "PLM_MODEL_FIRE_PROTECTION", ACTION_ASSIGN, "Fire protection related."
        if contains_any(text, ["TANK", "TANQUE", "TANQUES"]):
            return STATUS_ASSIGNABLE, "PLM_MODEL_TANKS", ACTION_ASSIGN, "Tanks related."
        if contains_any(text, ["VENT", "VENTILACION"]):
            return STATUS_ASSIGNABLE, "PLM_MODEL_VENT", ACTION_ASSIGN, "Vent related."
        if contains_any(text, ["WATER", "PLUMB", "PLUMMING", "PIPE", "PIPING"]):
            return STATUS_ASSIGNABLE, "PLM_MODEL_GENERAL", ACTION_ASSIGN, "General plumbing related."
        return STATUS_REVIEW, "PLM_MODEL", ACTION_REVIEW, "General plumbing guess."

    if discipline == "SITE":
        if contains_any(text, ["TOPO", "SURVEY"]):
            return STATUS_ASSIGNABLE, "SITE_MODEL_TOPO", ACTION_ASSIGN, "Topo related."
        if contains_any(text, ["ROAD", "EXTERNAL", "LAND", "PAVING"]):
            return STATUS_ASSIGNABLE, "SITE_MODEL_EXTERNAL", ACTION_ASSIGN, "External works related."
        return STATUS_REVIEW, "SITE_MODEL", ACTION_REVIEW, "General site guess."

    if discipline == "COORDINATION":
        return STATUS_REVIEW, "COORD_MODEL", ACTION_REVIEW, "Coordination guess."

    return STATUS_REVIEW, "", ACTION_REVIEW, "Could not classify safely."


def build_rows(discipline, include_links):
    rows = ObservableCollection[object]()
    count_map = build_workset_element_count_map()
    existing = get_existing_user_worksets()
    all_targets = get_profile_targets(discipline, include_links)

    for ws in existing:
        src = safe_upper(ws.Name)
        elem_count = count_map.get(ws.Id.IntegerValue, 0)

        category, suggestion, action, reason = classify_and_suggest(src, discipline, include_links)
        final_target = suggestion if suggestion else PLACEHOLDER_TARGET

        available_targets = [PLACEHOLDER_TARGET]
        for item in all_targets:
            if item not in available_targets:
                available_targets.append(item)
        if suggestion and suggestion not in available_targets:
            available_targets.append(suggestion)
        if src not in available_targets:
            available_targets.append(src)

        row = WorksetRow(
            src,
            elem_count,
            category,
            suggestion,
            final_target,
            action,
            reason,
            available_targets
        )

        if safe_upper(src) == safe_upper(suggestion):
            row.Action = ACTION_NO
            row.FinalTarget = src
            row.Result = "Already standard"

        rows.Add(row)

    return rows


# ==================================================
# DATA MODEL
# ==================================================


class WorksetRow(object):
    def __init__(self, source_name, element_count, category, suggested_target, final_target, action, reason, available_targets):
        self.SourceName = source_name
        self.ElementCount = element_count
        self.Status = category
        self.DetectedType = category
        self.SuggestedTarget = suggested_target
        self.FinalTarget = final_target if final_target else PLACEHOLDER_TARGET
        self.Action = action
        self.Reason = reason
        self.Result = ""

        self.AvailableTargets = ObservableCollection[object]()
        for item in available_targets:
            self.AvailableTargets.Add(item)

        self.AvailableActions = ObservableCollection[object]()
        for item in ACTION_VALUES:
            self.AvailableActions.Add(item)


# ==================================================
# PROCESS
# ==================================================


class WorksetStandardizerWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)

        self._all_rows = ObservableCollection[object]()
        self._view = None
        self._workset_map = get_workset_map()
        self._count_map = build_workset_element_count_map()

        load_logo_if_available(self)
        self.setup_grid()
        self.bind_events()
        self.initialize_defaults()
        self.ShowDialog()

    def setup_grid(self):
        self.dgMappings.ItemsSource = self._all_rows
        self._view = CollectionViewSource.GetDefaultView(self.dgMappings.ItemsSource)
        self._view.Filter = self.row_filter
        self.lstActiveTargets.ItemsSource = ObservableCollection[object]()

    def bind_events(self):
        self.btnAnalyze.Click += self.on_analyze
        self.btnRefresh.Click += self.on_refresh
        self.btnPreview.Click += self.on_preview
        self.btnApply.Click += self.on_apply
        self.btnClose.Click += self.on_close
        self.btnSuggestStandards.Click += self.on_suggest_profile
        self.btnCreateSelectedStandards.Click += self.on_create_profile_targets
        self.txtSearch.TextChanged += self.on_filter_changed
        self.cmbActionFilter.SelectionChanged += self.on_filter_changed
        self.cmbDisciplineMode.SelectionChanged += self.on_discipline_changed
        self.chkIncludeLinks.Checked += self.on_include_links_changed
        self.chkIncludeLinks.Unchecked += self.on_include_links_changed
        self.chkShowStandard.Checked += self.on_filter_changed
        self.chkShowStandard.Unchecked += self.on_filter_changed

    def initialize_defaults(self):
        names = get_existing_workset_names()
        detected = detect_discipline_from_names(names)

        self.chkIncludeLinks.IsChecked = False
        self.chkShowStandard.IsChecked = False
        self.chkDeleteEmptyAfterMap.IsChecked = False
        self.chkRenameDuplicates.IsChecked = True

        if hasattr(self, "chkDeleteEmptyAfterMap"):
            self.chkDeleteEmptyAfterMap.IsEnabled = api_supports_delete_workset()

        self.set_discipline(detected)
        self.update_active_targets_panel()
        self.reload_rows()
        self.update_status("Ready. Detected discipline: {}".format(detected))

    def update_status(self, text):
        self.txtStatus.Text = text

    def get_active_discipline(self):
        try:
            item = self.cmbDisciplineMode.SelectedItem
            if item is None:
                return "STRUCTURE"
            if hasattr(item, "Content"):
                return safe_upper(item.Content)
            return safe_upper(item.ToString())
        except Exception:
            return "STRUCTURE"

    def include_links(self):
        try:
            return bool(self.chkIncludeLinks.IsChecked)
        except Exception:
            return False

    def show_standard(self):
        try:
            return bool(self.chkShowStandard.IsChecked)
        except Exception:
            return False

    def get_action_filter(self):
        try:
            item = self.cmbActionFilter.SelectedItem
            if item is None:
                return "ALL"
            if hasattr(item, "Content"):
                return safe_upper(item.Content)
            return safe_upper(item.ToString())
        except Exception:
            return "ALL"

    def set_discipline(self, discipline):
        target = safe_upper(discipline)
        for i in range(self.cmbDisciplineMode.Items.Count):
            item = self.cmbDisciplineMode.Items[i]
            if hasattr(item, "Content"):
                text = safe_upper(item.Content)
            else:
                text = safe_upper(item.ToString())
            if text == target:
                self.cmbDisciplineMode.SelectedIndex = i
                return

    def update_active_targets_panel(self):
        targets = get_profile_targets(self.get_active_discipline(), self.include_links())
        items = ObservableCollection[object]()
        for name in targets:
            items.Add(name)
        self.lstActiveTargets.ItemsSource = items
        self.txtSelectedStandards.Text = str(len(targets))

    def row_filter(self, obj):
        row = obj
        search = safe_upper(self.txtSearch.Text)
        if search and search not in safe_upper(row.SourceName):
            return False

        if (not self.show_standard()) and safe_upper(row.Status) == STATUS_STANDARD:
            return False

        action_filter = self.get_action_filter()
        if action_filter != "ALL" and safe_upper(row.Action) != action_filter:
            return False

        return True

    def refresh_grid_ui(self):
        try:
            self.dgMappings.Items.Refresh()
        except Exception:
            pass
        try:
            self._view.Refresh()
        except Exception:
            pass
        self.update_summary()

    def sync_row_action_targets(self):
        for row in self._all_rows:
            action = safe_upper(row.Action)

            if action in [ACTION_NO, ACTION_KEEP]:
                row.FinalTarget = row.SourceName
            elif action == ACTION_ASSIGN:
                if is_empty_target(row.FinalTarget) and (not is_empty_target(row.SuggestedTarget)):
                    row.FinalTarget = row.SuggestedTarget
            elif action == ACTION_REVIEW:
                if is_empty_target(row.FinalTarget):
                    row.FinalTarget = row.SuggestedTarget if row.SuggestedTarget else PLACEHOLDER_TARGET

    def reload_rows(self):
        discipline = self.get_active_discipline()
        include_links = self.include_links()

        self._workset_map = get_workset_map()
        self._count_map = build_workset_element_count_map()
        self._all_rows = build_rows(discipline, include_links)
        self.dgMappings.ItemsSource = self._all_rows
        self._view = CollectionViewSource.GetDefaultView(self.dgMappings.ItemsSource)
        self._view.Filter = self.row_filter
        self.sync_row_action_targets()
        self.refresh_grid_ui()

    def update_summary(self):
        total_rows = len(list(self._all_rows))
        standard_count = 0
        assignable_count = 0
        review_count = 0
        keep_by_user_count = 0

        for row in self._all_rows:
            status = safe_upper(row.Status)
            if status == STATUS_STANDARD:
                standard_count += 1
            elif status == STATUS_ASSIGNABLE:
                assignable_count += 1
            elif status == STATUS_REVIEW:
                review_count += 1

            if safe_upper(row.Action) == ACTION_KEEP:
                keep_by_user_count += 1

        self.txtTotalRows.Text = str(total_rows)
        self.txtStandardRows.Text = str(standard_count)
        self.txtAssignableRows.Text = str(assignable_count)
        self.txtReviewRows.Text = str(review_count)
        self.txtKeepByUserRows.Text = str(keep_by_user_count)
        self.txtSelectedStandards.Text = str(len(get_profile_targets(self.get_active_discipline(), self.include_links())))

    def get_selected_profile_worksets(self):
        return get_profile_targets(self.get_active_discipline(), self.include_links())

    def build_preview_plan(self):
        self.sync_row_action_targets()

        rename_duplicates = False
        delete_empty = False
        try:
            rename_duplicates = bool(self.chkRenameDuplicates.IsChecked)
        except Exception:
            pass
        try:
            delete_empty = bool(self.chkDeleteEmptyAfterMap.IsChecked)
        except Exception:
            pass

        plan = []
        invalid = []

        for row in self._all_rows:
            row.Result = ""
            action = safe_upper(row.Action)

            if action != ACTION_ASSIGN:
                if action == ACTION_KEEP:
                    row.Result = "Keep"
                elif action == ACTION_NO:
                    row.Result = "Ignore"
                else:
                    row.Result = "Review"
                continue

            target = safe_upper(row.FinalTarget)
            if is_empty_target(target):
                row.Result = "Missing target"
                invalid.append(row)
                continue

            if safe_upper(row.SourceName) == target:
                row.Result = "Same name"
                continue

            target_ws = get_workset_by_name(target)
            if target_ws is None:
                row.Result = "Rename"
                plan.append((row, "RENAME"))
            else:
                if rename_duplicates:
                    row.Result = "Consolidate"
                    plan.append((row, "CONSOLIDATE"))
                else:
                    row.Result = "Conflict"
                    plan.append((row, "CONFLICT"))

        return plan, invalid, rename_duplicates, delete_empty

    def print_preview_output(self, plan, rename_duplicates, delete_empty):
        output.print_md("# pyMENVIC | WORKSET STANDARDIZER — PREVIEW")
        output.print_md("")
        output.print_md("## Summary")
        output.print_md("")
        output.print_md("- **Discipline mode:** {}".format(self.get_active_discipline()))
        output.print_md("- **Include link worksets:** {}".format("Yes" if self.include_links() else "No"))
        output.print_md("- **Rows in plan:** {}".format(len([1 for _, mode in plan if mode in ["RENAME", "CONSOLIDATE", "CONFLICT"]])))
        output.print_md("- **Consolidate duplicates:** {}".format("Yes" if rename_duplicates else "No"))
        output.print_md("- **Delete empty leftovers after apply:** {}".format("Yes" if delete_empty else "No"))
        output.print_md("")

        if plan:
            output.print_md("## Planned actions")
            output.print_md("")
            for row, mode in plan:
                output.print_md("- `{}` → `{}` | **{}** | elements: **{}** | {}".format(
                    row.SourceName,
                    row.FinalTarget,
                    mode,
                    row.ElementCount,
                    row.Reason
                ))
            output.print_md("")

    def on_analyze(self, sender, args):
        self.reload_rows()
        self.update_status("Analysis completed.")

    def on_refresh(self, sender, args):
        self.update_active_targets_panel()
        self.reload_rows()
        self.update_status("Refreshed.")

    def on_preview(self, sender, args):
        # TEMP CHANGE - PREVIEW REPORT DISABLED
        # Keep preview recalculation, skip preview report output.
        plan, invalid, rename_duplicates, delete_empty = self.build_preview_plan()
        self.refresh_grid_ui()

        if invalid:
            forms.alert("There are ASSIGN rows without Final Target. Please review them first.", title="Preview")
            self.update_status("Preview found rows without target.")
            return

        self.update_status("Preview report temporarily disabled. Table refreshed from current actions.")
        return

        self.print_preview_output(plan, rename_duplicates, delete_empty)
        self.update_status("Preview printed to pyRevit output.")

    def on_close(self, sender, args):
        self.Close()

    def on_filter_changed(self, sender, args):
        self.sync_row_action_targets()
        self.refresh_grid_ui()

    def on_include_links_changed(self, sender, args):
        self.update_active_targets_panel()
        self.reload_rows()
        self.update_status("Link visibility updated.")

    def on_discipline_changed(self, sender, args):
        self.update_active_targets_panel()
        self.reload_rows()
        self.update_status("Discipline changed to {}.".format(self.get_active_discipline()))

    def on_suggest_profile(self, sender, args):
        self.update_active_targets_panel()
        self.reload_rows()
        self.update_status("Suggested standards loaded for {}.".format(self.get_active_discipline()))

    def on_create_profile_targets(self, sender, args):
        selected = self.get_selected_profile_worksets()
        if not selected:
            forms.alert("No standard worksets available.", title="Create Selected Standards")
            return

        created = []
        skipped = []

        t = DB.Transaction(doc, "pyMENVIC | Create Selected Standardizer Worksets")
        t.Start()
        try:
            for name in selected:
                try:
                    ws, was_created = create_workset_if_missing(name)
                    if was_created:
                        created.append(name)
                    else:
                        skipped.append(name)
                except Exception as ex:
                    skipped.append("{} ({})".format(name, safe_str(ex)))
            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            forms.alert("Could not create standards:\n{}".format(safe_str(ex)), title="Create Selected Standards")
            return

        self.update_active_targets_panel()
        self.reload_rows()

        output.print_md("# pyMENVIC | WORKSET STANDARDIZER — CREATE SELECTED STANDARDS")
        output.print_md("")
        output.print_md("- **Created:** {}".format(len(created)))
        output.print_md("- **Skipped / already existed:** {}".format(len(skipped)))
        output.print_md("")

        self.update_status("Create selected standards completed.")

    def delete_empty_leftovers(self, processed_rows):
        deleted = []
        skipped = []
        failed = []

        if not api_supports_delete_workset():
            skipped.append(("-", "Delete workset API not available"))
            return deleted, skipped, failed

        t = DB.Transaction(doc, "pyMENVIC | Delete Empty Standardizer Leftovers")
        t.Start()
        try:
            for row in processed_rows:
                try:
                    source_ws = get_workset_by_name(row.SourceName)
                    target_ws = get_workset_by_name(row.FinalTarget)

                    if not source_ws:
                        skipped.append((row.SourceName, "Source workset not found"))
                        continue

                    if safe_upper(source_ws.Name) == safe_upper(row.FinalTarget):
                        skipped.append((row.SourceName, "Source now equals target"))
                        continue

                    if target_ws is None:
                        skipped.append((row.SourceName, "Target workset not found"))
                        continue

                    remaining = get_true_workset_element_count(source_ws.Id)
                    if remaining > 0:
                        move_settings = build_delete_move_settings(target_ws.Id)
                        if move_settings is None:
                            skipped.append((row.SourceName, "Delete move settings not available"))
                            continue
                        if not can_delete_workset_safe(source_ws.Id, move_settings):
                            skipped.append((row.SourceName, "Revit did not allow move-and-delete"))
                            continue
                        delete_workset_safe(source_ws.Id, move_settings)
                        deleted.append(row.SourceName)
                    else:
                        if not can_delete_workset_safe(source_ws.Id):
                            skipped.append((row.SourceName, "Revit did not allow deleting this workset"))
                            continue
                        delete_workset_safe(source_ws.Id)
                        deleted.append(row.SourceName)
                except Exception as ex:
                    failed.append((row.SourceName, safe_str(ex)))
            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            failed.append(("-", "Delete transaction failed: {}".format(safe_str(ex))))

        return deleted, skipped, failed

    def on_apply(self, sender, args):
        self.sync_row_action_targets()
        plan, invalid, rename_duplicates, delete_empty = self.build_preview_plan()
        self.refresh_grid_ui()

        if invalid:
            forms.alert("Some rows are marked as ASSIGN but have no Final Target.", title="Apply")
            self.update_status("Apply canceled. Missing target detected.")
            return

        actionable = [(row, mode) for row, mode in plan if mode in ["RENAME", "CONSOLIDATE", "CONFLICT"]]
        if not actionable:
            forms.alert("No ASSIGN rows to process.", title="Apply")
            self.update_status("Nothing to apply.")
            return

        conflict_count = len([1 for _, mode in actionable if mode == "CONFLICT"])
        if conflict_count and not rename_duplicates:
            forms.alert("There are target conflicts. Enable duplicate consolidation or review those rows.", title="Apply")
            self.update_status("Apply canceled. Target conflicts detected.")
            return

        confirm = forms.alert(
            "Apply standardization to {} row(s)?\n\nConsolidate duplicates: {}\nDelete empty leftovers after apply: {}".format(
                len([1 for _, mode in actionable if mode in ["RENAME", "CONSOLIDATE"]]),
                "Yes" if rename_duplicates else "No",
                "Yes" if delete_empty else "No"
            ),
            yes=True,
            no=True
        )
        if not confirm:
            self.update_status("Apply canceled by user.")
            return

        rows_renamed = 0
        rows_consolidated = 0
        rows_conflict = 0
        rows_failed = 0
        elements_moved = 0
        issues = []
        processed_rows = []

        ids_to_checkout = []
        ws_map = get_workset_map()
        for row, mode in actionable:
            source_ws = ws_map.get(safe_upper(row.SourceName))
            target_ws = ws_map.get(safe_upper(row.FinalTarget))
            if source_ws:
                ids_to_checkout.append(source_ws.Id)
            if target_ws:
                ids_to_checkout.append(target_ws.Id)

        checked_out, failed_checkout = checkout_worksets_if_possible(ids_to_checkout)
        for _, reason in failed_checkout:
            issues.append("WORKSET CHECKOUT | {}".format(reason))

        for row, mode in actionable:
            src_name = safe_upper(row.SourceName)
            dst_name = safe_upper(row.FinalTarget)

            if mode == "CONFLICT":
                row.Result = "Conflict"
                rows_conflict += 1
                issues.append("{} -> {} | Target already exists. Rename skipped.".format(src_name, dst_name))
                continue

            source_ws = get_workset_by_name(src_name)
            if source_ws is None:
                row.Result = "Failed"
                rows_failed += 1
                issues.append("{} -> {} | Source workset not found.".format(src_name, dst_name))
                continue

            if mode == "RENAME":
                ok, msg = rename_workset_safe(source_ws, dst_name)
                if ok:
                    row.Result = "Renamed"
                    rows_renamed += 1
                    processed_rows.append(row)
                else:
                    row.Result = "Failed"
                    rows_failed += 1
                    issues.append("{} -> {} | Rename failed: {}".format(src_name, dst_name, msg))
                continue

            if mode == "CONSOLIDATE":
                target_ws = get_workset_by_name(dst_name)
                if target_ws is None:
                    row.Result = "Failed"
                    rows_failed += 1
                    issues.append("{} -> {} | Target workset not found for consolidation.".format(src_name, dst_name))
                    continue

                t = DB.Transaction(doc, "pyMENVIC | Consolidate Workset - {}".format(src_name))
                t.Start()
                try:
                    changed, skipped, failed = move_elements_to_workset(source_ws.Id.IntegerValue, target_ws.Id.IntegerValue)
                    elements_moved += changed

                    remaining = get_true_workset_element_count(source_ws.Id)
                    delete_reason = ""
                    deleted = False

                    if remaining <= 0:
                        if can_delete_workset_safe(source_ws.Id):
                            delete_workset_safe(source_ws.Id)
                            deleted = True
                        else:
                            delete_reason = "Revit did not allow deleting consolidated source workset"
                    else:
                        move_settings = build_delete_move_settings(target_ws.Id)
                        if move_settings and can_delete_workset_safe(source_ws.Id, move_settings):
                            delete_workset_safe(source_ws.Id, move_settings)
                            deleted = True
                        else:
                            delete_reason = "Source workset still contains {} element(s)".format(remaining)

                    t.Commit()

                    row.Result = "Consolidated" if deleted else "Moved"
                    rows_consolidated += 1
                    processed_rows.append(row)

                    if failed:
                        issues.append("{} -> {} | {} moved / {} failed".format(src_name, dst_name, changed, len(failed)))
                    if delete_reason:
                        issues.append("{} -> {} | {}".format(src_name, dst_name, delete_reason))
                except Exception as ex:
                    try:
                        t.RollBack()
                    except Exception:
                        pass
                    row.Result = "Failed"
                    rows_failed += 1
                    issues.append("{} -> {} | Consolidate failed: {}".format(src_name, dst_name, safe_str(ex)))

        deleted_worksets = []
        skipped_delete = []
        failed_delete = []

        if delete_empty:
            deleted_worksets, skipped_delete, failed_delete = self.delete_empty_leftovers(processed_rows)

        self.update_active_targets_panel()
        self.reload_rows()

        output.print_md("# pyMENVIC | WORKSET STANDARDIZER — APPLY SUMMARY")
        output.print_md("")
        output.print_md("## Summary")
        output.print_md("")
        output.print_md("- **Mode:** Rename / Consolidate")
        output.print_md("- **Discipline mode:** {}".format(self.get_active_discipline()))
        output.print_md("- **Rows renamed:** {}".format(rows_renamed))
        output.print_md("- **Rows consolidated:** {}".format(rows_consolidated))
        output.print_md("- **Rows skipped by conflict:** {}".format(rows_conflict))
        output.print_md("- **Rows failed:** {}".format(rows_failed))
        output.print_md("- **Elements moved:** {}".format(elements_moved))
        output.print_md("- **Delete empty leftovers:** {}".format("Yes" if delete_empty else "No"))
        output.print_md("")

        if issues:
            output.print_md("## Issues")
            output.print_md("")
            for item in issues:
                output.print_md("- {}".format(item))
            output.print_md("")

        if delete_empty:
            output.print_md("## Cleanup")
            output.print_md("")
            output.print_md("- **Worksets deleted:** {}".format(len(deleted_worksets)))
            output.print_md("- **Delete skipped:** {}".format(len(skipped_delete)))
            output.print_md("- **Delete errors:** {}".format(len(failed_delete)))
            output.print_md("")

        self.update_status("Apply completed.")


# ==================================================
# CLEANUP
# ==================================================


if not is_workshared_document(doc):
    forms.alert(
        "This tool requires a workshared model with worksets enabled.\n\nEnable Worksharing first and run the tool again.",
        title="pyMENVIC | Worksets Required",
        warn_icon=True
    )
    raise SystemExit


if not has_custom_user_worksets(doc):
    forms.alert(
        "This model has Worksharing enabled, but no custom worksets were found.\n\nOnly the default 'Workset1' exists.\n\nCreate or seed the required pyMENVIC worksets first, then run this tool again.",
        title="pyMENVIC | No Custom Worksets Found",
        warn_icon=True
    )
    raise SystemExit


# ==================================================
# REPORT
# ==================================================


xaml_path = os.path.join(os.path.dirname(__file__), "WorksetStandardizer.xaml")
WorksetStandardizerWindow(xaml_path)

# -*- coding: utf-8 -*-

__title__ = "Analyze Worksets"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
ANALYZE WORKSETS
_____________________________________________________

Description:

Analyzes user worksets in the current model against the
pyMENVIC standard and produces a consolidated report.

This tool DOES NOT modify the model.

_____________________________________________________
What the tool does:

• checks if model is workshared
• reads all user worksets
• compares against pyMENVIC standard worksets
• reports standard worksets present
• reports missing standard worksets
• reports extra non-standard worksets
• classifies non-standard worksets
• suggests possible mappings to standard worksets
• highlights manual review items
• prints a formatted final report

_____________________________________________________
Usage:

1. Click the pyRevit button
2. Tool runs automatically
3. Review the report

_____________________________________________________

Author: Ricardo J. Mendieta
"""

from Autodesk.Revit.DB import WorksetKind, FilteredWorksetCollector
from pyrevit import revit, script

doc = revit.doc
output = script.get_output()

# ==================================================
# CONFIG
# ==================================================

STANDARD_WORKSETS = [
    "ARC_MODEL",
    "ARC_LEVELS_GRIDS",
    "STR_MODEL",
    "STR_LEVELS_GRIDS",
    "MEP_MODEL",
    "MEP_LEVELS_GRIDS",
    "ELE_MODEL",
    "ELE_LEVELS_GRIDS",
    "PLM_MODEL",
    "PLM_LEVELS_GRIDS",
    "LINK_ARC",
    "LINK_STR",
    "LINK_MEP",
    "LINK_ELE",
    "LINK_PLM",
    "LINK_SITE",
    "LINK_CAD",
    "LINK_REF"
]

DISCIPLINE_ORDER = [
    "ARC",
    "STR",
    "MEP",
    "ELE",
    "PLM",
    "LINK",
    "OTHER"
]

CATEGORY_ORDER = [
    "GENERIC",
    "LINK_RELATED",
    "SHARED_LEVELS_GRIDS",
    "DISCIPLINE_SPECIFIC",
    "LEGACY_OR_CUSTOM",
    "UNKNOWN"
]

GENERIC_NAMES = [
    "WORKSET1",
    "WORKSET 1",
    "SHARED LEVELS AND GRIDS",
    "SHARED VIEWS, LEVELS, GRIDS",
    "LINKED FILES",
    "LINKED REVIT FILES",
    "LINKED FILES FOR SHEETS",
    "INTERIOR",
    "EXTERIOR",
    "COLUMNS",
    "BEAMS",
    "WALLS",
    "FOUNDATION",
    "STAIRS",
    "ENLARGED",
    "MATERIAL SCHEDULE",
    "CORE AND SHELL",
    "MONUMENT SIGN"
]

LINK_KEYWORDS = [
    "LINK",
    "LINKED",
    "REVIT FILES",
    "SITE_LINKED",
    "_LINKED FILES",
    "LINKED FILES",
    "MECHANICAL LINK",
    "STRUCTURAL LINK"
]

SHARED_KEYWORDS = [
    "SHARED LEVELS AND GRIDS",
    "SHARED VIEWS, LEVELS, GRIDS",
    "LEVELS, GRIDS, VIEWS",
    "VIEWS, LEVELS, GRIDS"
]

DISCIPLINE_HINTS = [
    "ARQ",
    "AR_",
    "ARC",
    "A_",
    "STR",
    "ST_",
    "S_",
    "MEP",
    "ME_",
    "M_",
    "ELE",
    "EL_",
    "PLM",
    "PL_",
    "P_",
    "SITE",
    "ESTRUCT",
    "ARQUITECT",
    "MUROS",
    "VIGAS",
    "COLUMN",
    "LOSAS",
    "FUNDACION",
    "HORMIGON",
    "REBAR",
    "FIERRO",
    "PERFILES",
    "PLACA",
    "ESCALERA",
    "TOPOSOLID",
    "FURNITURE",
    "MOBILIARIO"
]

LEGACY_HINTS = [
    "OLD",
    "NEW",
    "TOWER",
    "ANCILLARY",
    "GUARD",
    "TEMPLO",
    "TRANSFOMADOR",
    "UTILITES"
]

# ==================================================
# HELPERS
# ==================================================

def normalize_name(name):
    try:
        return name.strip().upper()
    except:
        return ""

def tokenize_name(name):
    n = normalize_name(name)
    n = n.replace("-", "_").replace(" ", "_")
    parts = [p for p in n.split("_") if p]
    return parts

def starts_with_any(text, prefixes):
    for prefix in prefixes:
        if text.startswith(prefix):
            return True
    return False

def has_token(text, token):
    parts = tokenize_name(text)
    return token in parts

def contains_any(text, keywords):
    upper_text = normalize_name(text)
    for kw in keywords:
        if kw in upper_text:
            return True
    return False

def get_existing_user_workset_names():
    names = []
    collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    for ws in collector:
        names.append(ws.Name)
    return sorted(names)

def get_group_name(ws_name):
    n = normalize_name(ws_name)

    if n.startswith("ARC_"):
        return "ARC"
    elif n.startswith("STR_"):
        return "STR"
    elif n.startswith("MEP_"):
        return "MEP"
    elif n.startswith("ELE_"):
        return "ELE"
    elif n.startswith("PLM_"):
        return "PLM"
    elif n.startswith("LINK_"):
        return "LINK"
    else:
        return "OTHER"

def classify_nonstandard_workset(name):
    normalized = normalize_name(name)

    if normalized in GENERIC_NAMES:
        return "GENERIC"

    if contains_any(normalized, LINK_KEYWORDS):
        return "LINK_RELATED"

    if contains_any(normalized, SHARED_KEYWORDS):
        return "SHARED_LEVELS_GRIDS"

    if contains_any(normalized, LEGACY_HINTS):
        return "LEGACY_OR_CUSTOM"

    if contains_any(normalized, DISCIPLINE_HINTS):
        return "DISCIPLINE_SPECIFIC"

    return "UNKNOWN"

def group_by_category(names):
    grouped = {}
    for category in CATEGORY_ORDER:
        grouped[category] = []

    for name in names:
        category = classify_nonstandard_workset(name)
        grouped[category].append(name)

    return grouped

def group_by_discipline(names):
    grouped = {}
    for group_name in DISCIPLINE_ORDER:
        grouped[group_name] = []

    for name in names:
        group_name = get_group_name(name)
        grouped[group_name].append(name)

    return grouped

def print_grouped_discipline_section(title, items):
    if not items:
        return

    grouped = group_by_discipline(items)

    output.print_md("## {}".format(title))
    output.print_md("")

    for group_name in DISCIPLINE_ORDER:
        group_items = grouped.get(group_name, [])
        if not group_items:
            continue

        output.print_md("### {} ({})".format(group_name, len(group_items)))
        output.print_md("")
        for name in sorted(group_items):
            output.print_md("- `{}`".format(name))
        output.print_md("")

def print_grouped_category_section(title, items):
    if not items:
        return

    grouped = group_by_category(items)

    output.print_md("## {}".format(title))
    output.print_md("")

    for category in CATEGORY_ORDER:
        cat_items = grouped.get(category, [])
        if not cat_items:
            continue

        output.print_md("### {} ({})".format(category, len(cat_items)))
        output.print_md("")
        for name in sorted(cat_items):
            output.print_md("- `{}`".format(name))
        output.print_md("")

def get_present_standard_worksets(existing_names):
    present = []
    for ws_name in STANDARD_WORKSETS:
        if ws_name in existing_names:
            present.append(ws_name)
    return present

def get_missing_standard_worksets(existing_names):
    missing = []
    for ws_name in STANDARD_WORKSETS:
        if ws_name not in existing_names:
            missing.append(ws_name)
    return missing

def get_extra_nonstandard_worksets(existing_names):
    extra = []
    for ws_name in existing_names:
        if ws_name not in STANDARD_WORKSETS:
            extra.append(ws_name)
    return extra

def build_quick_findings(nonstandard_names):
    generic = []
    linked = []
    shared = []
    unknown = []

    for name in nonstandard_names:
        category = classify_nonstandard_workset(name)
        if category == "GENERIC":
            generic.append(name)
        elif category == "LINK_RELATED":
            linked.append(name)
        elif category == "SHARED_LEVELS_GRIDS":
            shared.append(name)
        elif category == "UNKNOWN":
            unknown.append(name)

    return generic, linked, shared, unknown

def suggest_mapping(name):
    n = normalize_name(name)

    is_arc = (
        starts_with_any(n, ["ARC_", "ARQ_", "AR_", "A_"]) or
        has_token(n, "ARC") or
        has_token(n, "ARQ") or
        has_token(n, "A") or
        "ARQUITECT" in n
    )

    is_str = (
        starts_with_any(n, ["STR_", "ST_", "S_"]) or
        has_token(n, "STR") or
        has_token(n, "ST") or
        has_token(n, "S") or
        has_token(n, "STRUCTURAL") or
        "ESTRUCT" in n
    )

    is_mep = (
        starts_with_any(n, ["MEP_", "ME_", "M_"]) or
        has_token(n, "MEP") or
        has_token(n, "ME") or
        has_token(n, "M") or
        has_token(n, "MECHANICAL")
    )

    is_ele = (
        starts_with_any(n, ["ELE_", "EL_"]) or
        has_token(n, "ELE") or
        has_token(n, "EL") or
        has_token(n, "ELECTRICAL")
    )

    is_plm = (
        starts_with_any(n, ["PLM_", "PL_", "P_"]) or
        has_token(n, "PLM") or
        has_token(n, "PL") or
        has_token(n, "P") or
        has_token(n, "PLUMBING")
    )

    is_fur = (
        starts_with_any(n, ["FE_", "FUR_"]) or
        has_token(n, "FE") or
        has_token(n, "FUR") or
        has_token(n, "FURNITURE") or
        "MOBILIARIO" in n
    )

    is_site = (
        starts_with_any(n, ["SITE_"]) or
        has_token(n, "SITE") or
        "TOPO" in n
    )

    # shared levels / grids
    if "LEVELS" in n or "GRIDS" in n:
        if is_site:
            return None
        if is_arc:
            return "ARC_LEVELS_GRIDS"
        if is_str:
            return "STR_LEVELS_GRIDS"
        if is_mep:
            return "MEP_LEVELS_GRIDS"
        if is_ele:
            return "ELE_LEVELS_GRIDS"
        if is_plm:
            return "PLM_LEVELS_GRIDS"
        if is_fur:
            return None

    # linked-related
    if "LINK" in n or "LINKED" in n:
        if is_arc:
            return "LINK_ARC"
        if is_str:
            return "LINK_STR"
        if is_mep or "MECHANICAL" in n:
            return "LINK_MEP"
        if is_ele:
            return "LINK_ELE"
        if is_plm:
            return "LINK_PLM"
        if is_site:
            return "LINK_SITE"
        if is_fur:
            return "LINK_REF"
        return "LINK_REF"

    # model-related
    if is_arc:
        return "ARC_MODEL"

    if (
        is_str or
        "VIGAS" in n or
        "BEAM" in n or
        "BEAMS" in n or
        "COLUMN" in n or
        "COLUMNS" in n or
        "WALL" in n or
        "WALLS" in n or
        "REBAR" in n or
        "REBARS" in n or
        "FIERRO" in n or
        "ARMADURA" in n or
        "HORMIGON" in n or
        "CONCRETE" in n or
        "LOSAS" in n or
        "SLAB" in n or
        "SLABS" in n or
        "FUNDACION" in n or
        "FOUNDATION" in n or
        "FOUNDATIONS" in n or
        "CERCHA" in n or
        "TRUSS" in n or
        "PERFILES" in n or
        "METAL" in n or
        "STEEL" in n or
        "ESCALERA" in n or
        "STAIR" in n or
        "STAIRS" in n or
        "STRUCTURAL" in n
    ):
        return "STR_MODEL"

    if is_mep:
        return "MEP_MODEL"

    if is_ele:
        return "ELE_MODEL"

    if is_plm:
        return "PLM_MODEL"

    if is_site:
        return "LINK_SITE"

    if is_fur:
        return None

    return None

def build_mapping_lists(nonstandard_names):
    suggested = []
    manual = []

    for name in nonstandard_names:
        target = suggest_mapping(name)
        if target:
            suggested.append((name, target))
        else:
            manual.append(name)

    return suggested, manual

def print_mapping_section(title, items):
    if not items:
        return

    output.print_md("## {}".format(title))
    output.print_md("")
    for src, dst in sorted(items, key=lambda x: x[0]):
        output.print_md("- `{}` → `{}`".format(src, dst))
    output.print_md("")

def get_final_status(is_workshared, missing_count, nonstandard_count):
    if not is_workshared:
        return "FAILED"
    if missing_count == 0 and nonstandard_count == 0:
        return "CLEAN"
    if missing_count == 0:
        return "COMPLETE WITH REVIEW NEEDED"
    return "PARTIAL"

# ==================================================
# HEADER
# ==================================================

output.print_md("# MENVIC | ANALYZE WORKSETS — REPORT")
output.print_md("")
output.print_md("**Version:** v3")
output.print_md("")

# ==================================================
# VALIDATION
# ==================================================

if not doc.IsWorkshared:
    output.print_md("## Result")
    output.print_md("")
    output.print_md("The model is not **workshared**.")
    output.print_md("")
    output.print_md("## Final status")
    output.print_md("")
    output.print_md("**FAILED**")
    script.exit()

# ==================================================
# SCAN
# ==================================================

existing_names = get_existing_user_workset_names()

present_standard = get_present_standard_worksets(existing_names)
missing_standard = get_missing_standard_worksets(existing_names)
nonstandard_worksets = get_extra_nonstandard_worksets(existing_names)

generic_names, linked_names, shared_names, unknown_names = build_quick_findings(nonstandard_worksets)
suggested_mappings, manual_review_mapping = build_mapping_lists(nonstandard_worksets)

total_standard = len(STANDARD_WORKSETS)
present_count = len(present_standard)
missing_count = len(missing_standard)
nonstandard_count = len(nonstandard_worksets)
total_user_worksets = len(existing_names)

coverage_percent = 0.0
if total_standard > 0:
    coverage_percent = (float(present_count) / float(total_standard)) * 100.0

final_status = get_final_status(True, missing_count, nonstandard_count)

# ==================================================
# REPORT
# ==================================================

output.print_md("## Summary")
output.print_md("")
output.print_md("- **Total user worksets in model:** {}".format(total_user_worksets))
output.print_md("- **Total standard worksets expected:** {}".format(total_standard))
output.print_md("- **Standard worksets present:** {}".format(present_count))
output.print_md("- **Standard worksets missing:** {}".format(missing_count))
output.print_md("- **Extra non-standard worksets:** {}".format(nonstandard_count))
output.print_md("")

print_grouped_discipline_section("Standard worksets present", present_standard)

if missing_standard:
    print_grouped_discipline_section("Missing standard worksets", missing_standard)

output.print_md("## Standard coverage")
output.print_md("")
output.print_md("- **Coverage:** {} / {}".format(present_count, total_standard))
output.print_md("- **Coverage percent:** {:.1f}%".format(coverage_percent))
output.print_md("")

if nonstandard_worksets:
    output.print_md("## Quick findings")
    output.print_md("")
    output.print_md("- **Generic names:** {}".format(len(generic_names)))
    output.print_md("- **Linked-related names:** {}".format(len(linked_names)))
    output.print_md("- **Shared levels / grids variants:** {}".format(len(shared_names)))
    output.print_md("- **Unknown names:** {}".format(len(unknown_names)))
    output.print_md("")

    print_grouped_category_section("Non-standard worksets by category", nonstandard_worksets)
    print_mapping_section("Suggested standard mapping", suggested_mappings)

    if manual_review_mapping:
        output.print_md("## Manual review required")
        output.print_md("")
        for name in sorted(manual_review_mapping):
            output.print_md("- `{}`".format(name))
        output.print_md("")

output.print_md("## Interpretation")
output.print_md("")
if missing_count == 0 and nonstandard_count == 0:
    output.print_md("The model fully matches the pyMENVIC workset standard and does not contain additional non-standard user worksets.")
elif missing_count == 0 and nonstandard_count > 0:
    output.print_md("The model contains the full pyMENVIC standard, but it also includes additional non-standard worksets that should be reviewed before cleanup or consolidation.")
else:
    output.print_md("The model does not fully match the pyMENVIC standard yet. Missing standard worksets and non-standard worksets should be reviewed before consolidation.")
output.print_md("")

output.print_md("## Final status")
output.print_md("")
output.print_md("**{}**".format(final_status))
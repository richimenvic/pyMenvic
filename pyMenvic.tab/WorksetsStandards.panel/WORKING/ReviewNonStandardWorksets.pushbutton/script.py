# -*- coding: utf-8 -*-

__title__ = "Review Non-Standard Worksets"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
REVIEW NON-STANDARD WORKSETS
_____________________________________________________

Description:

Reviews all non-standard user worksets in the model and
classifies them into useful categories for manual cleanup.

_____________________________________________________
What the tool does:

• checks if model is workshared
• reads all user worksets in the model
• excludes pyMENVIC standard worksets
• classifies non-standard worksets
• identifies generic names
• identifies possible legacy names
• identifies possible linked-related names
• prints a formatted review report

_____________________________________________________
Output:

Prints a grouped review report for non-standard worksets.

_____________________________________________________
Usage:

1. Click the pyRevit button
2. Tool runs automatically
3. Review the output

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
    "LINKED FILES"
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
    "STR",
    "ST_",
    "S_",
    "MEP",
    "ME_",
    "ELE",
    "EL_",
    "PLM",
    "PL_",
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
    "TOPOSOLID"
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

def get_existing_user_workset_names():
    names = []
    collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    for ws in collector:
        names.append(ws.Name)
    return sorted(names)

def is_standard_workset(name):
    return name in STANDARD_WORKSETS

def contains_any(text, keywords):
    upper_text = normalize_name(text)
    for kw in keywords:
        if kw in upper_text:
            return True
    return False

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

def print_grouped_section(title, items):
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

def get_final_status(is_workshared, nonstandard_count):
    if not is_workshared:
        return "FAILED"
    if nonstandard_count == 0:
        return "CLEAN"
    return "REVIEW NEEDED"

# ==================================================
# HEADER
# ==================================================

output.print_md("# MENVIC | REVIEW NON-STANDARD WORKSETS — REPORT")
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

all_user_worksets = get_existing_user_workset_names()

nonstandard_worksets = []
standard_present = []

for ws_name in all_user_worksets:
    if is_standard_workset(ws_name):
        standard_present.append(ws_name)
    else:
        nonstandard_worksets.append(ws_name)

generic_names, linked_names, shared_names, unknown_names = build_quick_findings(nonstandard_worksets)

final_status = get_final_status(True, len(nonstandard_worksets))

# ==================================================
# REPORT
# ==================================================

output.print_md("## Summary")
output.print_md("")
output.print_md("- **Total user worksets in model:** {}".format(len(all_user_worksets)))
output.print_md("- **Standard pyMENVIC worksets present:** {}".format(len(standard_present)))
output.print_md("- **Non-standard worksets found:** {}".format(len(nonstandard_worksets)))
output.print_md("")

output.print_md("## Quick findings")
output.print_md("")
output.print_md("- **Generic names:** {}".format(len(generic_names)))
output.print_md("- **Linked-related names:** {}".format(len(linked_names)))
output.print_md("- **Shared levels / grids variants:** {}".format(len(shared_names)))
output.print_md("- **Unknown names:** {}".format(len(unknown_names)))
output.print_md("")

print_grouped_section("Non-standard worksets by category", nonstandard_worksets)

if generic_names:
    output.print_md("## Generic names to review")
    output.print_md("")
    for name in sorted(generic_names):
        output.print_md("- `{}`".format(name))
    output.print_md("")

if unknown_names:
    output.print_md("## Unknown names to review manually")
    output.print_md("")
    for name in sorted(unknown_names):
        output.print_md("- `{}`".format(name))
    output.print_md("")

output.print_md("## Interpretation")
output.print_md("")
if len(nonstandard_worksets) == 0:
    output.print_md("The model contains only the standard pyMENVIC user worksets.")
else:
    output.print_md("The model includes additional non-standard worksets. This does not automatically mean the model is incorrect, but these names should be reviewed before cleanup, renaming, or consolidation.")
output.print_md("")

output.print_md("## Final status")
output.print_md("")
output.print_md("**{}**".format(final_status))
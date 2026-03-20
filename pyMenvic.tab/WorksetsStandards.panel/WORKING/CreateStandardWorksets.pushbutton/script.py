# -*- coding: utf-8 -*-

__title__ = "Create Standard Worksets"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
CREATE STANDARD WORKSETS
_____________________________________________________

Description:

Creates the standard pyMENVIC worksets if they do not already exist.

_____________________________________________________
What the tool does:

• checks if model is workshared
• creates missing standard worksets
• reports created, existing and failed worksets
• groups results by discipline
• reports final standard coverage
• reports final execution status

_____________________________________________________
Output:

Prints a formatted report grouped by discipline.

_____________________________________________________
Usage:

1. Click the pyRevit button
2. Tool runs automatically
3. Review the output

_____________________________________________________

Author: Ricardo J. Mendieta
"""

from Autodesk.Revit.DB import Workset, WorksetKind, WorksetTable, FilteredWorksetCollector, Transaction
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

# ==================================================
# HELPERS
# ==================================================

def get_error_message(ex):
    try:
        return str(ex).splitlines()[0]
    except:
        return "Unknown error"

def get_group_name(ws_name):
    if ws_name.startswith("ARC_"):
        return "ARC"
    elif ws_name.startswith("STR_"):
        return "STR"
    elif ws_name.startswith("MEP_"):
        return "MEP"
    elif ws_name.startswith("ELE_"):
        return "ELE"
    elif ws_name.startswith("PLM_"):
        return "PLM"
    elif ws_name.startswith("LINK_"):
        return "LINK"
    else:
        return "OTHER"

def group_names(items):
    grouped = {}
    for group_name in DISCIPLINE_ORDER:
        grouped[group_name] = []

    for item in items:
        group_name = get_group_name(item)
        grouped[group_name].append(item)

    return grouped

def print_grouped_section(title, items):
    if not items:
        return

    grouped = group_names(items)

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

def print_failed_section(title, failed_items):
    if not failed_items:
        return

    grouped = {}
    for group_name in DISCIPLINE_ORDER:
        grouped[group_name] = []

    for name, reason in failed_items:
        group_name = get_group_name(name)
        grouped[group_name].append((name, reason))

    output.print_md("## {}".format(title))
    output.print_md("")

    for group_name in DISCIPLINE_ORDER:
        group_items = grouped.get(group_name, [])
        if not group_items:
            continue

        output.print_md("### {} ({})".format(group_name, len(group_items)))
        output.print_md("")
        for name, reason in sorted(group_items, key=lambda x: x[0]):
            output.print_md("- `{}` | {}".format(name, reason))
        output.print_md("")

def get_existing_user_workset_names():
    names = []
    collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    for ws in collector:
        names.append(ws.Name)
    return names

def get_missing_standard_worksets(existing_names):
    missing = []
    for ws_name in STANDARD_WORKSETS:
        if ws_name not in existing_names:
            missing.append(ws_name)
    return missing

def get_final_status(is_workshared, failed_count, coverage_missing_count, processed_count):
    if not is_workshared:
        return "FAILED"

    if failed_count > 0 and processed_count > 0:
        return "PARTIAL"

    if failed_count > 0 and processed_count == 0:
        return "FAILED"

    if coverage_missing_count == 0:
        return "COMPLETE"

    return "PARTIAL"

# ==================================================
# HEADER
# ==================================================

output.print_md("# MENVIC | CREATE STANDARD WORKSETS — RUN")
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
# PROCESS
# ==================================================

created = []
existing_ws = []
failed = []

t = Transaction(doc, "Create Standard Worksets")
t.Start()

for ws_name in STANDARD_WORKSETS:
    try:
        if WorksetTable.IsWorksetNameUnique(doc, ws_name):
            Workset.Create(doc, ws_name)
            created.append(ws_name)
        else:
            existing_ws.append(ws_name)
    except Exception as ex:
        failed.append((ws_name, get_error_message(ex)))

t.Commit()

# ==================================================
# FINAL COVERAGE
# ==================================================

final_existing_names = get_existing_user_workset_names()
final_missing_standard = get_missing_standard_worksets(final_existing_names)

total_standard = len(STANDARD_WORKSETS)
present_standard = total_standard - len(final_missing_standard)
processed_count = len(created) + len(existing_ws) + len(failed)

final_status = get_final_status(
    True,
    len(failed),
    len(final_missing_standard),
    processed_count
)

# ==================================================
# REPORT
# ==================================================

output.print_md("## Summary")
output.print_md("")
output.print_md("- **Total standard worksets checked:** {}".format(total_standard))
output.print_md("- **Created:** {}".format(len(created)))
output.print_md("- **Already existing:** {}".format(len(existing_ws)))
output.print_md("- **Failed:** {}".format(len(failed)))
output.print_md("")

print_grouped_section("Created worksets", created)
print_grouped_section("Already existing worksets", existing_ws)
print_failed_section("Errors", failed)

output.print_md("## Standard coverage")
output.print_md("")
output.print_md("- **Present after run:** {} / {}".format(present_standard, total_standard))
output.print_md("- **Missing after run:** {}".format(len(final_missing_standard)))
output.print_md("")

if final_missing_standard:
    print_grouped_section("Missing standard worksets after run", final_missing_standard)

output.print_md("## Final status")
output.print_md("")
output.print_md("**{}**".format(final_status))
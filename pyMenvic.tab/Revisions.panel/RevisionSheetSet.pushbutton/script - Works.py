# -*- coding: utf-8 -*-
# # pylint: disable=E0401,C0103,C0111
from pyrevit import revit
from pyrevit import forms
from pyrevit import script


output = script.get_output()

# ----------------------------
# INPUTS
# ----------------------------
revisions = forms.select_revisions(button_name='Create Sheet Set',
                                   multiple=True)

# ----------------------------
# PROCESS
# ----------------------------
selected_switch = None
match_any = False
rev_sheetset = None
rev_sheetset_count = 0
empty_sheets = []

if revisions:
    if len(revisions) > 1:
        selected_switch = forms.CommandSwitchWindow.show(
            ['Matching ANY revision', 'Matching ALL revisions'],
            message='Pick an option:'
        )
    else:
        selected_switch = 'Matching ALL revisions'

    if selected_switch:
        match_any = (selected_switch == 'Matching ANY revision')

        with revit.Transaction('Create Revision Sheet Set'):
            rev_sheetset = revit.create.create_revision_sheetset(
                revisions,
                match_any=match_any
            )

        # rev_sheetset es iterable (sheets)
        for sheet in rev_sheetset:
            rev_sheetset_count += 1
            if revit.query.is_sheet_empty(sheet):
                empty_sheets.append(sheet)

        if empty_sheets:
            print('These sheets do not have any model contents and seem to be '
                  'placeholders for other content:')
            for esheet in empty_sheets:
                revit.report.print_sheet(esheet)
# ----------------------------
# GET SHEET SET NAME
# ----------------------------
from Autodesk.Revit.DB import FilteredElementCollector, ViewSheetSet

sheetset_name = "Unknown"

try:
    collector = FilteredElementCollector(revit.doc).OfClass(ViewSheetSet)
    for ss in collector:
        try:
            if ss.Name:
                sheetset_name = ss.Name
        except:
            pass
except:
    pass


# ----------------------------
# FINAL REPORT
# ----------------------------
def _safe_str(x):
    try:
        return "{}".format(x)
    except:
        return "<unprintable>"

def _get_revision_label(rev):
    seq = ""
    desc = ""
    try:
        seq = _safe_str(rev.SequenceNumber)
    except:
        pass
    try:
        desc = _safe_str(rev.Description)
    except:
        pass

    if seq and desc:
        return "{} - {}".format(seq, desc)
    elif desc:
        return desc
    else:
        return "Revision Id {}".format(rev.Id.IntegerValue)


# Build revision label
rev_labels = []
if revisions:
    for r in revisions:
        rev_labels.append(_get_revision_label(r))

if len(rev_labels) == 1:
    rev_one = rev_labels[0]
elif len(rev_labels) > 1:
    rev_one = "{} revisions".format(len(rev_labels))
else:
    rev_one = "None"


# Checks
ok_selected = bool(revisions)
ok_mode = bool(selected_switch)
ok_created = (rev_sheetset is not None)
ok_sheets = (rev_sheetset_count > 0)
ok_empty = (len(empty_sheets) == 0)

overall_ok = ok_selected and ok_mode and ok_created and ok_sheets
status_txt = "OK" if overall_ok else "FAIL"

mode_txt = selected_switch if selected_switch else "None"


# ----------------------------
# PRINT REPORT
# ----------------------------
output.print_md("### pyMenvic | Revision Sheet Set - RUN")

# Revision grande
output.print_md("## {}".format(rev_one))

# Sheet set creado
output.print_md("# Sheet Set: {}".format(sheetset_name))

# Resumen
output.print_md(
    "Status: {} | Mode: {} | Sheets: {} | Empty: {}".format(
        status_txt,
        mode_txt,
        rev_sheetset_count,
        len(empty_sheets)
    )
)


# ----------------------------
# GRAPHIC CHECK
# ----------------------------
checks = []

def add_check(label, ok, detail):
    checks.append([label, "[OK]" if ok else "[FAIL]", detail])

add_check("Revisions selected", ok_selected,
          "{}".format(len(revisions) if revisions else 0))

add_check("Match option picked", ok_mode,
          "{}".format(mode_txt))

add_check("Sheet set created", ok_created,
          "Yes" if ok_created else "No")

add_check("Sheets in set > 0", ok_sheets,
          "{}".format(rev_sheetset_count))

add_check("No empty sheets", ok_empty,
          "{}".format(len(empty_sheets)))

output.print_md("#### Graphic Check")
output.print_table(
    table_data=checks,
    columns=["Check", "Status", "Detail"]
)
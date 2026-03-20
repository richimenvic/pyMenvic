# -*- coding: utf-8 -*-
# pylint: disable=E0401,C0103,C0111

import Autodesk.Revit.DB as DB
from pyrevit import revit, forms, script


doc = revit.doc
output = script.get_output()

# ----------------------------
# INPUTS
# ----------------------------
revisions = forms.select_revisions(
    button_name='Create Sheet Set',
    multiple=True
)

# ----------------------------
# PROCESS VARS
# ----------------------------
selected_switch = None
match_any = False

rev_sheetset = None
rev_sheetset_count = 0
empty_sheets = []

match_count = 0
error_reason = None
sheetset_name = "Not created"

# ----------------------------
# HELPERS
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
        try:
            return "Revision Id {}".format(rev.Id.IntegerValue)
        except:
            return "Revision"

def _count_matching_sheets(selected_revision_ids, match_any_flag):
    """Counts ViewSheets that match selected revisions (ANY or ALL)."""
    count = 0

    sheets = DB.FilteredElementCollector(doc) \
               .OfClass(DB.ViewSheet) \
               .WhereElementIsNotElementType()

    for sh in sheets:
        try:
            sh_rev_ids = sh.GetAllRevisionIds()   # ICollection[ElementId]
        except:
            continue

        if not sh_rev_ids:
            continue

        if match_any_flag:
            found = False
            for rid in selected_revision_ids:
                if sh_rev_ids.Contains(rid):
                    found = True
                    break
            if found:
                count += 1
        else:
            all_found = True
            for rid in selected_revision_ids:
                if not sh_rev_ids.Contains(rid):
                    all_found = False
                    break
            if all_found:
                count += 1

    return count

def _try_get_current_sheetset_name():
    """Best-effort: reads current ViewSheetSet name from PrintManager."""
    try:
        pm = doc.PrintManager
        vss = pm.ViewSheetSetting
        # CurrentViewSheetSet is usually what create_revision_sheetset sets
        try:
            css = vss.CurrentViewSheetSet
            if css and css.Name:
                return css.Name
        except:
            pass

        # Fallback: sometimes InSession has the set info
        try:
            ins = vss.InSession
            if ins and ins.Name:
                return ins.Name
        except:
            pass
    except:
        pass

    return None


# ----------------------------
# MAIN
# ----------------------------
if revisions:
    # pick mode
    if len(revisions) > 1:
        selected_switch = forms.CommandSwitchWindow.show(
            ['Matching ANY revision', 'Matching ALL revisions'],
            message='Pick an option:'
        )
    else:
        selected_switch = 'Matching ALL revisions'

    if selected_switch:
        match_any = (selected_switch == 'Matching ANY revision')

        # PRE-CHECK: count matches
        sel_rev_ids = []
        for r in revisions:
            try:
                sel_rev_ids.append(r.Id)
            except:
                pass

        match_count = _count_matching_sheets(sel_rev_ids, match_any)

        if match_count == 0:
            error_reason = "No sheets match the selected revision(s)."
        else:
            # SAFE CREATE
            try:
                with revit.Transaction('Create Revision Sheet Set'):
                    rev_sheetset = revit.create.create_revision_sheetset(
                        revisions,
                        match_any=match_any
                    )
            except Exception as ex:
                rev_sheetset = None
                error_reason = "Failed to create sheet set: {}".format(ex)

            # Get name (only if created)
            if rev_sheetset:
                name_now = _try_get_current_sheetset_name()
                if name_now:
                    sheetset_name = name_now
                else:
                    sheetset_name = "Created (name unknown)"

            # Count sheets & empty (only if rev_sheetset iterable)
            if rev_sheetset:
                for sheet in rev_sheetset:
                    rev_sheetset_count += 1
                    try:
                        if revit.query.is_sheet_empty(sheet):
                            empty_sheets.append(sheet)
                    except:
                        pass
            else:
                # if creation failed but precheck had matches, show that info
                rev_sheetset_count = match_count

            # Optional: list empty sheets (console print)
            if empty_sheets:
                print('These sheets do not have any model contents and seem to be placeholders for other content:')
                for esheet in empty_sheets:
                    try:
                        revit.report.print_sheet(esheet)
                    except:
                        pass

# ----------------------------
# FINAL REPORT
# ----------------------------
# Revision label
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

mode_txt = selected_switch if selected_switch else "None"

ok_selected = bool(revisions)
ok_mode = bool(selected_switch)
ok_created = (rev_sheetset is not None)
ok_sheets = (rev_sheetset_count > 0)
ok_empty = (len(empty_sheets) == 0)

overall_ok = ok_selected and ok_mode and ok_created and ok_sheets and (error_reason is None)
status_txt = "OK" if overall_ok else "FAIL"

# PRINT
output.print_md("### pyMenvic | Revision Sheet Set - RUN")

# Sheet Set grande
output.print_md("# Sheet Set: {}".format(sheetset_name))

# Revision grande (subtitulo)
output.print_md("## {}".format(rev_one))

# Resumen
output.print_md("Status: {} | Mode: {} | Sheets: {} | Empty: {}".format(
    status_txt,
    mode_txt,
    rev_sheetset_count,
    len(empty_sheets)
))

if error_reason:
    output.print_md("**Note:** {}".format(error_reason))

# If multiple revisions, list them
if len(rev_labels) > 1:
    output.print_md("**Revisions included:**")
    for rl in rev_labels:
        output.print_md("- {}".format(rl))

# GRAPHIC CHECK (single)
checks = []

def add_check(label, ok, detail):
    checks.append([label, "[OK]" if ok else "[FAIL]", detail])

add_check("Revisions selected", ok_selected, "{}".format(len(revisions) if revisions else 0))
add_check("Match option picked", ok_mode, "{}".format(mode_txt))
add_check("Has matches (pre-check)", (match_count > 0), "{}".format(match_count))
add_check("Sheet set created", ok_created, "Yes" if ok_created else "No")
add_check("Sheets in set > 0", ok_sheets, "{}".format(rev_sheetset_count))
add_check("No empty sheets", ok_empty, "{}".format(len(empty_sheets)))

output.print_md("#### Graphic Check")
output.print_table(
    table_data=checks,
    columns=["Check", "Status", "Detail"]
)
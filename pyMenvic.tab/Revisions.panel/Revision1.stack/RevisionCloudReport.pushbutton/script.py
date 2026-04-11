# -*- coding: utf-8 -*-
__title__ = "Revision Cloud Report"
import re
from pyrevit import revit, DB, script

doc = revit.doc
output = script.get_output()
UNKNOWN_SHEET_LABEL = "Unplaced / Unknown"
UNKNOWN_SHEET_NAME = "Unknown Sheet"
UNKNOWN_VIEW_NAME = "Unknown View"

# ============================================================
# HELPERS
# ============================================================

def natural_key(s):
    if s is None:
        return []
    s = str(s)
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]


def safe_as_string(param):
    if not param:
        return None
    try:
        return param.AsString()
    except Exception:
        return None


def get_view_sheet_number(view):
    if not isinstance(view, DB.View):
        return None

    sheet_number = safe_as_string(view.get_Parameter(DB.BuiltInParameter.VIEWER_SHEET_NUMBER))
    if sheet_number and sheet_number != '---':
        return sheet_number

    return None


def get_revision_number(revision, sheet):
    if not revision:
        return "No Number"

    if sheet:
        try:
            rev_number_on_sheet = sheet.GetRevisionNumberOnSheet(revision.Id)
            if rev_number_on_sheet:
                return str(rev_number_on_sheet)
        except Exception:
            pass

    try:
        rev_number = revision.RevisionNumber
        if rev_number:
            return str(rev_number)
    except Exception:
        pass

    try:
        seq = revision.SequenceNumber
        if seq is not None:
            return str(seq)
    except Exception:
        pass

    return "No Number"


def get_comment_text(revision_cloud):
    comment = safe_as_string(
        revision_cloud.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    )
    return comment or ""

# Cache de sheets (evita collectors repetidos)
all_sheets = DB.FilteredElementCollector(doc) \
               .OfCategory(DB.BuiltInCategory.OST_Sheets) \
               .WhereElementIsNotElementType() \
               .ToElements()

sheet_by_number = {}
for s in all_sheets:
    sheet_number = getattr(s, 'SheetNumber', None)
    if sheet_number:
        sheet_by_number[sheet_number] = s

def get_sheet_info(revision_cloud):
    sheet_name = UNKNOWN_SHEET_NAME
    sheet_number = UNKNOWN_SHEET_LABEL
    sheet_key = "unknown:{}".format(revision_cloud.OwnerViewId.IntegerValue)
    sheet = None

    owner_view = doc.GetElement(revision_cloud.OwnerViewId)
    if not owner_view:
        return sheet_name, sheet_number, sheet_key, sheet, UNKNOWN_VIEW_NAME

    view_name = getattr(owner_view, 'Name', UNKNOWN_VIEW_NAME)

    if isinstance(owner_view, DB.ViewSheet):
        sheet = owner_view
        sheet_name = owner_view.Title
        sheet_number = owner_view.SheetNumber
        sheet_key = sheet_number

    elif isinstance(owner_view, DB.View):
        owner_sheet_number = get_view_sheet_number(owner_view)
        if owner_sheet_number:
            sheet = sheet_by_number.get(owner_sheet_number)
            if sheet:
                sheet_name = sheet.Title
                sheet_number = sheet.SheetNumber
                sheet_key = sheet_number
            else:
                sheet_number = owner_sheet_number
                sheet_name = "Missing Sheet for {}".format(owner_sheet_number)
                sheet_key = "missing:{}".format(owner_sheet_number)

    return sheet_name, sheet_number, sheet_key, sheet, view_name

# ============================================================
# COLLECT / SCAN
# ============================================================

revision_clouds = DB.FilteredElementCollector(doc) \
                    .OfCategory(DB.BuiltInCategory.OST_RevisionClouds) \
                    .WhereElementIsNotElementType() \
                    .ToElements()

# ============================================================
# PROCESS (Sheet -> View)
# ============================================================

data_by_sheet = {}
# data_by_sheet[sheet_number] = {'name': sheet_name, 'views': {view_name: [clouds...] } }

for revc in revision_clouds:
    sheet_name, sheet_number, sheet_key, sheet, view_name = get_sheet_info(revc)

    if sheet_key not in data_by_sheet:
        data_by_sheet[sheet_key] = {'number': sheet_number, 'name': sheet_name, 'views': {}}

    if view_name not in data_by_sheet[sheet_key]['views']:
        data_by_sheet[sheet_key]['views'][view_name] = []

    data_by_sheet[sheet_key]['views'][view_name].append((revc, sheet))

# ============================================================
# REPORT (tabla por vista)
# ============================================================

sorted_sheets = sorted(data_by_sheet.items(), key=lambda x: natural_key(x[1]['number']))

total_clouds = 0

for _, sheet_data in sorted_sheets:
    output.print_md('### **Sheet:** {} - {}\n'.format(sheet_data['number'], sheet_data['name']))

    view_items = list(sheet_data['views'].items())
    view_items_sorted = sorted(view_items, key=lambda x: natural_key(x[0]))

    for view_name, clouds in view_items_sorted:
        output.print_md('#### **View:** {}\n'.format(view_name))

        # Orden estable de clouds por Id
        clouds.sort(key=lambda item: item[0].Id.IntegerValue)

        rows = []
        headers = ["CloudId", "Revision", "Comment"]

        for revc, sheet in clouds:
            total_clouds += 1

            cloud_id_link = output.linkify([revc.Id])

            # Revision Number (si existe)
            revision = doc.GetElement(revc.RevisionId)
            rev_number = get_revision_number(revision, sheet)

            # Comment idioma-proof
            comment = get_comment_text(revc)

            rows.append([cloud_id_link, rev_number, comment])

        if rows:
            output.print_table(table_data=rows, columns=headers)
        else:
            output.print_md("_No revision clouds in this view._\n")

output.print_md('\n**SEARCH COMPLETED — {} revision clouds found.**'.format(total_clouds))

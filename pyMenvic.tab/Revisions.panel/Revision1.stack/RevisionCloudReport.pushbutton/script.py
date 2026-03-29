# -*- coding: utf-8 -*-
__title__ = "Revision Cloud Report"
import re
from pyrevit import revit, DB, script

doc = revit.doc
output = script.get_output()

# ============================================================
# HELPERS
# ============================================================

def natural_key(s):
    if s is None:
        return []
    s = str(s)
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]

# Cache de sheets (evita collectors repetidos)
all_sheets = DB.FilteredElementCollector(doc) \
               .OfCategory(DB.BuiltInCategory.OST_Sheets) \
               .WhereElementIsNotElementType() \
               .ToElements()

sheet_by_number = {}
for s in all_sheets:
    try:
        sheet_by_number[s.SheetNumber] = s
    except:
        pass

def get_sheet_info(revision_cloud):
    sheet_name = "Unknown Sheet"
    sheet_number = "Unknown Number"

    owner_view = doc.GetElement(revision_cloud.OwnerViewId)
    if not owner_view:
        return sheet_name, sheet_number

    if isinstance(owner_view, DB.ViewSheet):
        sheet_name = owner_view.Title
        sheet_number = owner_view.SheetNumber

    elif isinstance(owner_view, DB.View):
        p = owner_view.LookupParameter('Sheet Number')
        sheet_id = p.AsString() if p else None
        if sheet_id:
            sheet = sheet_by_number.get(sheet_id)
            if sheet:
                sheet_name = sheet.Title
                sheet_number = sheet.SheetNumber

    if sheet_name and " - " in sheet_name:
        sheet_name = sheet_name.split(" - ")[-1].strip()

    return sheet_name, sheet_number

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
    sheet_name, sheet_number = get_sheet_info(revc)

    if sheet_number not in data_by_sheet:
        data_by_sheet[sheet_number] = {'name': sheet_name, 'views': {}}

    view = doc.GetElement(revc.OwnerViewId)
    view_name = view.Name if view else "Unknown View"

    if view_name not in data_by_sheet[sheet_number]['views']:
        data_by_sheet[sheet_number]['views'][view_name] = []

    data_by_sheet[sheet_number]['views'][view_name].append(revc)

# ============================================================
# REPORT (tabla por vista)
# ============================================================

sorted_sheets = sorted(data_by_sheet.items(), key=lambda x: natural_key(x[0]))

total_clouds = 0

for sheet_number, sheet_data in sorted_sheets:
    output.print_md('### **Sheet:** {} - {}\n'.format(sheet_number, sheet_data['name']))

    view_items = list(sheet_data['views'].items())
    view_items_sorted = sorted(view_items, key=lambda x: natural_key(x[0]))

    for view_name, clouds in view_items_sorted:
        output.print_md('#### **View:** {}\n'.format(view_name))

        # Orden estable de clouds por Id
        try:
            clouds.sort(key=lambda c: c.Id.IntegerValue)
        except:
            pass

        rows = []
        headers = ["CloudId", "Revision", "Comment"]

        for revc in clouds:
            total_clouds += 1

            cloud_id_link = output.linkify([revc.Id])

            # Revision Number (si existe)
            revision = doc.GetElement(revc.RevisionId)
            rev_number_param = revision.LookupParameter('Revision Number') if revision else None
            rev_number = rev_number_param.AsString() if rev_number_param else "No Number"

            # Comment idioma-proof
            comment_param = revc.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            comment = comment_param.AsString() if comment_param else ""
            if comment is None:
                comment = ""

            rows.append([cloud_id_link, rev_number, comment])

        if rows:
            output.print_table(table_data=rows, columns=headers)
        else:
            output.print_md("_No revision clouds in this view._\n")

output.print_md('\n**SEARCH COMPLETED — {} revision clouds found.**'.format(total_clouds))
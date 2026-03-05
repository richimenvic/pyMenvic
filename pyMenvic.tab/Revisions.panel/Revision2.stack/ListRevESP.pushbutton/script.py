# -*- coding: utf-8 -*-
import re
from pyrevit import revit, DB, script, forms

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

def safe_as_string(param):
    try:
        return param.AsString() if param else ""
    except:
        return ""

def make_revision_display(rev_number, rev_description, rev_date):
    rn = (rev_number or "").strip()
    if rn.lower().startswith("rev"):
        return "{} | {} | {}".format(rn, rev_description, rev_date)
    else:
        return "{} - {} - {}".format(rn, rev_description, rev_date)

# ============================================================
# CACHE SHEETS
# ============================================================

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

# ============================================================
# COLLECT REVISIONS + UI
# ============================================================

revisions = DB.FilteredElementCollector(doc) \
              .OfCategory(DB.BuiltInCategory.OST_Revisions) \
              .WhereElementIsNotElementType() \
              .ToElements()

revision_items = []  # (display, id, rev_number)
for rev in revisions:
    pnum = rev.LookupParameter('Revision Number')
    pdate = rev.LookupParameter('Revision Date')
    pdesc = rev.LookupParameter('Revision Description')

    rev_number = safe_as_string(pnum)
    rev_date = safe_as_string(pdate)
    rev_description = safe_as_string(pdesc)

    display = make_revision_display(rev_number, rev_description, rev_date)
    revision_items.append((display, rev.Id, rev_number))

# Orden natural por número de revisión
revision_items_sorted = sorted(revision_items, key=lambda x: natural_key(x[2]))

display_list = [x[0] for x in revision_items_sorted]
selected_display = forms.SelectFromList.show(
    display_list,
    title='Select Revision',
    button_name='Select',
    multiple=False
)

if not selected_display:
    script.exit()

selected_revision_id = next(x[1] for x in revision_items_sorted if x[0] == selected_display)

# ============================================================
# COLLECT REVISION CLOUDS
# ============================================================

revision_clouds = DB.FilteredElementCollector(doc) \
                    .OfCategory(DB.BuiltInCategory.OST_RevisionClouds) \
                    .WhereElementIsNotElementType() \
                    .ToElements()

# ============================================================
# SHEET INFO (incluye Spanish Sheet)
# ============================================================

def get_sheet_info(revision_cloud):
    sheet_name = "Unknown Sheet"
    sheet_number = "Unknown Number"
    spanish_sheet = ""

    owner_view = doc.GetElement(revision_cloud.OwnerViewId)
    if not owner_view:
        return sheet_name, sheet_number, spanish_sheet

    if isinstance(owner_view, DB.ViewSheet):
        sheet_name = owner_view.Title.replace("Sheet: ", "").strip()
        sheet_number = owner_view.SheetNumber

        pspan = owner_view.LookupParameter('Spanish Sheet')
        spanish_sheet = safe_as_string(pspan)

    elif isinstance(owner_view, DB.View):
        p = owner_view.LookupParameter('Sheet Number')
        sheet_id = safe_as_string(p)

        if sheet_id:
            sheet = sheet_by_number.get(sheet_id)
            if sheet:
                sheet_name = sheet.Title.replace("Sheet: ", "").strip()
                sheet_number = sheet.SheetNumber

                pspan = sheet.LookupParameter('Spanish Sheet')
                spanish_sheet = safe_as_string(pspan)

    return sheet_name, sheet_number, spanish_sheet

# ============================================================
# PROCESS (solo sheets únicos para esa revisión)
# ============================================================

# sheet_number -> (sheet_name, spanish_sheet)
revision_clouds_by_sheet = {}

for revc in revision_clouds:
    if revc.RevisionId != selected_revision_id:
        continue

    sheet_name, sheet_number, spanish_sheet = get_sheet_info(revc)

    if " - " in sheet_name:
        sheet_name = sheet_name.split(" - ")[-1].strip()

    if sheet_number not in revision_clouds_by_sheet:
        revision_clouds_by_sheet[sheet_number] = (sheet_name, spanish_sheet)

# Orden natural de sheets
sorted_sheets = sorted(revision_clouds_by_sheet.items(), key=lambda x: natural_key(x[0]))

# ============================================================
# REPORT (sin conteo)
# ============================================================

output.print_md('### **Listado de Sheets**')

if not sorted_sheets:
    output.print_md('_No sheets found for this revision._')
else:
    for sheet_number, (sheet_name, spanish_sheet) in sorted_sheets:
        # Mostrar Spanish Sheet solo si existe
        if spanish_sheet:
            output.print_md('{} - {} ({})'.format(sheet_number, sheet_name, spanish_sheet))
        else:
            output.print_md('{} - {}'.format(sheet_number, sheet_name))

output.print_md('\n**SEARCH COMPLETED.**')
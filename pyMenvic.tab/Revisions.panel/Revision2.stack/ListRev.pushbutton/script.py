# -*- coding: utf-8 -*-
__title__ = "Sheet List"
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

# ============================================================
# COLLECT / SCAN (Revisions)
# ============================================================

revisions = DB.FilteredElementCollector(doc) \
              .OfCategory(DB.BuiltInCategory.OST_Revisions) \
              .WhereElementIsNotElementType() \
              .ToElements()

# Crear lista (display, id, rev_number) para ordenar bien y no perder el Id
revision_items = []
for rev in revisions:
    pnum = rev.LookupParameter('Revision Number')
    pdate = rev.LookupParameter('Revision Date')
    pdesc = rev.LookupParameter('Revision Description')

    rev_number = safe_as_string(pnum)
    rev_date = safe_as_string(pdate)
    rev_description = safe_as_string(pdesc)

    # Display como en la captura: "Rev 1 | Addendum 1 | 07/29/2024"
# (si rev_number ya trae "1", "2", etc.)
    display = "{} - {} - {}".format(rev_number, rev_description, rev_date)
    revision_items.append((display, rev.Id, rev_number))

# Ordenar por número de revisión (natural)
revision_items_sorted = sorted(revision_items, key=lambda x: natural_key(x[2]))

# UI: mostrar lista ya ordenada
display_list = [x[0] for x in revision_items_sorted]
selected_display = forms.SelectFromList.show(
    display_list,
    title='Select Revision',
    button_name='Select',
    multiple=False
)

if not selected_display:
    script.exit()

# Obtener el Id correcto (desde la lista ordenada)
selected_revision_id = next(x[1] for x in revision_items_sorted if x[0] == selected_display)

# ============================================================
# COLLECT / SCAN (Revision Clouds)
# ============================================================

revision_clouds = DB.FilteredElementCollector(doc) \
                    .OfCategory(DB.BuiltInCategory.OST_RevisionClouds) \
                    .WhereElementIsNotElementType() \
                    .ToElements()

# Función para obtener el nombre y número de la hoja donde se encuentra la nube de revisión
def get_sheet_info(revision_cloud):
    sheet_name = "Unknown Sheet"
    sheet_number = "Unknown Number"

    owner_view = doc.GetElement(revision_cloud.OwnerViewId)
    if not owner_view:
        return sheet_name, sheet_number

    if isinstance(owner_view, DB.ViewSheet):
        sheet_name = owner_view.Title.replace("Sheet: ", "").strip()
        sheet_number = owner_view.SheetNumber

    elif isinstance(owner_view, DB.View):
        p = owner_view.LookupParameter('Sheet Number')
        sheet_id = safe_as_string(p)
        if sheet_id:
            sheet = sheet_by_number.get(sheet_id)
            if sheet:
                sheet_name = sheet.Title.replace("Sheet: ", "").strip()
                sheet_number = sheet.SheetNumber

    return sheet_name, sheet_number

# ============================================================
# PROCESS (Sheets con clouds de la revisión seleccionada)
# ============================================================

# sheet_number -> {'name': sheet_name, 'count': n}
revision_clouds_by_sheet = {}

for revc in revision_clouds:
    if revc.RevisionId != selected_revision_id:
        continue

    sheet_name, sheet_number = get_sheet_info(revc)

    if " - " in sheet_name:
        sheet_name = sheet_name.split(" - ")[-1].strip()

    if sheet_number not in revision_clouds_by_sheet:
        revision_clouds_by_sheet[sheet_number] = {'name': sheet_name, 'count': 0}

    revision_clouds_by_sheet[sheet_number]['count'] += 1

# Orden natural de sheets (A2 antes que A10)
sorted_sheets = sorted(revision_clouds_by_sheet.items(), key=lambda x: natural_key(x[0]))

# ============================================================
# REPORT
# ============================================================

output.print_md('### **Sheets with Clouds**')

if not sorted_sheets:
    output.print_md('_No sheets found for this revision._')
else:
    for sheet_number, data in sorted_sheets:
        output.print_md('{} - {}'.format(sheet_number, data['name']))
        
output.print_md('\n**SEARCH COMPLETED.**')
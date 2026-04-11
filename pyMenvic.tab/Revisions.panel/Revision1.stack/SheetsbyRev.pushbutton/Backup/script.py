# -*- coding: utf-8 -*-
__title__ = "Sheets by Rev"

import re
from pyrevit import revit, DB, script, forms

# --- Natural sort helper: "2" < "10", "A2" < "A10"
def natural_key(s):
    if s is None:
        return []
    s = str(s)
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]

doc = revit.doc
output = script.get_output()

# Recopilar todas las revisiones
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

    rev_number = pnum.AsString() if pnum else ""
    rev_date = pdate.AsString() if pdate else ""
    rev_description = pdesc.AsString() if pdesc else ""

    # Display más legible (solo UI)
    display = "{} - {} - {}".format(rev_number, rev_description, rev_date)
    revision_items.append((display, rev.Id, rev_number))

# Ordenar por número de revisión (natural)
revision_items_sorted = sorted(revision_items, key=lambda x: natural_key(x[2]))

# Mostrar formulario con la lista ya ordenada (solo textos)
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

# Pre-cargar la revisión seleccionada (evita lookups repetidos)
selected_revision = doc.GetElement(selected_revision_id)
rev_number_param = selected_revision.LookupParameter('Revision Number') if selected_revision else None
selected_rev_number = rev_number_param.AsString() if rev_number_param else "No Number"

# Cache de sheets para evitar collectors repetidos
all_sheets = DB.FilteredElementCollector(doc) \
               .OfCategory(DB.BuiltInCategory.OST_Sheets) \
               .WhereElementIsNotElementType() \
               .ToElements()

sheet_by_number = {}
for s in all_sheets:
    sheet_by_number[s.SheetNumber] = s

# Recopilar todas las nubes de revisión
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
        sheet_id = p.AsString() if p else None
        if sheet_id:
            sheet = sheet_by_number.get(sheet_id)
            if sheet:
                sheet_name = sheet.Title.replace("Sheet: ", "").strip()
                sheet_number = sheet.SheetNumber

    return sheet_name, sheet_number

# ============================================================
# Agrupar nubes por hoja -> vista
# ============================================================
data_by_sheet = {}
# Estructura:
# data_by_sheet[sheet_number] = {
#   'name': sheet_name,
#   'views': { view_name: [clouds...] }
# }

for revc in revision_clouds:
    if revc.RevisionId != selected_revision_id:
        continue

    sheet_name, sheet_number = get_sheet_info(revc)

    if " - " in sheet_name:
        sheet_name = sheet_name.split(" - ")[-1].strip()

    if sheet_number not in data_by_sheet:
        data_by_sheet[sheet_number] = {'name': sheet_name, 'views': {}}

    view = doc.GetElement(revc.OwnerViewId)
    view_name = view.Name if view else "Unknown View"

    if view_name not in data_by_sheet[sheet_number]['views']:
        data_by_sheet[sheet_number]['views'][view_name] = []

    data_by_sheet[sheet_number]['views'][view_name].append(revc)

# Ordenar hojas por número (natural)
sorted_sheets = sorted(data_by_sheet.items(), key=lambda x: natural_key(x[0]))

# ============================================================
# Reporte
# ============================================================
total_clouds = 0

for sheet_number, sheet_data in sorted_sheets:
    output.print_md('### **Sheet:** {} - {}\n'.format(sheet_number, sheet_data['name']))

    # Ordenar vistas dentro de la hoja
    view_items = list(sheet_data['views'].items())
    view_items_sorted = sorted(view_items, key=lambda x: natural_key(x[0]))

    for view_name, clouds in view_items_sorted:
        output.print_md('#### **View:** {}\n'.format(view_name))

        # Orden estable de clouds: por ElementId (simple y consistente)
        try:
            clouds.sort(key=lambda c: c.Id.IntegerValue)
        except:
            pass

        # --- Tabla por vista (CloudId | Sheet | View | Comment)
        rows = []
        headers = ["CloudId", "Sheet", "View", "Comment"]

        for revc in clouds:
            total_clouds += 1

            cloud_id_link = output.linkify([revc.Id])

            # Leer comentario de forma segura (independiente del idioma)
            comment_param = revc.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            comment = comment_param.AsString() if comment_param else ""
            if comment is None:
                comment = ""

            rows.append([cloud_id_link, sheet_number, view_name, comment])

        # Imprimir tabla solo si hay filas
        if rows:
            output.print_table(table_data=rows, columns=headers)
        else:
            output.print_md("_No revision clouds in this view._\n")

output.print_md('\n**SEARCH COMPLETED — {} revision clouds found.**'.format(total_clouds))
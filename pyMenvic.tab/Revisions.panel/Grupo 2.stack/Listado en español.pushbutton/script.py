# -*- coding: utf-8 -*-
from pyrevit import revit, DB, script, forms

# Recopilar todas las revisiones
revisions = DB.FilteredElementCollector(revit.doc)\
              .OfCategory(DB.BuiltInCategory.OST_Revisions)\
              .WhereElementIsNotElementType()\
              .ToElements()

# Crear una lista de información de revisiones para seleccionar
revision_options = []
revision_ids = []
for rev in revisions:
    rev_number = rev.LookupParameter('Revision Number').AsString()
    rev_date = rev.LookupParameter('Revision Date').AsString()
    rev_description = rev.LookupParameter('Revision Description').AsString()
    rev_info = "{}-{}-{}".format(rev_number, rev_description, rev_date)
    revision_options.append(rev_info)
    revision_ids.append(rev.Id)

# Mostrar un formulario para seleccionar un número de revisión
selected_revision_option = forms.SelectFromList.show(
    sorted(revision_options),
    title='Select Revision',
    button_name='Select',
    multiple=False
)

# Salir si no se seleccionó ninguna revisión
if not selected_revision_option:
    script.exit()

# Extraer el Id de la revisión seleccionada
selected_revision_index = revision_options.index(selected_revision_option)
selected_revision_id = revision_ids[selected_revision_index]

output = script.get_output()

# Recopilar todas las nubes de revisión
revision_clouds = DB.FilteredElementCollector(revit.doc)\
                    .OfCategory(DB.BuiltInCategory.OST_RevisionClouds)\
                    .WhereElementIsNotElementType()\
                    .ToElements()

# Función para obtener el nombre, número y nombre en español de la hoja donde se encuentra la nube de revisión
def get_sheet_info(revision_cloud):
    sheet_name = "Unknown Sheet"
    sheet_number = "Unknown Number"
    spanish_sheet = "Unknown Spanish Sheet"
    owner_view = revit.doc.GetElement(revision_cloud.OwnerViewId)
    if isinstance(owner_view, DB.ViewSheet):
        sheet_name = owner_view.Title.replace("Sheet: ", "").strip()
        sheet_number = owner_view.SheetNumber
        spanish_sheet_param = owner_view.LookupParameter('Spanish Sheet')
        if spanish_sheet_param:
            spanish_sheet = spanish_sheet_param.AsString()
    elif isinstance(owner_view, DB.View):
        sheet_id = owner_view.LookupParameter('Sheet Number').AsString()
        if sheet_id:
            sheets = DB.FilteredElementCollector(revit.doc)\
                      .OfCategory(DB.BuiltInCategory.OST_Sheets)\
                      .WhereElementIsNotElementType()\
                      .ToElements()
            sheet = next((s for s in sheets if s.SheetNumber == sheet_id), None)
            if sheet:
                sheet_name = sheet.Title.replace("Sheet: ", "").strip()
                sheet_number = sheet.SheetNumber
                spanish_sheet_param = sheet.LookupParameter('Spanish Sheet')
                if spanish_sheet_param:
                    spanish_sheet = spanish_sheet_param.AsString()
    return sheet_name, sheet_number, spanish_sheet

# Diccionario para agrupar nubes de revisión por hoja
revision_clouds_by_sheet = {}

# Iterar sobre cada nube de revisión
for revc in revision_clouds:
    # Filtrar por Id de revisión seleccionada
    if revc.RevisionId != selected_revision_id:
        continue
    
    # Obtener la información de la hoja
    sheet_name, sheet_number, spanish_sheet = get_sheet_info(revc)
    
    # Eliminar redundancias de "Sheet: " en el nombre de la hoja
    if " - " in sheet_name:
        sheet_name = sheet_name.split(" - ")[-1].strip()
    
    # Agregar la nube de revisión al grupo de hojas correspondiente
    if sheet_number not in revision_clouds_by_sheet:
        revision_clouds_by_sheet[sheet_number] = (sheet_name, spanish_sheet)

# Ordenar las hojas por número de hoja en orden ascendente
sorted_sheets = sorted(revision_clouds_by_sheet.items(), key=lambda x: x[0])

# Imprimir solo el listado de las hojas
output.print_md('### **Listado de Sheets**')
for sheet_number, (sheet_name, spanish_sheet) in sorted_sheets:
    output.print_md('{} - {} ({})'.format(sheet_number, sheet_name, spanish_sheet))

output.print_md('\n**SEARCH COMPLETED.**')

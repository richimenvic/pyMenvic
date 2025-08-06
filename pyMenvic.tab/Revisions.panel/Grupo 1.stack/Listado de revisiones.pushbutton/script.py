# -*- coding: utf-8 -*-
from pyrevit import revit, DB, script

output = script.get_output()

# Recopilar todas las nubes de revisión
revision_clouds = DB.FilteredElementCollector(revit.doc)\
                    .OfCategory(DB.BuiltInCategory.OST_RevisionClouds)\
                    .WhereElementIsNotElementType()\
                    .ToElements()

# Función para obtener el nombre y número de la hoja donde se encuentra la nube de revisión
def get_sheet_info(revision_cloud):
    sheet_name = "Unknown Sheet"
    sheet_number = "Unknown Number"
    owner_view = revit.doc.GetElement(revision_cloud.OwnerViewId)
    if isinstance(owner_view, DB.ViewSheet):
        sheet_name = owner_view.Title
        sheet_number = owner_view.SheetNumber
    elif isinstance(owner_view, DB.View):
        sheet_id = owner_view.LookupParameter('Sheet Number').AsString()
        if sheet_id:
            sheets = DB.FilteredElementCollector(revit.doc)\
                      .OfCategory(DB.BuiltInCategory.OST_Sheets)\
                      .WhereElementIsNotElementType()\
                      .ToElements()
            sheet = next((s for s in sheets if s.SheetNumber == sheet_id), None)
            if sheet:
                sheet_name = sheet.Title
                sheet_number = sheet.SheetNumber
    return sheet_name, sheet_number

# Diccionario para agrupar nubes de revisión por hoja
revision_clouds_by_sheet = {}

# Iterar sobre cada nube de revisión
for revc in revision_clouds:
    # Obtener la información de la hoja
    sheet_name, sheet_number = get_sheet_info(revc)
    
    # Eliminar redundancias de "Sheet: " en el nombre de la hoja
    if " - " in sheet_name:
        sheet_name = sheet_name.split(" - ")[-1].strip()
    
    # Agregar la nube de revisión al grupo de hojas correspondiente
    if sheet_number not in revision_clouds_by_sheet:
        revision_clouds_by_sheet[sheet_number] = {
            'name': sheet_name,
            'clouds': []
        }
    revision_clouds_by_sheet[sheet_number]['clouds'].append(revc)

# Ordenar las hojas por número de hoja en orden ascendente
sorted_sheets = sorted(revision_clouds_by_sheet.items())

# Imprimir detalles de las nubes de revisión agrupadas por hoja
for sheet_number, data in sorted_sheets:
    output.print_md('### **Sheet:** {} - {}\n'.format(sheet_number, data['name']))
    for revc in data['clouds']:
        # Obtener el nombre de la vista
        view = revit.doc.GetElement(revc.OwnerViewId)
        view_name = view.Name if view else "Unknown View"
        
        # Obtener el número de revisión
        revision = revit.doc.GetElement(revc.RevisionId)
        rev_number_param = revision.LookupParameter('Revision Number')
        rev_number = rev_number_param.AsString() if rev_number_param else "No Number"
        
        # Imprimir los detalles
        selectable_cloud_id = output.linkify([revc.Id])
        output.print_md('**Revision Cloud ID:** {}, **Revision Number:** {}, **View Name:** {}'.format(selectable_cloud_id, rev_number, view_name))

output.print_md('\n**SEARCH COMPLETED.**')

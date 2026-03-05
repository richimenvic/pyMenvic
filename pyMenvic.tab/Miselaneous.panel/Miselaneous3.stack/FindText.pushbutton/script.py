# -*- coding: utf-8 -*-
import Autodesk.Revit.DB as DB
from pyrevit import revit, script, forms

doc = revit.doc
output = script.get_output()

# 1. COLECTAR TODO (Incluye Leyendas y Vistas)
all_notes = DB.FilteredElementCollector(doc).OfClass(DB.TextNote).WhereElementIsNotElementType().ToElements()
all_types = DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).ToElements()

# 2. CONTAR NOTAS POR TIPO
conteo = {}
for n in all_notes:
    tid = n.GetTypeId().IntegerValue
    conteo[tid] = conteo.get(tid, 0) + 1

# 3. PREPARAR CLASE PARA EL MENÚ
class EstiloItem(forms.TemplateListItem):
    @property
    def name(self):
        # Muestra: [Cantidad] Nombre del Estilo
        c = conteo.get(self.item.Id.IntegerValue, 0)
        return "[{}] {}".format(c, DB.Element.Name.__get__(self.item))

# Filtrar solo tipos que tienen al menos 1 nota
tipos_usados = [t for t in all_types if conteo.get(t.Id.IntegerValue, 0) > 0]
tipos_usados.sort(key=lambda x: DB.Element.Name.__get__(x))

# 4. MOSTRAR MENÚ DE SELECCIÓN (Usando función universal)
seleccion = forms.SelectFromList.show(
    [EstiloItem(t) for t in tipos_usados],
    title="Selecciona Estilos para Localizar (Incluye Leyendas)",
    width=500,
    height=600,
    multiselect=True,
    button_name="Localizar Notas"
)

if seleccion:
    ids_buscados = [t.Id.IntegerValue for t in seleccion]
    report_data = []

    output.print_md("### 📍 Ubicación de Notas Seleccionadas")

    for note in all_notes:
        if note.GetTypeId().IntegerValue in ids_buscados:
            t_type = doc.GetElement(note.GetTypeId())
            t_name = DB.Element.Name.__get__(t_type)
            
            # Ubicación (Vista)
            view = doc.GetElement(note.OwnerViewId)
            view_name = view.Name if view else "Desconocida"
            
            # Buscar el Plano (Sheet)
            sheet_info = "---"
            if view:
                # Caso Leyendas
                if view.ViewType == DB.ViewType.Legend:
                    sheet_info = "LEYENDA"
                else:
                    # Buscar si la vista está en un plano
                    p_sheet = view.get_Parameter(DB.BuiltInParameter.VIEWER_SHEET_NUMBER)
                    if p_sheet and p_sheet.AsString():
                        sheet_info = p_sheet.AsString()

            # Contenido y Link
            text_preview = (note.Text[:40] + "..") if len(note.Text) > 40 else note.Text
            text_preview = text_preview.replace('\r', ' ').replace('\n', ' ')
            id_link = output.linkify(note.Id)

            report_data.append([t_name, text_preview, view_name, sheet_info, id_link])

    # Imprimir tabla filtrada
    output.print_table(
        report_data,
        columns=["Estilo", "Contenido", "Vista", "Plano (Sheet)", "ID (Clic)"]
    )

    # 5. OPCIÓN DE SELECCIÓN FÍSICA
    if forms.alert("¿Quieres seleccionar todas estas notas en Revit ahora mismo?", yesno=True):
        ids_to_select = [n.Id for n in all_notes if n.GetTypeId().IntegerValue in ids_buscados]
        revit.get_selection().set_elements(ids_to_select)
        print("\n[OK] Se han seleccionado {} notas.".format(len(ids_to_select)))
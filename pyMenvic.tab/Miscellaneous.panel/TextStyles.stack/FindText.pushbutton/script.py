# -*- coding: utf-8 -*-
__title__  = "Find Text"
__author__ = "Ricardo J. Mendieta"

"""
==========================================================
pyMENVIC | TEXT NOTE LOCATOR (Style Usage + Location)
Revit + pyRevit

Descripción
-----------
Herramienta para localizar y analizar el uso de estilos
de texto (TextNoteType) dentro del proyecto.

El script:
- cuenta todas las TextNotes del modelo
- muestra cuántas usan cada estilo de texto
- permite seleccionar uno o varios estilos
- genera un reporte con la ubicación de cada nota
- permite seleccionar las notas encontradas en Revit

Compatibilidad
---------------
- Revit 2020-2027+
- Evita ElementId.acceso antiguo directo

Autor
-----
Ricardo J. Mendieta
pyMENVIC – Ayudas para MENVIC ARQ
==========================================================
"""

import Autodesk.Revit.DB as DB
from pyrevit import revit, script, forms
from System.Collections.Generic import List


doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# ==========================================================
# HELPERS
# ==========================================================

def element_id_int(eid, default=-1):
    """Devuelve un int seguro para ElementId en Revit 2020-2027+."""
    if eid is None:
        return default
    try:
        return int(eid.Value)          # Revit 2024/2025/2026+
    except:
        pass
    try:
        return int(eid.IntegerValue)   # Revit antiguo
    except:
        pass
    try:
        return int(str(eid))
    except:
        return default


def get_element_name(el):
    if not el:
        return "<Desconocido>"
    try:
        return DB.Element.Name.__get__(el)
    except:
        pass
    try:
        return el.Name
    except:
        return "<Sin nombre>"


def unwrap_template_item(item):
    """SelectFromList puede devolver TemplateListItem o el item interno."""
    try:
        return item.item
    except:
        return item


def safe_note_text(note):
    try:
        txt = note.Text or ""
    except:
        txt = ""
    txt = txt.replace('\r', ' ').replace('\n', ' ')
    if len(txt) > 40:
        return txt[:40] + ".."
    return txt


def ask_select_notes(count):
    """Ask before printing the report, so no dialog covers the output window."""
    try:
        return forms.CommandSwitchWindow.show(
            ["REPORT ONLY", "REPORT + SELECT FOUND NOTES"],
            message="{} text notes found. Choose how to continue.".format(count),
            title="pyMenvic | Find Text"
        ) == "REPORT + SELECT FOUND NOTES"
    except:
        return False


def select_element_ids_in_revit(element_ids):
    """Selecciona elementos usando la API nativa de Revit."""
    try:
        ids_net = List[DB.ElementId]()
        for eid in element_ids:
            if isinstance(eid, DB.ElementId):
                ids_net.Add(eid)
        uidoc.Selection.SetElementIds(ids_net)
        return True, len(ids_net)
    except Exception as ex:
        return False, str(ex).splitlines()[0]


# ==========================================================
# 1. COLECTAR TODO (Incluye Leyendas y Vistas)
# ==========================================================

all_notes = DB.FilteredElementCollector(doc) \
    .OfClass(DB.TextNote) \
    .WhereElementIsNotElementType() \
    .ToElements()

all_types = DB.FilteredElementCollector(doc) \
    .OfClass(DB.TextNoteType) \
    .ToElements()


# ==========================================================
# 2. CONTAR NOTAS POR TIPO
# ==========================================================

conteo = {}
for n in all_notes:
    try:
        tid_int = element_id_int(n.GetTypeId())
        if tid_int != -1:
            conteo[tid_int] = conteo.get(tid_int, 0) + 1
    except:
        pass


# ==========================================================
# 3. PREPARAR CLASE PARA EL MENÚ
# ==========================================================

class EstiloItem(forms.TemplateListItem):
    @property
    def name(self):
        c = conteo.get(element_id_int(self.item.Id), 0)
        return "[{}] {}".format(c, get_element_name(self.item))


# Filtrar solo tipos que tienen al menos 1 nota
tipos_usados = []
for t in all_types:
    try:
        if conteo.get(element_id_int(t.Id), 0) > 0:
            tipos_usados.append(t)
    except:
        pass

tipos_usados.sort(key=lambda x: get_element_name(x))


# ==========================================================
# 4. MOSTRAR MENÚ DE SELECCIÓN
# ==========================================================

seleccion = forms.SelectFromList.show(
    [EstiloItem(t) for t in tipos_usados],
    title="Select Text Styles to Locate",
    width=500,
    height=600,
    multiselect=True,
    button_name="LOCATE NOTES"
)


if seleccion:
    selected_types = [unwrap_template_item(t) for t in seleccion]
    ids_buscados = set([element_id_int(t.Id) for t in selected_types])

    report_data = []
    ids_to_select = []

    for note in all_notes:
        try:
            note_type_id = note.GetTypeId()
            note_type_int = element_id_int(note_type_id)
        except:
            continue

        if note_type_int not in ids_buscados:
            continue

        t_type = doc.GetElement(note_type_id)
        t_name = get_element_name(t_type)

        # Ubicación (Vista)
        view = None
        try:
            view = doc.GetElement(note.OwnerViewId)
        except:
            view = None

        view_name = get_element_name(view) if view else "Unknown"

        # Buscar el Plano (Sheet)
        sheet_info = "---"
        if view:
            try:
                if view.ViewType == DB.ViewType.Legend:
                    sheet_info = "LEGEND"
                else:
                    p_sheet = view.get_Parameter(DB.BuiltInParameter.VIEWER_SHEET_NUMBER)
                    if p_sheet and p_sheet.AsString():
                        sheet_info = p_sheet.AsString()
            except:
                pass

        text_preview = safe_note_text(note)

        try:
            id_link = output.linkify(note.Id)
        except:
            id_link = str(element_id_int(note.Id))

        report_data.append([t_name, text_preview, view_name, sheet_info, id_link])
        ids_to_select.append(note.Id)

    # Ask selection mode before showing the report, so the report is never hidden by the popup.
    select_after_report = False
    if ids_to_select:
        select_after_report = ask_select_notes(len(ids_to_select))

    output.print_md("### 📍 Selected Text Notes Location")

    # Print filtered table
    if report_data:
        output.print_table(
            report_data,
            columns=["Style", "Content", "View", "Sheet", "ID (Click)"]
        )
    else:
        output.print_md("No text notes were found with the selected styles.")

    # ======================================================
    # 5. OPCIÓN DE SELECCIÓN FÍSICA
    # ======================================================

    if ids_to_select and select_after_report:
        ok, result = select_element_ids_in_revit(ids_to_select)
        if ok:
            print("\n[OK] {} text notes selected in Revit.".format(result))
        else:
            output.print_md("\n⚠️ Could not select the text notes in Revit: `{}`".format(result))

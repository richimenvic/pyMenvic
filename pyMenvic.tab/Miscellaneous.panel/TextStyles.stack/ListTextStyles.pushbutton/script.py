# -*- coding: utf-8 -*-
__title__  = "List text styles"
__author__ = "Ricardo J. Mendieta"

"""
==========================================================
pyMENVIC | TEXT STYLE LISTER (Usage Counter)
Revit + pyRevit

Descripción
-----------
Herramienta para listar todos los estilos de texto
(TextNoteType) del proyecto y mostrar cuántas veces
se utilizan en el modelo.

El script crea una lista de notas de texto en la vista
activa con el siguiente formato:

Style Name (N)

donde N representa el número de TextNotes que utilizan
ese estilo en todo el proyecto.

Funcionamiento
--------------
1. Recorre todas las TextNotes del proyecto.
2. Cuenta el uso de cada TextNoteType.
3. Genera una lista ordenada de estilos con su conteo.
4. Solicita al usuario un punto de inserción.
5. Inserta una nota de texto por cada estilo con
   el formato "Nombre del Estilo (Cantidad)".

Características
---------------
- Los estilos se ordenan alfabéticamente.
- La separación vertical se ajusta automáticamente
  según el tamaño del texto y la escala de la vista.
- Funciona como herramienta de auditoría de estilos
  de anotación en el proyecto.

Autor
-----
Ricardo J. Mendieta
pyMENVIC – Ayudas para MENVIC ARQ
==========================================================
"""
from pyrevit import revit, DB, forms
from rpw.ui.forms import (FlexForm, Label, ComboBox, Separator, Button)
from collections import OrderedDict
from Autodesk.Revit import Exceptions

# -----------------------------
# Revit 2026+ safe ElementId integer helper
# -----------------------------
def element_id_int(eid, default=-1):
    """Return a stable integer from ElementId across Revit versions."""
    if eid is None:
        return default
    try:
        return int(eid.IntegerValue)
    except:
        pass
    try:
        return int(eid.Value)
    except:
        pass
    try:
        return int(str(eid))
    except:
        return default

# -----------------------------
# Count usage of each TextNoteType in the whole project
# -----------------------------
use_count = {}  # typeId int -> count

try:
    txt_notes = DB.FilteredElementCollector(revit.doc) \
        .OfClass(DB.TextNote) \
        .WhereElementIsNotElementType() \
        .ToElements()

    for tn in txt_notes:
        try:
            tid = tn.GetTypeId()
            if tid:
                k = element_id_int(tid)
                use_count[k] = use_count.get(k, 0) + 1
        except:
            pass
except:
    # If something goes wrong, keep counts empty (all will show 0)
    use_count = {}

# -----------------------------
# Pick text style types
# -----------------------------
txt_types = DB.FilteredElementCollector(revit.doc).OfClass(DB.TextNoteType)

# Build dict {TextNoteType : "Name (Count)"}
text_style_dict = {}
for txt_t in txt_types:
    try:
        name = txt_t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if not name:
            name = "<Unnamed>"
    except:
        name = "<Unnamed>"

    c = use_count.get(element_id_int(txt_t.Id), 0)
    label = "{} ({})".format(name, c)

    text_style_dict[txt_t] = label

# Sort styles by label (which includes the count at end)
sorted_text_styles = OrderedDict(sorted(text_style_dict.items(), key=lambda t: t[1]))

view = revit.active_view

# dims and scale
scale = float(view.Scale) / 200.0
shift = 5 * scale
offset = 0

text_height = 0

with forms.WarningBar(title="Pick Point"):
    try:
        pick_point = revit.uidoc.Selection.PickPoint()
    except Exceptions.OperationCanceledException:
        forms.alert("Cancelled", ok=True, exitscript=True)

origin = pick_point
with revit.Transaction("Place Text Notes"):
    for ts in sorted_text_styles:
        label_text = sorted_text_styles[ts]

        text_position = DB.XYZ(pick_point.X, (pick_point.Y - offset), 0)

        try:
            text_height = ts.get_Parameter(DB.BuiltInParameter.TEXT_SIZE).AsDouble()
        except:
            text_height = 0.01  # fallback

        offset += (text_height * 2.75 * float(view.Scale))

        DB.TextNote.Create(revit.doc, view.Id, text_position, label_text, ts.Id)
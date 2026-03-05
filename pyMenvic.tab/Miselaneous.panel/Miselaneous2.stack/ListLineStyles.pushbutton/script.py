# -*- coding: utf-8 -*-
"""List Line Styles + Usage Count (Detail Lines) + Label Text using a default TextNoteType"""

from pyrevit import revit, DB, forms
from rpw.ui.forms import FlexForm, Label, ComboBox, Button
from Autodesk.Revit import Exceptions
from collections import OrderedDict
import sys


# -----------------------------
# Helpers
# -----------------------------
def mm_to_internal(mm_value):
    """Convert mm to Revit internal units (feet)."""
    return mm_value / 304.8


def normalize_name(s):
    """Normalize a type name to improve matching across locales and odd spaces."""
    if not s:
        return ""
    # Replace NBSP with normal space, trim, collapse multiple spaces
    s = s.replace(u"\u00A0", " ").strip()
    while "  " in s:
        s = s.replace("  ", " ")
    # Unify comma/dot for decimals
    s = s.replace(",", ".")
    return s.upper()


# -----------------------------
# Active view check
# -----------------------------
view = revit.active_view
if view.ViewType == DB.ViewType.Legend:
    forms.alert(
        "Este script no funciona en vistas de Leyenda.\n"
        "Úselo en Drafting Views, plantas, secciones o alzados.",
        ok=True,
        exitscript=True
    )


# -----------------------------
# Pick or Find text style (TextNoteType)
# -----------------------------
doc = revit.doc

txt_types = DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).ToElements()

# Key by the visible name parameter (what you see in Revit UI)
text_style_dict = {}
for txt_t in txt_types:
    try:
        n = txt_t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    except:
        n = None
    if n:
        text_style_dict[n] = txt_t

chosen_text_style = None

# 1) Auto-detect robust: any "Arial" containing "2.5" (matches 2.50, 2,50, etc.)
for name, txt_t in text_style_dict.items():
    if not name:
        continue
    nn = normalize_name(name)  # comma->dot, trims, fixes odd spaces
    if "ARIAL" in nn and "2.5" in nn:
        chosen_text_style = txt_t
        break

# 2) If still none, try some preferred defaults (fallback)
if chosen_text_style is None:
    target_candidates = [
        "Arial 2.5mm", "Arial 2,5mm", "Arial 2.5 mm", "Arial 2,5 mm",
        "Arial 2.50mm", "Arial 2,50mm", "Arial 2.50 mm", "Arial 2,50 mm",
    ]

    # Exact matches
    for name in target_candidates:
        if name in text_style_dict:
            chosen_text_style = text_style_dict[name]
            break

    # Normalized matches
    if chosen_text_style is None:
        norm_map = {normalize_name(k): v for k, v in text_style_dict.items()}
        for name in target_candidates:
            nn = normalize_name(name)
            if nn in norm_map:
                chosen_text_style = norm_map[nn]
                break

# 3) If still none, show a picker UI
if chosen_text_style is None:
    if not text_style_dict:
        forms.alert("No se encontraron tipos de texto (TextNoteType) en el documento.", ok=True, exitscript=True)

    components = [
        Label("No se encontró un Arial 2.5mm compatible.\nSeleccione un estilo de texto:"),
        ComboBox(name="textstyle_combobox", options=text_style_dict),
        Button("Seleccionar")
    ]
    form = FlexForm("Estilo de Texto", components)
    ok = form.show()
    if ok:
        chosen_text_style = form.values["textstyle_combobox"]
    else:
        sys.exit()


# -----------------------------
# Configuration (paper dimensions)
# -----------------------------
vert_offset_mm = 10.0  # vertical gap between rows (on paper)
line_length_mm = 40.0  # line length (on paper)
text_gap_mm = 2.0      # gap from end of line to text (on paper)


# -----------------------------
# Collect line style subcategories
# -----------------------------
cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
subcats = [subcat for subcat in cat.SubCategories]

unsorted_dict = {sc: sc.Name for sc in subcats}
sorted_subcats = OrderedDict(sorted(unsorted_dict.items(), key=lambda t: t[1]))


# -----------------------------
# Count usage (Detail Lines in entire project)
# -----------------------------
all_curve_elems = (
    DB.FilteredElementCollector(doc)
    .OfClass(DB.CurveElement)
    .WhereElementIsNotElementType()
    .ToElements()
)

line_usage = {}  # key: GraphicsStyleId (ElementId), value: count

for ce in all_curve_elems:
    try:
        if isinstance(ce, DB.DetailCurve):
            gs_id = ce.LineStyle.Id
            line_usage[gs_id] = line_usage.get(gs_id, 0) + 1
    except:
        pass


# -----------------------------
# Dimensions and scale logic
# -----------------------------
view_scale = float(view.Scale)
shift = mm_to_internal(vert_offset_mm * view_scale)
w = mm_to_internal(line_length_mm * view_scale)
text_offset = mm_to_internal(text_gap_mm * view_scale)


# -----------------------------
# Pick start point
# -----------------------------
with forms.WarningBar(title="Click en el punto de inicio"):
    try:
        pick_point = revit.uidoc.Selection.PickPoint()
    except Exceptions.OperationCanceledException:
        forms.alert("Cancelado", ok=True, exitscript=True)

p1 = pick_point
p2 = DB.XYZ(pick_point.X + w, pick_point.Y, 0)
base_line = DB.Line.CreateBound(p1, p2)


# -----------------------------
# Create line samples + labels
# -----------------------------
with revit.Transaction("Listar Estilos de Línea"):
    l1 = base_line
    first_row = True

    for ls in sorted_subcats.keys():
        if not first_row:
            # Move down for next row
            t_down = DB.Transform.CreateTranslation(DB.XYZ(0, -shift, 0))
            l1 = l1.CreateTransformed(t_down)
        else:
            first_row = False

        # Create detail line
        new_line = doc.Create.NewDetailCurve(view, l1)

        # Apply style
        gs = ls.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
        new_line.LineStyle = gs

        # Label text = "StyleName (count)"
        count = line_usage.get(gs.Id, 0)
        label_text = "{} ({})".format(sorted_subcats[ls], count)

        # Text position at end of line + offset
        text_pos = l1.GetEndPoint(1).Add(DB.XYZ(text_offset, 0, 0))

        # Create text note
        tn = DB.TextNote.Create(
            doc,
            view.Id,
            text_pos,
            label_text,
            chosen_text_style.Id
        )

        # Vertical alignment
        tn.VerticalAlignment = DB.VerticalTextAlignment.Middle
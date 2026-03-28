# -*- coding: utf-8 -*-
__title__  = "MENVIC | DETAIL ITEMS → STYLE PREVIEW"
__author__ = "Ricardo J. Mendieta"

"""
==========================================================
pyMENVIC | DETAIL ITEMS STYLE PREVIEW
Revit + pyRevit

Descripción
-----------
Herramienta para visualizar los estilos gráficos definidos
en la categoría Detail Items mediante muestras de línea
generadas en la vista activa.

El script crea una serie de Detail Lines con overrides
gráficos que replican las propiedades de cada subcategoría:

- Line Weight
- Line Color
- Line Pattern

Cada línea se acompaña de una etiqueta con el nombre
del estilo correspondiente.

Funcionamiento
--------------
1. Lee las subcategorías de Detail Items
   (OST_DetailComponents).
2. Extrae sus propiedades gráficas:
   peso, color y patrón.
3. Solicita un punto de inserción.
4. Genera una lista vertical de líneas de muestra.
5. Aplica overrides de vista para simular el estilo real.

Notas importantes
-----------------
- Las líneas creadas son Detail Lines con overrides
  gráficos de vista.
- No es posible convertirlas realmente en subcategorías
  de Detail Items.
- El script funciona tanto en proyectos como en familias.

Uso típico
----------
Auditoría gráfica y verificación rápida de estándares
de detalle en plantillas BIM.

Autor
-----
Ricardo J. Mendieta
pyMENVIC – Ayudas para MENVIC ARQ
==========================================================
"""
import Autodesk.Revit.DB as DB
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, script

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# ---------------- CONFIG ----------------
LINE_LEN = 8.0     # feet
ROW_GAP  = 1.2     # feet
TEXT_X   = 0.4     # feet after line end

# ---------------- HELPERS ----------------
def safe_int(v, default=0):
    try:
        return default if v is None else int(v)
    except:
        return default

def get_pattern_id(subcat, gtype):
    try:
        pid = subcat.GetLinePatternId(gtype)
        if pid and pid.IntegerValue != -1:
            return pid
    except:
        pass
    return DB.ElementId.InvalidElementId

def get_text_type():
    try:
        return DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).FirstElement()
    except:
        return None

def get_default_line_style():
    """Return a safe LineStyle (GraphicsStyle) for DetailCurves.
    This avoids showing 'Subcategory: Detail Items' which is misleading for curves.
    """
    try:
        line_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
        # prefer <Thin Lines> if present
        for s in list(line_cat.SubCategories):
            if s and (s.Name or "").strip().upper() == "<THIN LINES>":
                gs = s.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
                if gs:
                    return gs
        # otherwise just use the first available
        for s in list(line_cat.SubCategories):
            if s and "<" in (s.Name or ""):
                gs = s.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
                if gs:
                    return gs
    except:
        pass
    return None

def creator():
    try:
        return doc.FamilyCreate if doc.IsFamilyDocument else doc.Create
    except:
        return doc.Create

def draw_sample(view, p0, label, weight, color, pattern_id, text_type, default_line_style):
    p1 = DB.XYZ(p0.X + LINE_LEN, p0.Y, p0.Z)
    geom_line = DB.Line.CreateBound(p0, p1)
    crv = creator().NewDetailCurve(view, geom_line)

    # Ensure the element stays as a normal Detail Line style (Lines category)
    # and use overrides for visual preview.
    if default_line_style:
        try:
            crv.LineStyle = default_line_style
        except:
            pass

    ogs = DB.OverrideGraphicSettings()
    try:
        ogs.SetProjectionLineWeight(int(weight))
    except:
        pass
    try:
        ogs.SetProjectionLineColor(color)
    except:
        pass
    if pattern_id and pattern_id.IntegerValue != -1:
        try:
            ogs.SetProjectionLinePatternId(pattern_id)
        except:
            pass

    try:
        view.SetElementOverrides(crv.Id, ogs)
    except:
        pass

    if text_type:
        tp = DB.XYZ(p1.X + TEXT_X, p0.Y, p0.Z)
        try:
            DB.TextNote.Create(doc, view.Id, tp, label, text_type.Id)
        except:
            try:
                DB.TextNote.Create(doc, view.Id, tp, label)
            except:
                pass

    return crv

# ---------------- MAIN ----------------
view = doc.ActiveView

try:
    ins_pt = uidoc.Selection.PickPoint("Pick insertion point for Detail Items style preview")
except OperationCanceledException:
    script.exit()
except:
    ins_pt = DB.XYZ(0, 0, 0)

# We read styles from Detail Items (OST_DetailComponents) because that's your standard source.
try:
    detail_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_DetailComponents)
except:
    detail_cat = None

if not detail_cat:
    output.print_md("# MENVIC | DETAIL ITEMS → STYLE PREVIEW")
    output.print_md("**ERROR:** Could not access Detail Items category (OST_DetailComponents).")
    script.exit()

gtype = DB.GraphicsStyleType.Projection

rows = []
for sub in list(detail_cat.SubCategories):
    try:
        if not sub:
            continue
        w = safe_int(sub.GetLineWeight(gtype), 0) or 1
        c = sub.LineColor
        pid = get_pattern_id(sub, gtype)
        rows.append((w, sub.Name, c, pid))
    except:
        pass

rows.sort(key=lambda x: (x[0], (x[1] or "").upper()))

text_type = get_text_type()
default_ls = get_default_line_style()

created = 0
failed = 0

with DB.Transaction(doc, "MENVIC: Detail Items Style Preview") as t:
    t.Start()
    for i, (w, name, col, pid) in enumerate(rows):
        try:
            p = DB.XYZ(ins_pt.X, ins_pt.Y - (i * ROW_GAP), ins_pt.Z)
            ww = max(1, min(16, int(w)))
            draw_sample(view, p, name, ww, col, pid, text_type, default_ls)
            created += 1
        except:
            failed += 1
    t.Commit()

output.print_md("# MENVIC | DETAIL ITEMS → STYLE PREVIEW")
output.print_md("View: **{0}**".format(view.Name))
output.print_md("Document: **{0}**".format("FAMILY" if doc.IsFamilyDocument else "PROJECT"))
output.print_md("---")
output.print_md("SAMPLES CREATED: **{0}**".format(created))
output.print_md("FAILED ROWS: **{0}**".format(failed))
output.print_md("---")
output.print_md("NOTE: These are Detail Lines with view overrides (for accurate visual preview). They cannot truly become Detail Items subcategories.")
output.print_md("END")

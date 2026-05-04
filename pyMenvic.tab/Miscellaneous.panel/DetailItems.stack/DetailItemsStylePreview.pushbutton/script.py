# -*- coding: utf-8 -*-
__title__  = "MENVIC | DETAIL ITEMS → STYLE PREVIEW"
__author__ = "Ricardo J. Mendieta"

"""
==========================================================
pyMENVIC | DETAIL ITEMS STYLE PREVIEW
Revit + pyRevit

Creates a visual preview of Detail Items subcategories.

IMPORTANT
---------
Inside a Detail Item family, the preview lines are assigned to the
actual Detail Items subcategories, so the Properties palette will show
the real subcategory instead of <Thin Lines>.

Inside a project, Revit Detail Lines cannot truly become Detail Items
subcategories. In that case, the tool uses Detail Lines with view
overrides as a visual preview only.
==========================================================
"""

import Autodesk.Revit.DB as DB
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, script, forms

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# ---------------- CONFIG ----------------
LINE_LEN = 8.0     # feet
ROW_GAP  = 1.2     # feet
TEXT_X   = 0.4     # feet after line end


# ---------------- HELPERS ----------------
def element_id_int(eid, default=-1):
    """Safe ElementId integer for Revit 2020-2027."""
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


def is_valid_element_id(eid):
    try:
        if eid == DB.ElementId.InvalidElementId:
            return False
    except:
        pass
    return element_id_int(eid, -1) != -1


def safe_int(v, default=0):
    try:
        return default if v is None else int(v)
    except:
        return default


def get_pattern_id(subcat, gtype):
    try:
        pid = subcat.GetLinePatternId(gtype)
        if is_valid_element_id(pid):
            return pid
    except:
        pass
    return DB.ElementId.InvalidElementId


def get_text_type():
    try:
        return DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).FirstElement()
    except:
        return None


def get_fallback_line_style():
    """Safe fallback LineStyle for project mode or failed family assignment."""
    try:
        line_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)

        # Prefer <Thin Lines> only as fallback, not as the intended final style.
        for s in list(line_cat.SubCategories):
            if s and (s.Name or "").strip().upper() == "<THIN LINES>":
                gs = s.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
                if gs:
                    return gs

        for s in list(line_cat.SubCategories):
            if s:
                gs = s.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
                if gs:
                    return gs
    except:
        pass
    return None


def get_subcategory_graphics_style(subcat):
    """GraphicsStyle of the Detail Items subcategory."""
    try:
        gs = subcat.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
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


def can_create_detail_curve_in_view(view):
    if not view:
        return False
    try:
        if view.IsTemplate:
            return False
    except:
        pass
    try:
        bad_types = [
            DB.ViewType.ThreeD,
            DB.ViewType.Schedule,
            DB.ViewType.DrawingSheet,
            DB.ViewType.ProjectBrowser,
            DB.ViewType.SystemBrowser
        ]
        if view.ViewType in bad_types:
            return False
    except:
        pass
    return True


def apply_visual_overrides(view, element_id, weight, color, pattern_id):
    """Visual backup. Useful in project mode and harmless in most family views."""
    ogs = DB.OverrideGraphicSettings()

    try:
        ogs.SetProjectionLineWeight(int(weight))
    except:
        pass

    try:
        ogs.SetProjectionLineColor(color)
    except:
        pass

    if is_valid_element_id(pattern_id):
        try:
            ogs.SetProjectionLinePatternId(pattern_id)
        except:
            pass

    try:
        view.SetElementOverrides(element_id, ogs)
        return True
    except:
        return False


def draw_sample(view, p0, label, weight, color, pattern_id, text_type, subcat_style, fallback_line_style):
    p1 = DB.XYZ(p0.X + LINE_LEN, p0.Y, p0.Z)
    geom_line = DB.Line.CreateBound(p0, p1)
    crv = creator().NewDetailCurve(view, geom_line)

    used_real_subcategory = False

    # FAMILY MODE:
    # In a Detail Item family, Detail Curves can use the family subcategories.
    # This makes the Properties palette show the correct subcategory name.
    if doc.IsFamilyDocument and subcat_style:
        try:
            crv.LineStyle = subcat_style
            used_real_subcategory = True
        except:
            used_real_subcategory = False

    # PROJECT MODE / FALLBACK:
    # Project Detail Lines cannot truly be Detail Items subcategories.
    # Keep a safe line style and use overrides for preview.
    if not used_real_subcategory and fallback_line_style:
        try:
            crv.LineStyle = fallback_line_style
        except:
            pass

    # Keep overrides as visual backup. In family mode, the real subcategory is still assigned.
    apply_visual_overrides(view, crv.Id, weight, color, pattern_id)

    if text_type:
        tp = DB.XYZ(p1.X + TEXT_X, p0.Y, p0.Z)
        try:
            DB.TextNote.Create(doc, view.Id, tp, label, text_type.Id)
        except:
            try:
                DB.TextNote.Create(doc, view.Id, tp, label)
            except:
                pass

    return crv, used_real_subcategory


# ---------------- MAIN ----------------
view = doc.ActiveView

if not can_create_detail_curve_in_view(view):
    forms.alert(
        "This tool needs an active 2D/detail view that supports Detail Lines.\n\n"
        "Open a plan/detail/drafting view or the main view inside a Detail Item family, then run it again.",
        title="MENVIC | Detail Items Style Preview",
        exitscript=True
    )

try:
    ins_pt = uidoc.Selection.PickPoint("Pick insertion point for Detail Items style preview")
except OperationCanceledException:
    script.exit()
except:
    ins_pt = DB.XYZ(0, 0, 0)

try:
    detail_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_DetailComponents)
except:
    detail_cat = None

if not detail_cat:
    output.print_md("# MENVIC | DETAIL ITEMS → STYLE PREVIEW")
    output.print_md("**ERROR:** Could not access Detail Items category (OST_DetailComponents).")
    output.print_md("If this is a family, confirm it is a Detail Item family, not a Generic Model or annotation family.")
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
        gs = get_subcategory_graphics_style(sub)

        rows.append((w, sub.Name, c, pid, gs))
    except:
        pass

rows.sort(key=lambda x: (x[0], (x[1] or "").upper()))

if not rows:
    forms.alert(
        "No Detail Items subcategories were found in this document.",
        title="MENVIC | Detail Items Style Preview",
        exitscript=True
    )

text_type = get_text_type()
fallback_ls = get_fallback_line_style()

created = 0
failed = 0
real_subcat_count = 0
override_only_count = 0

with DB.Transaction(doc, "MENVIC: Detail Items Style Preview") as t:
    t.Start()

    for i, (w, name, col, pid, gs) in enumerate(rows):
        try:
            p = DB.XYZ(ins_pt.X, ins_pt.Y - (i * ROW_GAP), ins_pt.Z)
            ww = max(1, min(16, int(w)))

            crv, used_real = draw_sample(
                view,
                p,
                name,
                ww,
                col,
                pid,
                text_type,
                gs,
                fallback_ls
            )

            created += 1
            if used_real:
                real_subcat_count += 1
            else:
                override_only_count += 1

        except:
            failed += 1

    t.Commit()

output.print_md("# MENVIC | DETAIL ITEMS → STYLE PREVIEW")
output.print_md("View: **{0}**".format(view.Name))
output.print_md("Document: **{0}**".format("FAMILY" if doc.IsFamilyDocument else "PROJECT"))
output.print_md("---")
output.print_md("SAMPLES CREATED: **{0}**".format(created))
output.print_md("REAL DETAIL ITEMS SUBCATEGORY LINES: **{0}**".format(real_subcat_count))
output.print_md("OVERRIDE-ONLY PREVIEW LINES: **{0}**".format(override_only_count))
output.print_md("FAILED ROWS: **{0}**".format(failed))
output.print_md("---")

if doc.IsFamilyDocument:
    output.print_md("NOTE: In a Detail Item family, the generated lines should show the actual Detail Items subcategory in Properties.")
else:
    output.print_md("NOTE: In projects, Revit Detail Lines cannot truly become Detail Items subcategories. The tool uses view overrides for visual preview.")

output.print_md("TIP: If all line weights look identical on screen, check whether Revit's **Thin Lines** display mode is enabled.")
output.print_md("END")

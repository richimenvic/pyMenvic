# -*- coding: utf-8 -*-
"""List Line Styles + Usage Count (Detail Lines)"""

from pyrevit import revit, DB, forms, HOST_APP
from rpw.ui.forms import FlexForm, Label, ComboBox, Separator, Button, TextBox
from collections import OrderedDict
from Autodesk.Revit import Exceptions
import sys


def convert_length_to_internal(d_units):
    units = revit.doc.GetUnits()
    if HOST_APP.is_newer_than(2021):
        internal_units = units.GetFormatOptions(DB.SpecTypeId.Length).GetUnitTypeId()
    else:
        internal_units = units.GetFormatOptions(DB.UnitType.UT_Length).DisplayUnits
    return DB.UnitUtils.ConvertToInternalUnits(d_units, internal_units)


# --- Active view check ---
view = revit.active_view
if view.ViewType == DB.ViewType.Legend:
    forms.alert(
        "This script does not work in Legend views.\n"
        "Use it in a Drafting View, plan, section or elevation.",
        ok=True,
        exitscript=True
    )


# --- Pick text style ---
txt_types = DB.FilteredElementCollector(revit.doc).OfClass(DB.TextNoteType)
text_style_dict = {
    txt_t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString(): txt_t
    for txt_t in txt_types
}

components = [
    Label("Pick Text Style:"),
    ComboBox(name="textstyle_combobox", options=text_style_dict),
    Separator(),
    Label("Vertical Offset (mm):"),
    TextBox(name="offset", Text="500"),
    Button("Select")
]

form = FlexForm("Appearance", components)
ok = form.show()

if ok:
    chosen_text_style = form.values["textstyle_combobox"]
    vert_offset = float(form.values["offset"])
else:
    sys.exit()


# --- Collect line style subcategories ---
cat = revit.doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
subcats = [subcat for subcat in cat.SubCategories]

unsorted_dict = {sc: sc.Name for sc in subcats}
sorted_subcats = OrderedDict(sorted(unsorted_dict.items(), key=lambda t: t[1]))


# --- Count usage (Detail Lines in entire project) ---
# Use CurveElement (supported), then filter to DetailCurve instances
all_curve_elems = (
    DB.FilteredElementCollector(revit.doc)
    .OfClass(DB.CurveElement)
    .WhereElementIsNotElementType()
    .ToElements()
)

line_usage = {}  # key: GraphicsStyleId (ElementId), value: count

for ce in all_curve_elems:
    try:
        # Keep only detail curves (detail lines)
        if isinstance(ce, DB.DetailCurve):
            gs_id = ce.LineStyle.Id
            line_usage[gs_id] = line_usage.get(gs_id, 0) + 1
    except:
        pass


# --- Dimensions and scale ---
scale = float(view.Scale) / 100.0
w = 20 * scale
text_offset = 1 * scale
shift = convert_length_to_internal(vert_offset) * scale


# --- Pick start point ---
with forms.WarningBar(title="Pick Point"):
    try:
        pick_point = revit.uidoc.Selection.PickPoint()
    except Exceptions.OperationCanceledException:
        forms.alert("Cancelled", ok=True, exitscript=True)

p1 = pick_point
p2 = DB.XYZ(pick_point.X + w, pick_point.Y, 0)

base_line = DB.Line.CreateBound(p1, p2)


with revit.Transaction("Draw Lines"):
    l1 = base_line
    for ls in sorted_subcats.keys():

        # Move down
        t1 = DB.Transform.CreateTranslation(DB.XYZ(0, -shift, 0))
        l1 = l1.CreateTransformed(t1)

        # Create detail line
        new_line = revit.doc.Create.NewDetailCurve(view, l1)

        # Apply style
        gs = ls.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
        new_line.LineStyle = gs

        # Label text
        count = line_usage.get(gs.Id, 0)
        label_text = "{} ({})".format(sorted_subcats[ls], count)

        # Place text to the right of the line end
        t2 = DB.Transform.CreateTranslation(DB.XYZ(text_offset, 0, 0))
        text_position = l1.CreateTransformed(t2).GetEndPoint(1)

        DB.TextNote.Create(
            revit.doc,
            view.Id,
            text_position,
            label_text,
            chosen_text_style.Id
        )

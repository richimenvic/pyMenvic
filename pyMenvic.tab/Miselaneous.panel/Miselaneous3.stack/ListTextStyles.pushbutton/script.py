# -*- coding: utf-8 -*-
"""List Text Styles and show usage count in project: 'Style Name (N)' """

from pyrevit import revit, DB, forms
from rpw.ui.forms import (FlexForm, Label, ComboBox, Separator, Button)
from collections import OrderedDict
from Autodesk.Revit import Exceptions

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
                k = tid.IntegerValue
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

    c = use_count.get(txt_t.Id.IntegerValue, 0)
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
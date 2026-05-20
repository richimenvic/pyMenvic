__title__ = "Purge imported dwg line patterns"
__doc__ = "Delete all the pesky imported dwg line patterns from the Project."

from pyrevit import forms
from pyrevit import revit, DB
from pyrevit import script
from pyrevit.framework import List

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

cl = DB.FilteredElementCollector(doc)

CAD_IMPORT_TOKENS = [
    ".DWG",
    ".DXF",
    "DWG",
    "DXF",
    "XREF",
    "DEFPOINTS",
    "$0$",
    "|",
    "IMPORTED",
    "IMPORT",
]

def element_id_int(element_id):
    try:
        return int(element_id.Value)
    except:
        try:
            return int(element_id.IntegerValue)
        except:
            return int(element_id)

def matched_token(name):
    name_upper = (name or "").upper()
    for token in CAD_IMPORT_TOKENS:
        if token in name_upper:
            return token
    return None

pattern_rows = []
pat_imports = List[DB.ElementId]()

for pat in cl.OfClass(DB.LinePatternElement):
    token = matched_token(pat.Name)
    if token:
        pat_imports.Add(pat.Id)
        pattern_rows.append([pat.Name, element_id_int(pat.Id), token])

l_num = str(len(pat_imports))
message = 'There are {} imported line patterns in the model. Only the listed IDs will be deleted. Are you sure you want to delete them?'.format(l_num)

if len(pat_imports) == 0:
    forms.alert("No Imported Line Patterns, well done!")
else:
    output.print_md("## Imported Line Pattern Delete Candidates")
    output.print_table(
        table_data=pattern_rows,
        columns=["Line Pattern Name", "ElementId", "Matched Token"],
    )

    if forms.alert(message, ok=False, yes=True, no=True, exitscript=True):
        with revit.Transaction("Delete imported patterns"):
            doc.Delete(pat_imports)
            forms.alert("{} line patterns deleted.".format(l_num), warn_icon=False)





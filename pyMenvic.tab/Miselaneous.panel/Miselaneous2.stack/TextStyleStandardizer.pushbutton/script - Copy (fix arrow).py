# -*- coding: utf-8 -*-
__title__ = "FIX | Copiar Arrowhead desde un Tipo BUENO"
__author__ = "pyMenvic"

import Autodesk.Revit.DB as DB
from pyrevit import revit, script, forms
import re

from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

def norm(s):
    return re.sub(r"\s+", " ", (s or "")).strip()

class TextNoteFilter(ISelectionFilter):
    def AllowElement(self, e):
        try:
            return isinstance(e, DB.TextNote)
        except:
            return False
    def AllowReference(self, r, p):
        return False

def get_arrow_eid_from_texttype(tt):
    try:
        p = tt.get_Parameter(DB.BuiltInParameter.LEADER_ARROWHEAD)
        if not p or p.StorageType != DB.StorageType.ElementId:
            return None, "missing-or-not-elementid"
        return p.AsElementId(), None
    except Exception as ex:
        return None, "read-failed: {}".format(ex)

def arrow_is_resolvable(eid):
    """En tu caso: 'resolvable' = existe y tiene Category o tiene TypeName string."""
    try:
        if not eid or eid == DB.ElementId.InvalidElementId:
            return False
        if eid.IntegerValue < 0:
            return False
        e = doc.GetElement(eid)
        if not e:
            return False
        # si tiene categoría, casi seguro aparece en UI
        if getattr(e, "Category", None) is not None:
            return True
        # o si tiene Type Name real
        try:
            ptn = e.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if ptn and ptn.StorageType == DB.StorageType.String and norm(ptn.AsString() or ""):
                return True
        except:
            pass
        return False
    except:
        return False

def set_arrow_on_texttype(tt, target_eid):
    try:
        p = tt.get_Parameter(DB.BuiltInParameter.LEADER_ARROWHEAD)
        if not p:
            return False, "missing-param"
        if p.IsReadOnly:
            return False, "readonly"
        if p.StorageType != DB.StorageType.ElementId:
            return False, "not-elementid"
        p.Set(target_eid)
        after = p.AsElementId()
        if after and after.IntegerValue == target_eid.IntegerValue:
            return True, None
        return False, "set-did-not-stick"
    except Exception as ex:
        return False, "set-failed: {}".format(ex)

# ------------------------------------------------------------
# 1) Seleccionar un TextNote que use el tipo "bueno"
# ------------------------------------------------------------
output.print_md("# FIX | Copiar Arrowhead desde un Tipo BUENO")
output.print_md("Selecciona un **TextNote** que use el Text Style que ya corregiste (donde se ve 'Arrow 30 Degree').")

try:
    ref = uidoc.Selection.PickObject(ObjectType.Element, TextNoteFilter(), "Selecciona un TextNote (texto) con el estilo BUENO")
except:
    forms.alert("Selección cancelada.", exitscript=True)

note = doc.GetElement(ref.ElementId)
tt = doc.GetElement(note.GetTypeId())

good_type_name = norm(getattr(tt, "Name", "") or "")
good_type_id = tt.Id.IntegerValue

good_arrow_eid, err = get_arrow_eid_from_texttype(tt)
if err or not good_arrow_eid:
    forms.alert("No pude leer LEADER_ARROWHEAD del tipo seleccionado: {}".format(err), exitscript=True)

# Mostrar info del “arrow” seleccionado
arrow_elem = doc.GetElement(good_arrow_eid)
arrow_class = ""
arrow_cat = ""
arrow_name = ""
try:
    arrow_class = arrow_elem.GetType().ToString() if arrow_elem else "NotFound"
except:
    arrow_class = "?"
try:
    arrow_cat = norm(arrow_elem.Category.Name) if (arrow_elem and arrow_elem.Category) else "None"
except:
    arrow_cat = "?"
try:
    arrow_name = norm(getattr(arrow_elem, "Name", "") or "")
except:
    arrow_name = ""

output.print_md("## Origen (tipo BUENO)")
output.print_md("- TextNoteType: **{}** (Id:{})".format(good_type_name if good_type_name else "∅", good_type_id))
output.print_md("- LEADER_ARROWHEAD Id: **{}**".format(good_arrow_eid.IntegerValue))
output.print_md("- Arrow elem: Name=`{}` | Class=`{}` | Category=`{}`".format(arrow_name if arrow_name else "∅", arrow_class, arrow_cat))

# ------------------------------------------------------------
# 2) Elegir modo
# ------------------------------------------------------------
mode = forms.CommandSwitchWindow.show(
    ["APLICAR SOLO A TIPOS MALOS", "FORZAR A TODOS", "CANCEL"],
    message="¿Cómo quieres aplicar el Arrowhead del tipo BUENO al resto?",
    title="pyMenvic"
)
if not mode or mode == "CANCEL":
    script.exit()

apply_all = (mode == "FORZAR A TODOS")

# ------------------------------------------------------------
# 3) Aplicar a todos los TextNoteType
# ------------------------------------------------------------
types = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))

changed = 0
skipped = 0
failed = []
already = 0

with revit.Transaction("pyMenvic | Copy Leader Arrowhead From Good Type"):
    for t in sorted(types, key=lambda x: x.Id.IntegerValue):
        tname = norm(getattr(t, "Name", "") or "")
        cur_eid, cur_err = get_arrow_eid_from_texttype(t)

        if cur_err:
            failed.append((t.Id.IntegerValue, tname, "read", cur_err))
            continue

        if cur_eid and cur_eid.IntegerValue == good_arrow_eid.IntegerValue:
            already += 1
            # si es "solo malos", lo dejamos igual
            if not apply_all:
                skipped += 1
                continue

        needs = apply_all or (not arrow_is_resolvable(cur_eid))
        if not needs:
            skipped += 1
            continue

        ok, reason = set_arrow_on_texttype(t, good_arrow_eid)
        if ok:
            changed += 1
        else:
            failed.append((t.Id.IntegerValue, tname, "set", reason))

output.print_md("---")
output.print_md("## Resultado")
output.print_md("- Aplicación: **{}**".format("FORZAR A TODOS" if apply_all else "SOLO TIPOS MALOS"))
output.print_md("- Ya tenían ese ArrowId: **{}**".format(already))
output.print_md("- Cambiados: **{}**".format(changed))
output.print_md("- Saltados: **{}**".format(skipped))
output.print_md("- Fallos: **{}**".format(len(failed)))

if failed:
    output.print_md("## Fallos (muestra)")
    output.print_md("| TypeId | TypeName | Etapa | Motivo |")
    output.print_md("|---:|---|---|---|")
    for tid, tn, stage, reason in failed[:200]:
        output.print_md("| {} | {} | {} | {} |".format(tid, tn if tn else "∅", stage, reason))

output.print_md("_Fin._")
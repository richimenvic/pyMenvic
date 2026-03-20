# -*- coding: utf-8 -*-
"""MENVIC | FAMILY DETAIL ITEMS - SEED STANDARD

Family Editor tool to standardize Object Styles -> Detail Items subcategories.

STANDARD (Projection):
- Creates/updates 8 canonical subcategories under Detail Items:
  1 - Extra Fine Line
  2 - Fine Line
  3 - Medium Line
  4 - Medium Heavy Line
  5 - Heavy Line
  6 - Very Heavy Line
  7 - Extra Heavy Line
  8 - Ultra Heavy Line
- Forces for canonicals:
  - Line Weight (Projection): 1..8
  - Line Color: Black
  - Line Pattern: Solid

RUN mode also:
- Reassigns CurveElements using any non-canonical Detail Items subcategory
  (excluding <...>) to the canonical subcategory matching its Projection weight.
- Purges (deletes) those old subcategories where possible.

Notes:
- Revit API does not reliably allow renaming subcategories. This script achieves
  the practical result by reassigning elements to canonical subcategories and
  deleting unused ones.

Author: Ricardo J. Mendieta (pyMenvic)
"""

import Autodesk.Revit.DB as DB
from pyrevit import revit, script, forms


doc = revit.doc
output = script.get_output()

# ---------------- CONFIG ----------------
WEIGHT_TO_NAME = {
    1: "1 - Extra Fine Line",
    2: "2 - Fine Line",
    3: "3 - Medium Line",
    4: "4 - Medium Heavy Line",
    5: "5 - Heavy Line",
    6: "6 - Very Heavy Line",
    7: "7 - Extra Heavy Line",
    8: "8 - Ultra Heavy Line",
}

BLACK = DB.Color(0, 0, 0)
GTYPE = DB.GraphicsStyleType.Projection


# ---------------- UTIL ----------------
def safe_int(v, default=0):
    try:
        return default if v is None else int(v)
    except:
        return default


def get_solid_pattern_id():
    """Best-effort solid pattern. If not found, InvalidElementId is treated as solid in many contexts."""
    try:
        pats = DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement)
        for p in pats:
            n = (p.Name or "").strip().upper()
            if n == "SOLID":
                return p.Id
    except:
        pass
    return DB.ElementId.InvalidElementId


def ensure_subcategory(parent_cat, name):
    try:
        if parent_cat.SubCategories.Contains(name):
            return parent_cat.SubCategories.get_Item(name)
        return doc.Settings.Categories.NewSubcategory(parent_cat, name)
    except:
        return None


def apply_props(subcat, weight, color_obj, pattern_id, gtype=GTYPE):
    # Color
    try:
        subcat.LineColor = color_obj
    except:
        pass

    # Weight
    try:
        subcat.SetLineWeight(int(weight), gtype)
    except:
        pass

    # Pattern
    try:
        # Setting InvalidElementId often equals solid
        subcat.SetLinePatternId(pattern_id, gtype)
    except:
        pass


def get_subcat_gsid(subcat, gtype=GTYPE):
    try:
        gs = subcat.GetGraphicsStyle(gtype)
        if not gs:
            return None
        return gs.Id.IntegerValue
    except:
        return None


def iter_user_subcats(parent_cat):
    """Iterate Detail Items subcats excluding built-ins like <Hidden Lines>."""
    for sub in list(parent_cat.SubCategories):
        if not sub:
            continue
        try:
            nm = sub.Name or ""
            if "<" in nm:
                continue
        except:
            continue
        yield sub


def count_curve_usage():
    """Counts CurveElement.LineStyle usage by GraphicsStyleId (int)."""
    usage = {}
    curves = DB.FilteredElementCollector(doc).OfClass(DB.CurveElement).WhereElementIsNotElementType()
    for cv in curves:
        try:
            ls = cv.LineStyle
            if ls:
                sid = ls.Id.IntegerValue
                usage[sid] = usage.get(sid, 0) + 1
        except:
            pass
    return usage


def seed_canonicals(parent_cat, solid_id):
    """Create/update canonical subcategories and return (created, updated, protected_names)."""
    created = 0
    updated = 0
    protected = set()

    for w in range(1, 9):
        name = WEIGHT_TO_NAME[w]
        protected.add(name)

        existed = False
        try:
            existed = parent_cat.SubCategories.Contains(name)
        except:
            existed = False

        sub = ensure_subcategory(parent_cat, name)
        if not sub:
            continue

        apply_props(sub, w, BLACK, solid_id, GTYPE)
        if existed:
            updated += 1
        else:
            created += 1

    return created, updated, protected


# ---------------- GUARDS ----------------
if not doc.IsFamilyDocument:
    forms.alert(
        "This tool is for Family Editor only (doc.IsFamilyDocument == True).",
        title="MENVIC | FAMILY DETAIL ITEMS STANDARD",
    )
    script.exit()

try:
    parent_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_DetailComponents)
except:
    parent_cat = None

if not parent_cat:
    forms.alert(
        "Cannot access 'Detail Items' category (OST_DetailComponents).",
        title="MENVIC | FAMILY DETAIL ITEMS STANDARD",
    )
    script.exit()


# ---------------- UI ----------------
options = [
    "RUN: SEED + CLEAN",
    "SEED STANDARD",
    "CANCEL",
]

choice = forms.CommandSwitchWindow.show(
    options,
    title="MENVIC | FAMILY DETAIL ITEMS",
    message="OBJECT STYLES → DETAIL ITEMS  |  Standard 1..8 (Black / Solid)",
)

if choice is None or choice == "CANCEL":
    script.exit()

SEED_ONLY = (choice == "SEED STANDARD")


# ---------------- EXEC ----------------
solid_id = get_solid_pattern_id()

# Always seed first
created = 0
updated = 0
mapped = 0
reassigned = 0
purge_deleted_ok = 0
purge_deleted_fail = 0

usage_before = count_curve_usage()

output.print_md("# MENVIC | FAMILY DETAIL ITEMS STANDARD")
output.print_md("---")

with DB.Transaction(doc, "MENVIC: Seed Detail Items Standard") as t:
    t.Start()
    c, u, protected_names = seed_canonicals(parent_cat, solid_id)
    created += c
    updated += u
    t.Commit()

output.print_md("## SEED SUMMARY")
output.print_md("CANONICALS CREATED: {0}".format(created))
output.print_md("CANONICALS UPDATED: {0}".format(updated))

if SEED_ONLY:
    output.print_md("---")
    output.print_md("MODE: SEED ONLY (NO CLEAN)")
    output.print_md("END")
    script.exit()

# Build canonical GS lookup
canon_name_to_gs = {}
for w in range(1, 9):
    nm = WEIGHT_TO_NAME[w]
    try:
        sub = parent_cat.SubCategories.get_Item(nm)
        gs = sub.GetGraphicsStyle(GTYPE) if sub else None
        if gs:
            canon_name_to_gs[nm] = gs
    except:
        pass

# Map old styles -> canonical GS by weight
old_gsid_to_new_gs = {}
subcats_to_try_delete = []
canon_names = set(WEIGHT_TO_NAME.values())

for sub in iter_user_subcats(parent_cat):
    try:
        nm = sub.Name or ""
        if nm in canon_names:
            # already canonical
            continue

        w = safe_int(sub.GetLineWeight(GTYPE), 0)
        if w not in WEIGHT_TO_NAME:
            continue

        target_name = WEIGHT_TO_NAME[w]
        canon_gs = canon_name_to_gs.get(target_name)
        if not canon_gs:
            continue

        old_gsid = get_subcat_gsid(sub, GTYPE)
        if old_gsid is None:
            continue

        old_gsid_to_new_gs[old_gsid] = canon_gs
        subcats_to_try_delete.append(sub.Id)
        mapped += 1

        # Also normalize props of the old one (optional, keeps things consistent if delete fails)
        apply_props(sub, w, BLACK, solid_id, GTYPE)
    except:
        pass

# Reassign CurveElements
with DB.Transaction(doc, "MENVIC: Reassign to Canonical Detail Items") as t2:
    t2.Start()
    curves = DB.FilteredElementCollector(doc).OfClass(DB.CurveElement).WhereElementIsNotElementType()
    for cv in curves:
        try:
            ls = cv.LineStyle
            if not ls:
                continue
            old_id = ls.Id.IntegerValue
            if old_id in old_gsid_to_new_gs:
                cv.LineStyle = old_gsid_to_new_gs[old_id]
                reassigned += 1
        except:
            pass
    t2.Commit()

# Purge: delete mapped subcats + delete any remaining unused (and not protected)
usage_after = count_curve_usage()

with DB.Transaction(doc, "MENVIC: Purge Detail Items Subcategories") as t3:
    t3.Start()

    # Delete the specific mapped subcategories first
    for sid in subcats_to_try_delete:
        try:
            doc.Delete(sid)
            purge_deleted_ok += 1
        except:
            purge_deleted_fail += 1

    # Delete anything else (non <...>, non protected) that has 0 curve usage
    for sub in iter_user_subcats(parent_cat):
        try:
            nm = sub.Name or ""
            if nm in protected_names:
                continue

            gsid = get_subcat_gsid(sub, GTYPE)
            if gsid is None:
                continue

            if usage_after.get(gsid, 0) == 0:
                doc.Delete(sub.Id)
                purge_deleted_ok += 1
        except:
            purge_deleted_fail += 1

    t3.Commit()

# ---------------- REPORT ----------------
output.print_md("---")
output.print_md("## CLEAN SUMMARY")
output.print_md("NON-CANONICAL STYLES MAPPED:  {0}".format(mapped))
output.print_md("CURVE ELEMENTS REASSIGNED:    {0}".format(reassigned))
output.print_md("STYLES DELETED:               {0}".format(purge_deleted_ok))
output.print_md("DELETE FAILED (IN USE/LOCK):  {0}".format(purge_deleted_fail))
output.print_md("---")
output.print_md("END")

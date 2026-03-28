# -*- coding: utf-8 -*-
import Autodesk.Revit.DB as DB
from pyrevit import revit, script


doc = revit.doc
output = script.get_output()

WEIGHT_LABELS = {
    1: "Extra Fine",
    2: "Fine",
    3: "Medium",
    4: "Medium Heavy",
    5: "Heavy",
    6: "Very Heavy",
    7: "Extra Heavy",
    8: "Ultra Heavy",
}

GTYPE = DB.GraphicsStyleType.Projection


def canonical_name(weight, color_label, pattern_label):
    parts = [str(weight), WEIGHT_LABELS.get(weight, "Custom")]
    if color_label and color_label != "BLACK":
        parts.append(color_label)
    if pattern_label:
        parts.append("({0})".format(pattern_label))
    return " - ".join(parts)


def get_lines_category():
    try:
        return doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
    except Exception:
        return None


def get_subcat_exact(line_cat, name):
    if not line_cat:
        return None
    try:
        if line_cat.SubCategories.Contains(name):
            return line_cat.SubCategories.get_Item(name)
    except Exception:
        pass
    try:
        for sub in line_cat.SubCategories:
            try:
                if (sub.Name or "").strip() == name:
                    return sub
            except Exception:
                pass
    except Exception:
        pass
    return None


def find_hidden_pattern_id():
    try:
        pats = DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement)
        for p in pats:
            try:
                n = (p.Name or "").strip()
            except Exception:
                n = ""
            if n and "HIDDEN" in n.upper():
                return p.Id, n
    except Exception:
        pass
    return DB.ElementId.InvalidElementId, None


def apply_props(subcat, weight, color_obj, pattern_id):
    subcat.LineColor = color_obj
    subcat.SetLineWeight(int(weight), GTYPE)
    if pattern_id and pattern_id.IntegerValue != -1:
        subcat.SetLinePatternId(pattern_id, GTYPE)


def count_curve_usage():
    usage = {}
    try:
        curves = DB.FilteredElementCollector(doc).OfClass(DB.CurveElement).WhereElementIsNotElementType()
        for cv in curves:
            try:
                ls = cv.LineStyle
                if ls:
                    sid = ls.Id.IntegerValue
                    usage[sid] = usage.get(sid, 0) + 1
            except Exception:
                pass
    except Exception:
        pass
    return usage


def get_graphics_style_id(subcat):
    try:
        gs = subcat.GetGraphicsStyle(GTYPE)
        if gs:
            return gs.Id.IntegerValue
    except Exception:
        pass
    return None


def build_seed_specs():
    hidden_id, hidden_name = find_hidden_pattern_id()
    specs = []
    variants = [
        ("BLACK", DB.Color(0, 0, 0), None, DB.ElementId.InvalidElementId),
        ("RED", DB.Color(255, 0, 0), None, DB.ElementId.InvalidElementId),
        ("BLUE", DB.Color(0, 0, 255), None, DB.ElementId.InvalidElementId),
        ("BLACK", DB.Color(0, 0, 0), "Hidden", hidden_id),
    ]
    for w in sorted(WEIGHT_LABELS.keys()):
        for color_label, color_obj, pattern_label, pattern_id in variants:
            specs.append({
                "name": canonical_name(w, color_label, pattern_label),
                "weight": w,
                "color": color_obj,
                "pattern_label": pattern_label,
                "pattern_id": pattern_id,
            })
    return specs, hidden_name


def ensure_seed(line_cat, spec):
    name = spec["name"]
    existing = get_subcat_exact(line_cat, name)
    mode = "UPDATED"
    reason = ""

    t = DB.Transaction(doc, "MENVIC | Ensure Seed | {0}".format(name))
    t.Start()
    created_now = False
    try:
        sub = existing
        if not sub:
            sub = doc.Settings.Categories.NewSubcategory(line_cat, name)
            created_now = True
        apply_props(sub, spec["weight"], spec["color"], spec["pattern_id"])
        doc.Regenerate()
        status = t.Commit()
        if status != DB.TransactionStatus.Committed:
            return False, "FAILED", "Commit status: {0}".format(status)
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        return False, "FAILED", str(ex)

    fresh_lines = get_lines_category()
    fresh_sub = get_subcat_exact(fresh_lines, name)
    if not fresh_sub:
        return False, "FAILED", "Not found after committed transaction"

    return True, ("CREATED" if created_now else mode), reason


output.print_md("# MENVIC | LINE MASTER | SEED ONLY STANDALONE")
output.print_md("\nSEED MODE (CREATE/UPDATE CANONICAL STANDARDS ONLY)\n")

line_cat = get_lines_category()
if not line_cat:
    output.print_md("**ERROR:** Could not access OST_Lines category.")
    script.exit()

specs, hidden_pattern_name = build_seed_specs()
created = 0
updated = 0
failed = []

for spec in specs:
    ok, status, reason = ensure_seed(line_cat, spec)
    if ok:
        if status == "CREATED":
            created += 1
        else:
            updated += 1
    else:
        failed.append((spec["name"], reason))

usage = count_curve_usage()
fresh_lines = get_lines_category()
rows = []
with_use = 0
zero_use = 0
missing = 0

for spec in specs:
    name = spec["name"]
    sub = get_subcat_exact(fresh_lines, name)
    if not sub:
        rows.append((name, 0, "MISSING"))
        missing += 1
        continue
    gsid = get_graphics_style_id(sub)
    cnt = usage.get(gsid, 0) if gsid is not None else 0
    if cnt > 0:
        rows.append((name, cnt, "OK"))
        with_use += 1
    else:
        rows.append((name, 0, "ZERO USE"))
        zero_use += 1

output.print_md("\nCANONICAL (SEED) STYLES CREATED: **{0}**  ".format(created))
output.print_md("CANONICAL (SEED) STYLES UPDATED: **{0}**  ".format(updated))
output.print_md("CANONICAL (SEED) STYLES WITH USE: **{0}**  ".format(with_use))
output.print_md("CANONICAL (SEED) STYLES WITH 0 USE: **{0}**  ".format(zero_use))
output.print_md("CANONICAL (SEED) STYLES MISSING/FAILED: **{0}**\n".format(missing + len(failed)))

if hidden_pattern_name:
    output.print_md("Hidden pattern used: **{0}**\n".format(hidden_pattern_name))
else:
    output.print_md("Hidden pattern used: **NOT FOUND** (hidden seeds keep current/default pattern)\n")

output.print_md("## Canonical Seed Styles Usage")
output.print_md("| Line Style | Uses | Status |")
output.print_md("|---|---:|---|")
for name, uses, status in rows:
    output.print_md("| {0} | {1} | {2} |".format(name, uses, status))

if failed:
    output.print_md("\n## Seed creation issues")
    output.print_md("| Requested Canonical Name | Reason |")
    output.print_md("|---|---|")
    for name, reason in failed:
        output.print_md("| {0} | {1} |".format(name, reason.replace("|", "/")))

output.print_md("\nEND")

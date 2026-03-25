# -*- coding: utf-8 -*-
import Autodesk.Revit.DB as DB
import math
import re
from pyrevit import revit, script, forms


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

COLOR_MAP = {
    (0, 0, 0): "BLACK",
    (255, 255, 255): "WHITE",
    (128, 128, 128): "GRAY",
    (192, 192, 192): "SILVER",
    (255, 0, 0): "RED",
    (0, 255, 0): "LIME",
    (0, 128, 0): "GREEN",
    (0, 0, 255): "BLUE",
    (255, 255, 0): "YELLOW",
    (255, 165, 0): "ORANGE",
    (255, 140, 0): "DARKORANGE",
    (255, 127, 80): "CORAL",
    (0, 255, 255): "CYAN",
    (0, 128, 255): "DEEPSKYBLUE",
    (30, 144, 255): "DODGERBLUE",
    (0, 191, 255): "DEEPSKYBLUE",
    (128, 0, 128): "PURPLE",
    (255, 0, 255): "MAGENTA",
}

COLOR_TOL = 15.0
GTYPE = DB.GraphicsStyleType.Projection

NAME_TO_RGB = {}
for _rgb, _name in COLOR_MAP.items():
    try:
        nm = (_name or "").upper()
        if nm and nm not in NAME_TO_RGB:
            NAME_TO_RGB[nm] = _rgb
    except Exception:
        pass


# ----------------------------------------------------------
# CORE HELPERS
# ----------------------------------------------------------
def safe_int(v, default=1):
    try:
        return default if v is None else int(v)
    except Exception:
        return default


def rgb_tuple(c):
    return (int(c.Red), int(c.Green), int(c.Blue))


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


def ensure_subcategory(line_cat, name):
    try:
        if line_cat.SubCategories.Contains(name):
            return line_cat.SubCategories.get_Item(name)
        return doc.Settings.Categories.NewSubcategory(line_cat, name)
    except Exception:
        return None


def _match_root_suffix(name, root):
    try:
        if name == root:
            return 0
        pattern = r'^' + re.escape(root) + r' \((\d+)\)$'
        mm = re.match(pattern, name)
        if mm:
            return safe_int(mm.group(1), 0)
    except Exception:
        pass
    return None


def find_best_subcat_by_root(subcats, root):
    best = None
    best_suf = None
    for s in subcats:
        try:
            suf = _match_root_suffix(s.Name, root)
            if suf is None:
                continue
            if best is None or suf < best_suf:
                best = s
                best_suf = suf
        except Exception:
            continue
    return best


def get_unique_subcat_name(line_cat, base_name):
    try:
        if not line_cat.SubCategories.Contains(base_name):
            return base_name
    except Exception:
        return base_name

    root = base_name
    try:
        m = re.match(r"^(.*?)(\s\((\d+)\))?$", base_name)
        if m:
            root = m.group(1)
    except Exception:
        root = base_name

    i = 2
    while True:
        cand = "{0} ({1})".format(root, i)
        try:
            if not line_cat.SubCategories.Contains(cand):
                return cand
        except Exception:
            return cand
        i += 1


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


def get_pattern_name(cat, gtype):
    try:
        pid = cat.GetLinePatternId(gtype)
        if not pid or pid.IntegerValue == -1:
            return None, None
        pel = doc.GetElement(pid)
        if isinstance(pel, DB.LinePatternElement):
            n = (pel.Name or "").strip()
            if not n:
                return None, None
            if "HIDDEN" in n.upper():
                return "Hidden", pid
            return n, pid
    except Exception:
        pass
    return None, None


def apply_props(subcat, weight, color_obj, pattern_id, gtype=GTYPE):
    try:
        subcat.LineColor = color_obj
    except Exception:
        pass
    try:
        subcat.SetLineWeight(int(weight), gtype)
    except Exception:
        pass
    if pattern_id and pattern_id.IntegerValue != -1:
        try:
            pel = doc.GetElement(pattern_id)
            if isinstance(pel, DB.LinePatternElement):
                subcat.SetLinePatternId(pattern_id, gtype)
        except Exception:
            pass


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


def get_graphics_style_id(subcat, gtype=GTYPE):
    try:
        gs = subcat.GetGraphicsStyle(gtype)
        if gs:
            return gs.Id.IntegerValue
    except Exception:
        pass
    return None


def color_name_or_rgb(c):
    r, g, b = rgb_tuple(c)
    best, best_dist = None, COLOR_TOL
    for (rr, gg, bb), name in COLOR_MAP.items():
        dist = math.sqrt((r - rr) ** 2 + (g - gg) ** 2 + (b - bb) ** 2)
        if dist < best_dist:
            best_dist = dist
            best = name
    if best:
        return best.upper(), True
    if abs(r - g) < 15 and abs(r - b) < 15 and abs(g - b) < 15:
        return "GRAY ({0},{1},{2})".format(r, g, b), False
    if r > 150 and g > 60 and b < 100:
        return "ORANGE ({0},{1},{2})".format(r, g, b), False
    if r >= g and r >= b:
        base = "RED"
    elif g >= r and g >= b:
        base = "GREEN"
    else:
        base = "BLUE"
    return "{0} ({1},{2},{3})".format(base, r, g, b), False


def key_from_subcat(sub, gtype=GTYPE):
    try:
        w = safe_int(sub.GetLineWeight(gtype), default=1)
        c = sub.LineColor
        cname, is_named = color_name_or_rgb(c)
        if is_named:
            color_key = ("NAMED", cname)
        else:
            color_key = ("RGB", rgb_tuple(c))
        pname, _pid = get_pattern_name(sub, gtype)
        pattern_key = pname
        return (w, color_key, pattern_key), cname, pname
    except Exception:
        return None, None, None


def label_from_color_key(color_key):
    try:
        if color_key[0] == "NAMED":
            return color_key[1]
        r, g, b = color_key[1]
        if abs(r - g) < 15 and abs(r - b) < 15 and abs(g - b) < 15:
            base = "GRAY"
        elif r > 150 and g > 60 and b < 100:
            base = "ORANGE"
        elif r >= g and r >= b:
            base = "RED"
        elif g >= r and g >= b:
            base = "GREEN"
        else:
            base = "BLUE"
        return "{0} ({1},{2},{3})".format(base, r, g, b)
    except Exception:
        return None


def color_obj_from_key(color_key, fallback_color):
    try:
        if color_key[0] == "RGB":
            rr, gg, bb = color_key[1]
            return DB.Color(rr, gg, bb)
        nm = (color_key[1] or "").upper()
        if nm in NAME_TO_RGB:
            rr, gg, bb = NAME_TO_RGB[nm]
            return DB.Color(rr, gg, bb)
    except Exception:
        pass
    return fallback_color


# ----------------------------------------------------------
# SEED / STANDARD LINES
# ----------------------------------------------------------
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
    t = DB.Transaction(doc, "MENVIC | Ensure Seed | {0}".format(name))
    t.Start()
    created_now = False
    try:
        sub = existing
        if not sub:
            sub = doc.Settings.Categories.NewSubcategory(line_cat, name)
            created_now = True
        apply_props(sub, spec["weight"], spec["color"], spec["pattern_id"], GTYPE)
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

    return True, ("CREATED" if created_now else "UPDATED"), ""


def seed_canonical_standards(line_cat, gtype, seed_names):
    hidden_id, _hidden_name = find_hidden_pattern_id()
    seeds = [
        ("BLACK", DB.Color(0, 0, 0), None, DB.ElementId.InvalidElementId),
        ("RED", DB.Color(255, 0, 0), None, DB.ElementId.InvalidElementId),
        ("BLUE", DB.Color(0, 0, 255), None, DB.ElementId.InvalidElementId),
        ("BLACK", DB.Color(0, 0, 0), "Hidden", hidden_id),
    ]

    created_local = 0
    updated_local = 0

    for w in sorted(WEIGHT_LABELS.keys()):
        for color_label, color_obj, pattern_label, pattern_id in seeds:
            canon_name = canonical_name(w, color_label, pattern_label)
            seed_names.add(canon_name)

            canon_sub = None
            try:
                if line_cat.SubCategories.Contains(canon_name):
                    canon_sub = line_cat.SubCategories.get_Item(canon_name)
            except Exception:
                canon_sub = None

            if not canon_sub:
                canon_sub = find_best_subcat_by_root(list(line_cat.SubCategories), canon_name)

            existed = True if canon_sub else False

            if not canon_sub:
                canon_name_to_create = get_unique_subcat_name(line_cat, canon_name)
                canon_sub = ensure_subcategory(line_cat, canon_name_to_create)

            if canon_sub:
                try:
                    seed_names.add(canon_sub.Name)
                except Exception:
                    pass
            if not canon_sub:
                continue

            apply_props(canon_sub, w, color_obj, pattern_id, gtype)

            if existed:
                updated_local += 1
            else:
                created_local += 1

    return created_local, updated_local


def run_add_standard_lines():
    output.print_md("# MENVIC | LINE MASTER | ADD STANDARD LINES")
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
        gsid = get_graphics_style_id(sub, GTYPE)
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


# ----------------------------------------------------------
# CLEAN AND STANDARDIZE
# ----------------------------------------------------------
def run_clean_and_standardize():
    line_cat = get_lines_category()
    gtype = GTYPE
    if not line_cat:
        output.print_md("**ERROR:** Could not access OST_Lines category.")
        script.exit()

    usage_before = count_curve_usage()

    output.print_md("# MENVIC | LINE MASTER")
    output.print_md("## CLEAN AND STANDARDIZE")
    output.print_md("---")

    old_gsid_to_new_gs = {}
    subcats_to_try_delete = []
    seed_names = set()
    protected_names = set()

    created = 0
    forced = 0
    mapped = 0
    reassigned = 0

    with DB.Transaction(doc, "MENVIC: Line Master (Property Based)") as t:
        t.Start()

        c0, u0 = seed_canonical_standards(line_cat, gtype, seed_names)
        created += c0
        forced += u0
        protected_names |= set(seed_names)

        groups = {}
        for sub in list(line_cat.SubCategories):
            if not sub:
                continue
            try:
                if "<" in sub.Name:
                    continue
            except Exception:
                continue

            k, _cname, _pname = key_from_subcat(sub, gtype)
            if not k:
                continue
            if k not in groups:
                groups[k] = []
            groups[k].append(sub)

        for k, subs in groups.items():
            w, color_key, pattern_key = k

            group_has_use = False
            for s in subs:
                gsid = get_graphics_style_id(s, gtype)
                if gsid is None:
                    continue
                if usage_before.get(gsid, 0) > 0:
                    group_has_use = True
                    break

            if not group_has_use:
                for s in subs:
                    try:
                        if s.Name in seed_names:
                            continue
                        subcats_to_try_delete.append(s.Id)
                    except Exception:
                        pass
                continue

            cname = label_from_color_key(color_key)
            pname = pattern_key
            canon_name = canonical_name(w, cname, pname)

            canon_sub = find_best_subcat_by_root(subs, canon_name)
            if not canon_sub:
                try:
                    if line_cat.SubCategories.Contains(canon_name):
                        canon_sub = line_cat.SubCategories.get_Item(canon_name)
                except Exception:
                    canon_sub = None
            if not canon_sub:
                canon_sub = find_best_subcat_by_root(list(line_cat.SubCategories), canon_name)

            existed = True if canon_sub else False

            if not canon_sub:
                canon_name_to_create = get_unique_subcat_name(line_cat, canon_name)
                canon_sub = ensure_subcategory(line_cat, canon_name_to_create)

            if not canon_sub:
                continue

            try:
                canon_color = color_obj_from_key(color_key, subs[0].LineColor)
            except Exception:
                canon_color = subs[0].LineColor

            _p_label, p_id = get_pattern_name(subs[0], gtype)
            apply_props(canon_sub, w, canon_color, p_id, gtype)

            if existed:
                forced += 1
            else:
                created += 1

            protected_names.add(canon_sub.Name)

            try:
                canon_gs = canon_sub.GetGraphicsStyle(gtype)
                if not canon_gs:
                    continue
            except Exception:
                continue

            for s in subs:
                try:
                    if s.Id.IntegerValue == canon_sub.Id.IntegerValue:
                        continue
                    old_gs = s.GetGraphicsStyle(gtype)
                    if not old_gs:
                        continue
                    old_gsid_to_new_gs[old_gs.Id.IntegerValue] = canon_gs
                    subcats_to_try_delete.append(s.Id)
                    mapped += 1
                except Exception:
                    pass

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
            except Exception:
                pass

        t.Commit()

    deleted_ok = 0
    deleted_fail = 0
    usage_after = count_curve_usage()

    with DB.Transaction(doc, "MENVIC: Purge Line Styles") as t2:
        t2.Start()

        for sid in subcats_to_try_delete:
            try:
                doc.Delete(sid)
                deleted_ok += 1
            except Exception:
                deleted_fail += 1

        for sub in list(line_cat.SubCategories):
            try:
                if not sub:
                    continue
                if "<" in sub.Name:
                    continue
                if sub.Name in seed_names:
                    continue
                if sub.Name in protected_names:
                    continue

                gsid = get_graphics_style_id(sub, gtype)
                if gsid is None:
                    continue

                if usage_after.get(gsid, 0) == 0:
                    doc.Delete(sub.Id)
                    deleted_ok += 1
            except Exception:
                deleted_fail += 1

        t2.Commit()

    output.print_md("## EXECUTION SUMMARY")
    output.print_md("---")
    output.print_md("CANONICAL STYLES CREATED: {0}".format(created))
    output.print_md("CANONICAL STYLES UPDATED: {0}".format(forced))
    output.print_md("DUPLICATE STYLES MAPPED:  {0}".format(mapped))
    output.print_md("CURVE ELEMENTS UPDATED:   {0}".format(reassigned))
    output.print_md("STYLES DELETED:           {0}".format(deleted_ok))
    output.print_md("DELETE FAILED (IN USE):   {0}".format(deleted_fail))
    output.print_md("---")

    output.print_md("### TOP LINE STYLES BY USAGE (CURVEELEMENT)")

    rows = []
    for sub in list(line_cat.SubCategories):
        try:
            if not sub:
                continue
            if "<" in sub.Name:
                continue
            gsid = get_graphics_style_id(sub, gtype)
            if gsid is None:
                continue
            cnt = usage_after.get(gsid, 0)
            if cnt > 0:
                weight = safe_int(sub.GetLineWeight(gtype), 0)
                rows.append((weight, sub.Name, cnt))
        except Exception:
            pass

    rows.sort(key=lambda x: (x[0], x[1].upper()))
    for weight, name, cnt in rows[:15]:
        output.print_md("- {0}  (USE: {1})".format(name, cnt))

    output.print_md("---")
    output.print_md("END")


# ----------------------------------------------------------
# MENU
# ----------------------------------------------------------
def show_main_menu():
    options = [
        "RUN: CLEAN AND STANDARDIZE",
        "ADD STANDARD LINES",
        "CANCEL",
    ]
    try:
        return forms.CommandSwitchWindow.show(
            options,
            message="pyMenvic | Text Style Standardizer (v2.4.9)",
        )
    except Exception:
        return forms.SelectFromList.show(
            options,
            title="pyMenvic | Text Style Standardizer (v2.4.9)",
            multiselect=False,
            button_name="Run",
        )


def main():
    selection = show_main_menu()
    if not selection:
        script.exit()

    if selection == "RUN: CLEAN AND STANDARDIZE":
        run_clean_and_standardize()
    elif selection == "ADD STANDARD LINES":
        run_add_standard_lines()


if __name__ == "__main__":
    main()

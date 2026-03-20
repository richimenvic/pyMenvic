# -*- coding: utf-8 -*-
# ==========================================================
# pyMENVIC | LINE MASTER (PROPERTY-BASED STANDARDIZATION)
# Revit + pyRevit
#
# CAMBIOS CLAVE (vs tu versión):
# 1) Se protegen SOLO los SEEDS (standards base) siempre.
# 2) Los canónicos por grupo SOLO se crean/protegen si el grupo tiene USO.
# 3) Los estilos con (0) y que NO son seed, NO generan canónico y se intentan borrar.
#
# ESTRUCTURA EN BLOQUES: busca "BLOCK:" para reemplazar fácil.
# ==========================================================

import Autodesk.Revit.DB as DB
import math
from pyrevit import revit, script, forms

doc = revit.doc
output = script.get_output()

# ==========================================================
# BLOCK: SAFETY / NAME HELPERS
# ==========================================================

def is_reserved_name(name):
    try:
        n = name or ""
        return ("<" in n) or (">" in n)
    except:
        return False


def ensure_unique_name_in_cat(line_cat, desired_name):
    """Return desired_name if free; else append (2),(3)... without duplicating suffixes."""
    try:
        if not line_cat.SubCategories.Contains(desired_name):
            return desired_name
    except:
        return desired_name

    m = re.match(r"^(.*?)(\s\((\d+)\))?$", desired_name)
    root = desired_name
    if m:
        root = m.group(1)

    i = 2
    while True:
        cand = "{0} ({1})".format(root, i)
        try:
            if not line_cat.SubCategories.Contains(cand):
                return cand
        except:
            return cand
        i += 1


def find_subcat_by_root(line_cat, root_name):
    """Find exact root or root (n)."""
    if not root_name:
        return None
    try:
        if line_cat.SubCategories.Contains(root_name):
            return line_cat.SubCategories.get_Item(root_name)
    except:
        pass

    try:
        pat = re.compile(r"^" + re.escape(root_name) + r"\s\((\d+)\)$")
        for sc in list(line_cat.SubCategories):
            try:
                if pat.match(sc.Name):
                    return sc
            except:
                pass
    except:
        pass
    return None


def footprint_tuple(weight, color_obj, pattern_id):
    try:
        pid_int = pattern_id.IntegerValue if pattern_id else -1
    except:
        pid_int = -1
    try:
        return (int(weight), int(color_obj.Red), int(color_obj.Green), int(color_obj.Blue), int(pid_int))
    except:
        return (int(weight), 0, 0, 0, int(pid_int))


def subcat_footprint(sub, gtype):
    try:
        w = safe_int(sub.GetLineWeight(gtype), default=1)
    except:
        w = 1
    try:
        c = sub.LineColor
    except:
        c = DB.Color(0, 0, 0)
    try:
        _pn, pid = get_pattern_name(sub, gtype)
        pid_int = pid.IntegerValue if pid else -1
    except:
        pid_int = -1
    return (int(w), int(c.Red), int(c.Green), int(c.Blue), int(pid_int))


def find_subcat_by_footprint(line_cat, fp, gtype):
    try:
        for sc in list(line_cat.SubCategories):
            try:
                if is_reserved_name(sc.Name):
                    continue
                if subcat_footprint(sc, gtype) == fp:
                    return sc
            except:
                pass
    except:
        pass
    return None


def ensure_subcategory_safe(line_cat, name):
    """Create subcategory with safe unique naming; never touch reserved names."""
    if not name:
        return None
    if is_reserved_name(name):
        return None
    try:
        if line_cat.SubCategories.Contains(name):
            return line_cat.SubCategories.get_Item(name)
    except:
        pass

    unique = ensure_unique_name_in_cat(line_cat, name)
    if is_reserved_name(unique):
        return None
    try:
        return doc.Settings.Categories.NewSubcategory(line_cat, unique)
    except:
        return None


def stable_color_for_key(color_key):
    """Return DB.Color deterministically from color_key.
    color_key is ('RGB',(r,g,b))
    """
    try:
        rr, gg, bb = color_key
        return DB.Color(int(rr), int(gg), int(bb))
    except:
        return DB.Color(0, 0, 0)


# ==========================================================
# BLOCK: CONFIG (WEIGHTS, COLORS, TOLERANCES)
# ==========================================================
WEIGHT_LABELS = {
    1: "Extra Fine", 2: "Fine", 3: "Medium", 4: "Medium Heavy",
    5: "Heavy", 6: "Very Heavy", 7: "Extra Heavy", 8: "Ultra Heavy"
}

# COLORES CSS/X11 (MAYUSCULAS)
# Nota: (0,128,255) se etiqueta como DEEPSKYBLUE por estandar interno.
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

COLOR_TOL = 15.0  # tolerancia para "nombrar" colores por cercania


# ==========================================================
# BLOCK: UTIL (SAFE CAST, COLOR NAME, PATTERN)
# ==========================================================
def safe_int(v, default=1):
    try:
        return default if v is None else int(v)
    except:
        return default

def rgb_tuple(c):
    return (int(c.Red), int(c.Green), int(c.Blue))

def color_name_or_rgb(c):
    """
    Devuelve (label, is_named)
    is_named True => label es un nombre del COLOR_MAP (o cercano).
    """
    r, g, b = rgb_tuple(c)

    # 1) Intentar coincidencia CSS por tolerancia
    best, best_dist = None, COLOR_TOL
    for (rr, gg, bb), name in COLOR_MAP.items():
        dist = math.sqrt((r - rr) ** 2 + (g - gg) ** 2 + (b - bb) ** 2)
        if dist < best_dist:
            best_dist = dist
            best = name

    if best:
        return best.upper(), True

    # 2) Detectar GRAY (si los 3 canales son similares)
    if abs(r - g) < 15 and abs(r - b) < 15 and abs(g - b) < 15:
        return "GRAY ({0},{1},{2})".format(r, g, b), False

    # 3) Detectar ORANGE (rojo dominante + verde medio)
    if r > 150 and g > 60 and b < 100:
        return "ORANGE ({0},{1},{2})".format(r, g, b), False

    # 4) Canal dominante
    if r >= g and r >= b:
        base = "RED"
    elif g >= r and g >= b:
        base = "GREEN"
    else:
        base = "BLUE"

    return "{0} ({1},{2},{3})".format(base, r, g, b), False

def get_pattern_name(cat, gtype):
    """
    Devuelve (pattern_label, pattern_id) o (None, None)
    Normaliza Hidden.
    """
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
        return None, None
    except:
        return None, None


# ==========================================================
# BLOCK: NAMING (CANONICAL NAME, GROUP KEY)
# ==========================================================
def canonical_name(weight, color_label, pattern_label):
    """
    Convención:
      "<w> - <WeightLabel> - <COLOR> - (<Pattern>)"
    Regla: si el color es BLACK, no se añade al nombre.
    """
    parts = ["{0}".format(weight), WEIGHT_LABELS.get(weight, "Custom")]
    if color_label and color_label != "BLACK":
        parts.append(color_label)
    if pattern_label:
        parts.append("({0})".format(pattern_label))
    return " - ".join(parts)

def key_from_subcat(sub, gtype=DB.GraphicsStyleType.Projection):
    """Key estable por propiedades REALES: (weight, rgb_tuple, patternIdInt).
    Retorna: (key, color_label_for_name, pattern_label_for_name)
    """
    try:
        w = safe_int(sub.GetLineWeight(gtype), default=1)

        c = sub.LineColor
        rgb = rgb_tuple(c)

        # label para nombre (usa tolerancia), pero NO para agrupar
        cname, _is_named = color_name_or_rgb(c)

        pname, pid = get_pattern_name(sub, gtype)
        try:
            pid_int = pid.IntegerValue if pid else -1
        except:
            pid_int = -1

        return (int(w), rgb, int(pid_int)), cname, pname
    except:
        return None, None, None


def apply_props(sub, weight, color_obj, pattern_id, gtype=DB.GraphicsStyleType.Projection):
    # Color
    try:
        sub.LineColor = color_obj
    except:
        pass

    # Weight
    try:
        sub.SetLineWeight(int(weight), gtype)
    except:
        pass

    # Pattern (solo si es válido)
    if pattern_id and pattern_id.IntegerValue != -1:
        try:
            pel = doc.GetElement(pattern_id)
            if isinstance(pel, DB.LinePatternElement):
                sub.SetLinePatternId(pattern_id, gtype)
        except:
            pass

def find_hidden_pattern_id():
    """Devuelve ElementId del patrón que contenga 'Hidden' (si existe)."""
    try:
        pats = DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement)
        for p in pats:
            n = (p.Name or "").strip()
            if n and "HIDDEN" in n.upper():
                return p.Id
    except:
        pass
    return DB.ElementId.InvalidElementId


# ==========================================================
# BLOCK: USAGE SCAN (CURVEELEMENT USAGE)
# NOTA: Revit puede usar estilos en más sitios (anotaciones/familias).
# El Delete() es la verdad final: si falla, está en uso o bloqueado.
# ==========================================================
def get_subcat_gsid(sub, gtype):
    try:
        gs = sub.GetGraphicsStyle(gtype)
        if not gs:
            return None
        return gs.Id.IntegerValue
    except:
        return None

def count_curve_usage():
    """
    Cuenta uso por GraphicsStyleId en CurveElement (líneas modelo/detalle).
    """
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


# ==========================================================
# BLOCK: SEED CANONICAL STANDARDS (BASE STANDARDS)
# Crea base por weight (1..8):
#   - RED
#   - BLUE
#   - (Hidden) [sin color en el nombre]
# ==========================================================
def seed_canonical_standards(line_cat, gtype, seed_names):
    hidden_id = find_hidden_pattern_id()

    # Seeds base por peso:
    # - BLACK continuo (sin patrón)  -> crea "1 - Extra Fine", etc.
    # - RED continuo
    # - BLUE continuo
    # - BLACK Hidden                 -> crea "... - (Hidden)"
    seeds = [
        ("BLACK", DB.Color(0, 0, 0), None,     DB.ElementId.InvalidElementId),
        ("RED",   DB.Color(255, 0, 0), None,   DB.ElementId.InvalidElementId),
        ("BLUE",  DB.Color(0, 0, 255), None,   DB.ElementId.InvalidElementId),
        ("BLACK", DB.Color(0, 0, 0),  "Hidden", hidden_id),
    ]

    created_local = 0
    updated_local = 0

    for w in sorted(WEIGHT_LABELS.keys()):
        for color_label, color_obj, pattern_label, pattern_id in seeds:
            canon_name = canonical_name(w, color_label, pattern_label)

            # Reusar por nombre (root o root (n))
            canon_sub = find_subcat_by_root(line_cat, canon_name)

            # Reusar por huella
            if canon_sub is None:
                fp = footprint_tuple(w, color_obj, pattern_id)
                canon_sub = find_subcat_by_footprint(line_cat, fp, gtype)

            if canon_sub is None:
                canon_sub = ensure_subcategory_safe(line_cat, canon_name)
                if not canon_sub:
                    continue
                created_local += 1
            else:
                updated_local += 1

            try:
                seed_names.add(canon_sub.Name)  # proteger nombre real
            except:
                seed_names.add(canon_name)

            apply_props(canon_sub, w, color_obj, pattern_id, gtype)

    return created_local, updated_local


# ==========================================================
# BLOCK: UI
# ==========================================================
options = [
    "RUN: CLEAN AND STANDARDIZED",
    "ADD SEED (CREATE/UPDATE CANONICALS)",
    "CANCEL"
]

choice = forms.CommandSwitchWindow.show(
    options,
    title="pyMENVIC | LINE MASTER",
    message="PROPERTY-BASED STANDARDIZATION (WEIGHT + COLOR + PATTERN)"
)

if choice is None or choice == "CANCEL":
    script.exit()

SEED_ONLY = (choice == "ADD SEED (CREATE/UPDATE CANONICALS)")


# ==========================================================
# BLOCK: CONTEXT
# ==========================================================
line_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
gtype = DB.GraphicsStyleType.Projection

usage_before = count_curve_usage()


# ==========================================================
# BLOCK: SEED ONLY MODE
# ==========================================================
if SEED_ONLY:
    output.print_md("# pyMENVIC | LINE MASTER")
    output.print_md("## SEED MODE (CREATE/UPDATE CANONICAL STANDARDS)")
    output.print_md("---")

    seed_names = set()
    created = 0
    updated = 0

    with DB.Transaction(doc, "pyMENVIC: Seed Canonical Line Styles") as ts:
        ts.Start()
        c0, u0 = seed_canonical_standards(line_cat, gtype, seed_names)
        created += c0
        updated += u0
        ts.Commit()

    output.print_md("CANONICAL (SEED) STYLES CREATED: {0}".format(created))
    output.print_md("CANONICAL (SEED) STYLES UPDATED: {0}".format(updated))
    output.print_md("---")
    output.print_md("END")
    script.exit()


# ==========================================================
# BLOCK: EXECUTION (CLEAN)
# ==========================================================
output.print_md("# pyMENVIC | LINE MASTER")
output.print_md("## EXECUTION STARTED")
output.print_md("---")

old_gsid_to_new_gs = {}     # old GraphicsStyleId(int) -> new GraphicsStyle
subcats_to_try_delete = []  # [(ElementId, Name)] candidatas a borrar
failed_deletes = []  # (name, reason)

seed_names = set()          # SOLO seeds intocables
protected_names = set()     # seeds + canónicos CON USO (no borrar)

created = 0
forced = 0
mapped = 0
reassigned = 0

with DB.Transaction(doc, "pyMENVIC: Line Master (Property Based)") as t:
    t.Start()

    # 0) SEED: canónicos base (RED / BLUE / (Hidden))
    c0, u0 = seed_canonical_standards(line_cat, gtype, seed_names)
    created += c0
    forced += u0
    protected_names |= set(seed_names)

    # 1) AGRUPAR por propiedades
    groups = {}
    for sub in list(line_cat.SubCategories):
        if not sub:
            continue
        try:
            if is_reserved_name(sub.Name):  # built-in / reservado
                continue
        except:
            continue

        k, cname, pname = key_from_subcat(sub, gtype)
        if not k:
            continue
        groups.setdefault(k, []).append((sub, cname, pname))

    # 2) POR GRUPO: solo crear/proteger canónico si el grupo TIENE USO
    for k, items in groups.items():
        subs = [it[0] for it in items]
        # labels para nombre: usar la primera ocurrencia (determinístico por sort)
        try:
            items_sorted = sorted(items, key=lambda x: (x[0].Name or "").upper())
            cname = items_sorted[0][1]
            pname = items_sorted[0][2]
        except:
            cname = items[0][1]
            pname = items[0][2]
        w, color_key, pattern_key = k

        # --- BLOQUE CLAVE: medir uso del grupo (CurveElement) ---
        group_has_use = False
        for s in subs:
            gsid = get_subcat_gsid(s, gtype)
            if gsid is None:
                continue
            if usage_before.get(gsid, 0) > 0:
                group_has_use = True
                break

        # Si NO hay uso y NO es seed -> NO crear canónico. Intentar borrar todos.
        if not group_has_use:
            for s in subs:
                try:
                    if s.Name in seed_names:
                        continue
                    subcats_to_try_delete.append((s.Id, s.Name))
                except:
                    pass
            continue

        # Si hay uso: asegurar canónico, protegerlo, y mapear duplicados al canónico.
        canon_name = canonical_name(w, cname, pname)

        existed = False
        try:
            existed = line_cat.SubCategories.Contains(canon_name)
        except:
            existed = False

        canon_sub = ensure_subcategory_safe(line_cat, canon_name)
        if not canon_sub:
            continue

        # color "real" del canónico (determinístico desde la key RGB)
        try:
            rr, gg, bb = color_key
            canon_color = DB.Color(int(rr), int(gg), int(bb))
        except:
            canon_color = subs[0].LineColor

        # patrón desde la key (determinístico)
        try:
            pid_int = pattern_key
            p_id = DB.ElementId(pid_int) if int(pid_int) != -1 else DB.ElementId.InvalidElementId
        except:
            p_id = DB.ElementId.InvalidElementId

        apply_props(canon_sub, w, canon_color, p_id, gtype)

        if existed:
            forced += 1
        else:
            created += 1

        # proteger este canónico (tiene uso real)
        protected_names.add(canon_name)

        # graphics style canónico
        try:
            canon_gs = canon_sub.GetGraphicsStyle(gtype)
            if not canon_gs:
                continue
        except:
            continue

        # mapear todos los del grupo que no sean el canónico
        for s in subs:
            try:
                if s.Id.IntegerValue == canon_sub.Id.IntegerValue:
                    continue
                old_gs = s.GetGraphicsStyle(gtype)
                if not old_gs:
                    continue
                old_gsid_to_new_gs[old_gs.Id.IntegerValue] = canon_gs
                subcats_to_try_delete.append((s.Id, s.Name))
                mapped += 1
            except:
                pass

    # 3) REASIGNAR CurveElement
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

    t.Commit()


# ==========================================================
# BLOCK: PURGE
# ==========================================================
deleted_ok = 0
deleted_fail = 0

usage_after = count_curve_usage()

# TOP rows (CurveElement) para reporte
top_rows = []
try:
    for sub in list(line_cat.SubCategories):
        try:
            if not sub:
                continue
            if is_reserved_name(sub.Name):
                continue
            gsid = get_subcat_gsid(sub, gtype)
            if gsid is None:
                continue
            cnt = usage_after.get(gsid, 0)
            if cnt > 0:
                top_rows.append((sub.Name, cnt))
        except:
            pass
    top_rows.sort(key=lambda x: (-x[1], x[0].upper()))
except:
    top_rows = []

t2 = DB.Transaction(doc, "pyMENVIC: Purge Line Styles")
t2.Start()

# 4.1) borrar duplicados mapeados + todos los "sin uso" que metimos a lista
seen_del = set()
for sid, sname in subcats_to_try_delete:
    try:
        key = sid.IntegerValue
    except:
        key = None

    if key is not None and key in seen_del:
        continue
    if key is not None:
        seen_del.add(key)

    try:
        doc.Delete(sid)
        deleted_ok += 1
    except Exception as ex:
        deleted_fail += 1
        reason = str(ex).split("\\n")[0]
        failed_deletes.append((sname, reason))

# 4.2) borrar lo que quede sin uso (CurveElement) y que NO sea seed ni canónico con uso
for sub in list(line_cat.SubCategories):
    try:
        if not sub:
            continue
        if is_reserved_name(sub.Name):
            continue
        if sub.Name in seed_names:
            continue
        if sub.Name in protected_names:
            continue

        gsid = get_subcat_gsid(sub, gtype)
        if gsid is None:
            continue

        if usage_after.get(gsid, 0) == 0:
            doc.Delete(sub.Id)
            deleted_ok += 1
    except:
        deleted_fail += 1

t2.Commit()


# ==========================================================
# BLOCK: REPORT
# ==========================================================
output.print_md("# pyMENVIC | LINE MASTER — {}".format(choice))
output.print_md("**Line Style Standardization Report**")
output.print_md("## Resumen")
output.print_md("")
output.print_md("- 🆕 Canonical styles creados: **{}**".format(created))
output.print_md("- 🔄 Canonical styles actualizados: **{}**".format(forced))
output.print_md("- 🔁 Duplicates mapeados: **{}**".format(mapped))
output.print_md("- ✏ CurveElements reasignados: **{}**".format(reassigned))
output.print_md("- 🗑 Styles eliminados: **{}**".format(deleted_ok))
output.print_md("- ⚠ Delete fallido (in use / protegido): **{}**".format(deleted_fail))

total_changes = created + forced + mapped + reassigned + deleted_ok
output.print_md("---")

# TOP (si hay datos)
if top_rows:
    output.print_md("## Top Line Styles por uso (CurveElement)")
    output.print_md("| Line Style | Usos |")
    output.print_md("|---|---:|")
    for name, cnt in top_rows[:15]:
        output.print_md("| {} | {} |".format(name, cnt))

# Deletes fallidos (detalle)
if failed_deletes:
    output.print_md("## ⚠ Styles que no pudieron eliminarse")
    output.print_md("| Line Style | Motivo |")
    output.print_md("|---|---|")
    seen = set()
    for name, reason in failed_deletes:
        k = "{}||{}".format(name, reason)
        if k in seen:
            continue
        seen.add(k)
        output.print_md("| {} | {} |".format(name, reason))

if total_changes == 0 and deleted_fail == 0:
    output.print_md("### ✅ No se requirieron cambios. Todo ya estaba estandarizado.")
else:
    output.print_md("### ✔ Proceso completado.")

output.print_md("---")
output.print_md("END")
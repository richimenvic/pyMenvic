# -*- coding: utf-8 -*-
# ==========================================================
# MENVIC | LINE MASTER (PROPERTY-BASED STANDARDIZATION)
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
import re
from pyrevit import revit, script, forms

doc = revit.doc
output = script.get_output()


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

# Reverse map for stable named colors (first RGB wins)
NAME_TO_RGB = {}
for _rgb, _name in COLOR_MAP.items():
    try:
        nm = (_name or "").upper()
        if nm and nm not in NAME_TO_RGB:
            NAME_TO_RGB[nm] = _rgb
    except:
        pass


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
    """
    Key de unificacion por propiedades:
    (weight, color_key, pattern_key)
      color_key: ("NAMED","GRAY") o ("RGB",(r,g,b))
      pattern_key: None o "Hidden" o "Dash 1/8\""
    Retorna: (key, cname_label, pname_label)
    """
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
    except:
        return None, None, None



def label_from_color_key(color_key):
    """Stable label used for naming from a color_key."""
    try:
        if color_key[0] == "NAMED":
            return color_key[1]
        r, g, b = color_key[1]
        if abs(r - g) < 15 and abs(r - b) < 15 and abs(g - b) < 15:
            base = "GRAY"
        elif r > 150 and g > 60 and b < 100:
            base = "ORANGE"
        else:
            if r >= g and r >= b:
                base = "RED"
            elif g >= r and g >= b:
                base = "GREEN"
            else:
                base = "BLUE"
        return "{0} ({1},{2},{3})".format(base, r, g, b)
    except:
        return None

def color_obj_from_key(color_key, fallback_color):
    """Return DB.Color from color_key with stable exact RGB for named colors."""
    try:
        if color_key[0] == "RGB":
            rr, gg, bb = color_key[1]
            return DB.Color(rr, gg, bb)
        nm = (color_key[1] or "").upper()
        if nm in NAME_TO_RGB:
            rr, gg, bb = NAME_TO_RGB[nm]
            return DB.Color(rr, gg, bb)
    except:
        pass
    return fallback_color

# ==========================================================
# BLOCK: REVIT HELPERS (ENSURE SUBCAT, APPLY PROPS)
# ==========================================================
def ensure_subcategory(line_cat, name):
    try:
        if line_cat.SubCategories.Contains(name):
            return line_cat.SubCategories.get_Item(name)
        return doc.Settings.Categories.NewSubcategory(line_cat, name)
    except:
        return None


def _match_root_suffix(name, root):
    """Return suffix int if name matches root or root (n). None if not."""
    try:
        if name == root:
            return 0
        # root may contain parentheses, escape in regex
        pattern = r'^' + re.escape(root) + r' \((\d+)\)$'
        mm = re.match(pattern, name)
        if mm:
            return safe_int(mm.group(1), 0)
    except:
        pass
    return None

def find_best_subcat_by_root(subcats, root):
    """Pick existing subcategory matching root or root (n) with smallest suffix."""
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
        except:
            continue
    return best

def get_unique_subcat_name(line_cat, base_name):
    """Return a name that does not collide. Uses ' (2)', ' (3)'... and avoids duplicating suffixes."""
    try:
        if not line_cat.SubCategories.Contains(base_name):
            return base_name
    except:
        return base_name

    # strip existing numeric suffix if present
    root = base_name
    try:
        m = re.match(r"^(.*?)(\s\((\d+)\))?$", base_name)
        if m:
            root = m.group(1)
    except:
        root = base_name

    i = 2
    while True:
        cand = "{0} ({1})".format(root, i)
        try:
            if not line_cat.SubCategories.Contains(cand):
                return cand
        except:
            return cand
        i += 1

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

    seeds = [
        ("BLACK", DB.Color(0, 0, 0),   None,     DB.ElementId.InvalidElementId),  # base (sin color en nombre)
        ("RED",   DB.Color(255, 0, 0), None,     DB.ElementId.InvalidElementId),
        ("BLUE",  DB.Color(0, 0, 255), None,     DB.ElementId.InvalidElementId),
        ("BLACK", DB.Color(0, 0, 0),   "Hidden", hidden_id),
    ]

    created_local = 0
    updated_local = 0

    for w in sorted(WEIGHT_LABELS.keys()):
        for color_label, color_obj, pattern_label, pattern_id in seeds:
            canon_name = canonical_name(w, color_label, pattern_label)
            # proteger el nombre ideal (aunque se cree con sufijo por colisión)
            seed_names.add(canon_name)

            # Reusar si ya existe un seed con la misma "raíz" (root o root (n))
            canon_sub = None
            try:
                if line_cat.SubCategories.Contains(canon_name):
                    canon_sub = line_cat.SubCategories.get_Item(canon_name)
            except:
                canon_sub = None

            if not canon_sub:
                # buscar por raíz (incluye sufijos)
                canon_sub = find_best_subcat_by_root(list(line_cat.SubCategories), canon_name)

            existed = True if canon_sub else False

            if not canon_sub:
                # crear con nombre libre (evita auto-renombre raro de Revit)
                canon_name_to_create = get_unique_subcat_name(line_cat, canon_name)
                canon_sub = ensure_subcategory(line_cat, canon_name_to_create)

            if canon_sub:
                # proteger el nombre real existente/creado
                try:
                    seed_names.add(canon_sub.Name)
                except:
                    pass
            if not canon_sub:
                continue

            apply_props(canon_sub, w, color_obj, pattern_id, gtype)

            if existed:
                updated_local += 1
            else:
                created_local += 1

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
    title="MENVIC | LINE MASTER",
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
    output.print_md("# MENVIC | LINE MASTER")
    output.print_md("## SEED MODE (CREATE/UPDATE CANONICAL STANDARDS)")
    output.print_md("---")

    seed_names = set()
    created = 0
    updated = 0

    with DB.Transaction(doc, "MENVIC: Seed Canonical Line Styles") as ts:
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
output.print_md("# MENVIC | LINE MASTER")
output.print_md("## EXECUTION STARTED")
output.print_md("---")

old_gsid_to_new_gs = {}     # old GraphicsStyleId(int) -> new GraphicsStyle
subcats_to_try_delete = []  # ids candidatas a borrar

seed_names = set()          # SOLO seeds intocables
protected_names = set()     # seeds + canónicos CON USO (no borrar)

created = 0
forced = 0
mapped = 0
reassigned = 0

with DB.Transaction(doc, "MENVIC: Line Master (Property Based)") as t:
    t.Start()

    # 0) SEED: canónicos base (RED / BLUE / (Hidden))
    c0, u0 = seed_canonical_standards(line_cat, gtype, seed_names)
    created += c0
    forced += u0
    protected_names |= set(seed_names)

    # 1) AGRUPAR por propiedades (SOLO por huella k)
    # Importante: NO incluir cname/pname en la llave del dict,
    # porque puede variar por tolerancias y romper idempotencia.
    groups = {}
    for sub in list(line_cat.SubCategories):
        if not sub:
            continue
        try:
            if "<" in sub.Name:  # built-in o reservado
                continue
        except:
            continue

        k, _cname, _pname = key_from_subcat(sub, gtype)
        if not k:
            continue
        if k not in groups:
            groups[k] = []
        groups[k].append(sub)

    # 2) POR GRUPO: solo crear/proteger canónico si el grupo TIENE USO
    for k, subs in groups.items():
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
                    subcats_to_try_delete.append(s.Id)
                except:
                    pass
            continue

        # Si hay uso: asegurar canónico, protegerlo, y mapear duplicados al canónico.
        cname = label_from_color_key(color_key)
        pname = pattern_key
        canon_name = canonical_name(w, cname, pname)

        # Preferir un canónico existente en este mismo grupo (root o root (n)).
        canon_sub = find_best_subcat_by_root(subs, canon_name)

        # Si no existe en el grupo, buscar en el proyecto por raíz (por si el grupo cambió de miembros)
        if not canon_sub:
            try:
                if line_cat.SubCategories.Contains(canon_name):
                    canon_sub = line_cat.SubCategories.get_Item(canon_name)
            except:
                canon_sub = None
        if not canon_sub:
            canon_sub = find_best_subcat_by_root(list(line_cat.SubCategories), canon_name)

        existed = True if canon_sub else False

        # Si aun no existe, crear SOLO porque este grupo tiene uso real
        if not canon_sub:
            canon_name_to_create = get_unique_subcat_name(line_cat, canon_name)
            canon_sub = ensure_subcategory(line_cat, canon_name_to_create)

        if not canon_sub:
            continue

        # color "real" del canónico (estable)
        try:
            canon_color = color_obj_from_key(color_key, subs[0].LineColor)
        except:
            canon_color = subs[0].LineColor

        # patrón desde el primero del grupo
        p_label, p_id = get_pattern_name(subs[0], gtype)
        apply_props(canon_sub, w, canon_color, p_id, gtype)

        if existed:
            forced += 1
        else:
            created += 1

        # proteger este canónico (tiene uso real)
        protected_names.add(canon_sub.Name)

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
                subcats_to_try_delete.append(s.Id)
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

with DB.Transaction(doc, "MENVIC: Purge Line Styles") as t2:
    t2.Start()

    # 4.1) borrar duplicados mapeados + todos los "sin uso" que metimos a lista
    for sid in subcats_to_try_delete:
        try:
            doc.Delete(sid)
            deleted_ok += 1
        except:
            deleted_fail += 1

    # 4.2) borrar lo que quede sin uso (CurveElement) y que NO sea seed ni canónico con uso
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
output.print_md("# MENVIC | LINE MASTER")
output.print_md("## EXECUTION SUMMARY")
output.print_md("---")
output.print_md("CANONICAL STYLES CREATED: {0}".format(created))
output.print_md("CANONICAL STYLES UPDATED: {0}".format(forced))
output.print_md("DUPLICATE STYLES MAPPED:  {0}".format(mapped))
output.print_md("CURVE ELEMENTS UPDATED:   {0}".format(reassigned))
output.print_md("STYLES DELETED:           {0}".format(deleted_ok))
output.print_md("DELETE FAILED (IN USE):   {0}".format(deleted_fail))
output.print_md("---")

# TOP 15 USOS DESPUES (ORDEN POR LINE WEIGHT)
output.print_md("### TOP LINE STYLES BY USAGE (CURVEELEMENT)")

rows = []
for sub in list(line_cat.SubCategories):
    try:
        if not sub:
            continue
        if "<" in sub.Name:
            continue

        gsid = get_subcat_gsid(sub, gtype)
        if gsid is None:
            continue

        cnt = usage_after.get(gsid, 0)
        if cnt > 0:
            weight = safe_int(sub.GetLineWeight(gtype), 0)
            rows.append((weight, sub.Name, cnt))
    except:
        pass

rows.sort(key=lambda x: (x[0], x[1].upper()))

for weight, name, cnt in rows[:15]:
    output.print_md("- {0}  (USE: {1})".format(name, cnt))

output.print_md("---")
output.print_md("END")
# -*- coding: utf-8 -*-
__title__  = "Standardize Text Styles"
__author__ = "Ricardo J. Mendieta"

"""
==========================================================
pyMENVIC | TEXT STYLE STANDARDIZER
Revit + pyRevit

Descripción
-----------
Herramienta avanzada para estandarizar los estilos de texto
(TextNoteType) dentro del proyecto según el estándar de
oficina definido para pyMENVIC.

El script puede:

- limpiar y normalizar fuentes
- ajustar tamaños de texto a estándares
- fusionar tipos duplicados
- renombrar tipos según su configuración real
- eliminar tipos no utilizados
- crear automáticamente los estilos estándar de oficina

Estándar de Oficina
-------------------
Fuente: Arial  
Unidades: milímetros  
Background: Opaque  
Width Scale: 1.0  
Sin Bold / Italic / Underline  
Sin Box  
Arrowhead: Arrow 30 Degree

Tamaños estándar:
1.80, 2.00, 2.20, 2.50, 3.00, 3.50, 4.00, 4.50,
5.00, 5.50, 6.00, 7.00, 8.00 mm

Reglas importantes
------------------
- "(T)" en el nombre indica fondo transparente.
- TEXT_BACKGROUND:
    1 = Transparent
    0 = Opaque
- El parámetro LEADER_ARROWHEAD no se modifica
  directamente (solo lectura).
- Si existe un tipo con el mismo nombre pero con
  una flecha distinta, se crea un nuevo estándar
  limpio en lugar de modificar el existente.

Funciones principales
---------------------
RUN: CLEAN AND STANDARDIZE
    Limpia, normaliza, renombra, fusiona y purga estilos.

ADD STANDARD FONTS
    Crea los estilos estándar de oficina.

RENAME ONLY
    Solo renombra los estilos según su configuración.

Autor
-----
Ricardo J. Mendieta  
pyMENVIC – Ayudas para MENVIC ARQ
==========================================================
"""

import Autodesk.Revit.DB as DB
from pyrevit import revit, script, forms
import re
import System
from System.Reflection import BindingFlags

doc = revit.doc
output = script.get_output()

# ============================================================
# CONFIG
# ============================================================

# NO usar ElementId fijo para A30 (varía por plantilla)
DEF_ARROW30_ID = None  # desactivado

# Opciones principales
MERGE_DUPLICATES = True
SNAP_TEXT_SIZES  = True
SNAP_TOL_MM      = 0.06

EXCLUDE_NAME_CONTAINS = ["default"]

# Renombrado
INCLUDE_WIDTH_IN_NAME_IF_NOT_1 = True
ARIAL_NARROW_THRESHOLD = 0.85    # Arial comprimido -> Arial Narrow (si existe)

# Office standard (creación)
OFFICE_STANDARD_CREATE = True
OFFICE_STANDARD_FONT = "Arial"

# >>> STANDARD SOLO MM, SIN 11.70, SOLO OPACO <<<
OFFICE_STANDARD_SIZES_MM = [
    1.80, 2.00, 2.20, 2.50, 3.00, 3.50, 4.00, 4.50, 5.00, 5.50, 6.00, 7.00, 8.00
]

# Pulgadas típicas SOLO para detectar y etiquetar en nombre (NO para snap)
# Incluye 7/64" porque te apareció ~2.73mm
MAPEO_PULGADAS = {
    0.396875: '1/64"',
    0.79375:  '1/32"',
    1.5875:   '1/16"',
    2.38125:  '3/32"',
    2.778125: '7/64"',
    3.175:    '1/8"',
    3.96875:  '5/32"',
    4.7625:   '3/16"',
    6.35:     '1/4"',
    7.9375:   '5/16"',
    9.525:    '3/8"',
    11.1125:  '7/16"',
    12.7:     '1/2"',
}
TOL_PULGADAS_MM = 0.001  # tolerancia para reconocer etiqueta en pulgadas

CANON_FONT_CASE = {
    "calibri": "Calibri",
    "arial": "Arial",
    "arial narrow": "Arial Narrow",
    "arial black": "Arial Black",
    "romans": "RomanS",
    "romans-1": "RomanS",
    "romans.shx": "RomanS",
    "swis721 lt bt": "Swis721 Lt BT",
}

# ============================================================
# HELPERS
# ============================================================

def norm(s):
    return re.sub(r"\s+", " ", (s or "")).strip()

def low(s):
    return norm(s).lower()

def is_excluded_name(name):
    return any(x in low(name) for x in EXCLUDE_NAME_CONTAINS)

# Revit se queja con: [] {} | ; < > ? ` ~
_ILLEGAL_NAME_CHARS = set(list(r'{}[]|;<>?`~'))

def sanitize_name(name):
    if not name:
        return ""
    out = "".join((" " if c in _ILLEGAL_NAME_CHARS else c) for c in name)
    return norm(out)

def get_type_name(tt):
    try:
        p = tt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return norm(p.AsString() or "")
    except:
        pass
    return ""

def set_clr_name(obj, new_name):
    t = obj.GetType()
    t.InvokeMember(
        "Name",
        BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty,
        None,
        obj,
        System.Array[object]([new_name])
    )

def ensure_unique_name(desired_name, existing_names):
    if desired_name not in existing_names:
        return desired_name

    m = re.match(r"^(.*)\s\((\d+)\)$", desired_name)
    if m:
        stem = m.group(1)
        i = int(m.group(2)) + 1
    else:
        stem = desired_name
        i = 2

    while True:
        cand = "{} ({})".format(stem, i)
        if cand not in existing_names:
            return cand
        i += 1

def mm_to_ft(mm):
    return float(mm) / 304.8

def ft_to_mm(ft):
    return float(ft) * 304.8

def _nearest_key(value, keys, tol):
    best_key = None
    best_d = 1e9
    for k in keys:
        d = abs(float(value) - float(k))
        if d < best_d:
            best_d = d
            best_key = float(k)
    if best_key is not None and best_d <= tol:
        return best_key
    return None

def detect_inch_fraction(size_mm):
    key = _nearest_key(size_mm, MAPEO_PULGADAS.keys(), TOL_PULGADAS_MM)
    if key is None:
        return None, None
    return float(key), MAPEO_PULGADAS[float(key)]

def size_label_mm_with_inch(size_mm_value):
    mmv = float(size_mm_value)
    inch_mm, frac = detect_inch_fraction(mmv)
    if frac:
        return "{:.2f}mm ({})".format(round(inch_mm, 2), frac)
    return "{:.2f}mm".format(round(mmv, 2))

def snap_mm_to_standard_only(size_mm):
    size_mm = float(size_mm)
    best = None
    best_d = 1e9
    for c in OFFICE_STANDARD_SIZES_MM:
        d = abs(size_mm - float(c))
        if d < best_d:
            best_d = d
            best = float(c)
    if best is not None and best_d <= SNAP_TOL_MM:
        return best, True
    return size_mm, False

# ============================================================
# ARROW (READ ONLY)
# ============================================================

def get_arrow_id(tt):
    try:
        p = tt.get_Parameter(DB.BuiltInParameter.LEADER_ARROWHEAD)
        if p:
            return p.AsElementId()
    except:
        pass
    return DB.ElementId.InvalidElementId

def element_display_name(el):
    if not el:
        return ""
    try:
        n = norm(getattr(el, "Name", "") or "")
        if n:
            return n
    except:
        pass
    try:
        p = el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            n = norm(p.AsString() or "")
            if n:
                return n
    except:
        pass
    for pname in ["Type Name", "Nombre de tipo"]:
        try:
            p = el.LookupParameter(pname)
            if p:
                n = norm(p.AsString() or "")
                if n:
                    return n
        except:
            pass
    return ""

def arrow_name_from_id(aid):
    if not aid or aid == DB.ElementId.InvalidElementId:
        return ""
    try:
        el = doc.GetElement(aid)
    except:
        el = None
    return element_display_name(el)

def is_default_arrow(aid, a_name):
    # STANDARD = siempre Arrow 30 Degree
    return low(a_name) == "arrow 30 degree"

def type_has_default_arrow(tt):
    try:
        aid = get_arrow_id(tt)
        an = arrow_name_from_id(aid)
        return is_default_arrow(aid, an)
    except:
        return False

def _extract_frac(arrow_name):
    """
    Extract fraction like 1/16 from arrowhead names.
    Supports:
      - 1/16"
      - 1 / 16"
      - 1/16″ (double-prime) or smart quotes
      - fraction without quotes
    """
    s = arrow_name or ""
    m = re.search(r'(\d+\s*/\s*\d+)\s*["\u2033\u201d]', s)
    if m:
        return m.group(1).replace(" ", "")
    m2 = re.search(r'(\d+\s*/\s*\d+)', s)
    return m2.group(1).replace(" ", "") if m2 else None

def _extract_mm(arrow_name):
    s = (arrow_name or "").lower()
    m = re.search(r'(\d+(?:\.\d+)?)\s*mm', s)
    return m.group(1) if m else None

def arrow_code_from_name(a_name):
    n = norm(a_name)
    ln = n.lower()
    if not n:
        return ""

    # Default A30 (no se escribe)
    if ln == "arrow 30 degree":
        return "A30"

    frac = _extract_frac(n)
    mmv  = _extract_mm(n)

    def _mm_trim(v):
        try:
            return str(float(v)).rstrip("0").rstrip(".")
        except:
            return str(v).rstrip("0").rstrip(".")

    def _unit_suffix_no_underscore():
        if frac:
            return "_" + frac
        if mmv:
            return _mm_trim(mmv)
        return ""

    is_mep = ln.startswith("mep") or ("mep -" in ln)

    if "arrow filled" in ln:
        m = re.search(r'(\d+)', ln)
        code = "AF{}".format(m.group(1)) if m else "AF"
        return ("MEP_" + code) if is_mep else code

    if "arrow open" in ln and "90" in ln:
        if frac:
            return 'AO90_{}"'.format(frac)
        if mmv:
            mm_txt = str(float(mmv)).rstrip("0").rstrip(".")
            return "AO90_{}mm".format(mm_txt)
        return "AO90"

    if ln.startswith("arrow") and ("degree" in ln or "deg" in ln or "°" in ln):
        m = re.search(r'(\d+)', ln)
        if m:
            return "A{}".format(m.group(1))
        return "A"

    if ("dot filled" in ln) or ("filled dot" in ln):
        suf = _unit_suffix_no_underscore()
        return "DF{}".format(suf) if suf else "DF"

    if ("dot open" in ln) or ("open dot" in ln):
        suf = _unit_suffix_no_underscore()
        return "DO{}".format(suf) if suf else "DO"

    if "diagonal" in ln:
        suf = _unit_suffix_no_underscore()
        return "DG{}".format(suf) if suf else "DG"

    if ("box filled" in ln) or ("filled box" in ln):
        suf = _unit_suffix_no_underscore()
        return "BX{}".format(suf) if suf else "BX"

    if "heavy end" in ln:
        suf = _unit_suffix_no_underscore()
        return "HE{}".format(suf) if suf else "HE"

    if "filled elevation target" in ln:
        suf = _unit_suffix_no_underscore()
        return "ET{}".format(suf) if suf else "ET"

    if "triangle filled" in ln:
        suf = _unit_suffix_no_underscore()
        return "TRF{}".format(suf) if suf else "TRF"

    fallback = re.sub(r"\s+", "", n)
    return fallback[:24]

# ============================================================
# FONT / WIDTH
# ============================================================

try:
    from System.Drawing.Text import InstalledFontCollection
    _ifc = InstalledFontCollection()
    INSTALLED_FONTS = set([norm(f.Name) for f in _ifc.Families])
except:
    INSTALLED_FONTS = set()

def is_font_installed(font_name):
    if not INSTALLED_FONTS:
        return False
    return bool(font_name) and (norm(font_name) in INSTALLED_FONTS)

def canonicalize_case_only(font_value):
    if not font_value:
        return None
    f = norm(font_value)
    return CANON_FONT_CASE.get(f.lower(), None)

def to_title_case_font(font_value):
    if not font_value:
        return ""
    return " ".join([w[:1].upper() + w[1:].lower() for w in norm(font_value).split()])

def fix_text_font(tt):
    try:
        p_font = tt.get_Parameter(DB.BuiltInParameter.TEXT_FONT)
        p_width = tt.get_Parameter(DB.BuiltInParameter.TEXT_WIDTH_SCALE)

        if not p_font:
            return (False, "TEXT_FONT missing")
        if p_font.IsReadOnly:
            return (False, "TEXT_FONT readonly")

        font = norm(p_font.AsString() or "")
        if not font:
            return (False, None)

        canon = canonicalize_case_only(font)
        normalized_font = canon if canon else to_title_case_font(font)

        if p_width and (not p_width.IsReadOnly):
            try:
                width = float(p_width.AsDouble())
            except:
                width = 1.0

            if normalized_font == "Arial" and width <= ARIAL_NARROW_THRESHOLD and is_font_installed("Arial Narrow"):
                p_font.Set("Arial Narrow")
                p_width.Set(1.0)
                return (True, "Arial->Narrow")

        if normalized_font and normalized_font != font:
            p_font.Set(normalized_font)
            return (True, "case-normalized")

        return (False, None)

    except Exception as ex:
        return (False, str(ex).splitlines()[0])

# ============================================================
# SIZE SNAP
# ============================================================

def snap_text_size_param(tt):
    try:
        p = tt.get_Parameter(DB.BuiltInParameter.TEXT_SIZE)
        if not p or p.IsReadOnly:
            return (False, "TEXT_SIZE missing/readonly")

        size_mm = ft_to_mm(p.AsDouble())
        snapped, did = snap_mm_to_standard_only(size_mm)
        if did:
            p.Set(mm_to_ft(snapped))
            return (True, "snapped {:.4f}->{:.2f}mm".format(size_mm, snapped))
        return (False, None)
    except Exception as ex:
        return (False, str(ex).splitlines()[0])

# ============================================================
# FOOTPRINT + NAME
# ============================================================

def has_show_border(tt):
    for pname in ["Show Border", "Mostrar borde", "Border", "Borde"]:
        try:
            p = tt.LookupParameter(pname)
            if p and p.StorageType == DB.StorageType.Integer and p.AsInteger() == 1:
                return True
        except:
            pass
    return False

def get_footprint(tt):
    try:
        p_font = tt.get_Parameter(DB.BuiltInParameter.TEXT_FONT)
        font = norm(p_font.AsString() or "") if p_font else ""
    except:
        font = ""

    try:
        size_ft = tt.get_Parameter(DB.BuiltInParameter.TEXT_SIZE).AsDouble()
    except:
        size_ft = 0.0

    try:
        width = float(tt.get_Parameter(DB.BuiltInParameter.TEXT_WIDTH_SCALE).AsDouble())
    except:
        width = 1.0

    try:
        bold = tt.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_BOLD).AsInteger()
    except:
        bold = 0
    try:
        italic = tt.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_ITALIC).AsInteger()
    except:
        italic = 0
    try:
        underline = tt.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_UNDERLINE).AsInteger()
    except:
        underline = 0

    try:
        bg = tt.get_Parameter(DB.BuiltInParameter.TEXT_BACKGROUND).AsInteger()
    except:
        bg = 0

    arrow_id = get_arrow_id(tt)
    box = 1 if has_show_border(tt) else 0

    size_mm = ft_to_mm(size_ft)
    snapped_mm, _ = snap_mm_to_standard_only(size_mm)
    snapped_ft = mm_to_ft(snapped_mm)

    return (
        font,
        round(snapped_ft, 10),
        round(width, 6),
        bold, italic, underline,
        bg,
        arrow_id.IntegerValue if arrow_id else -1,
        box
    )

def make_standard_name_from_fp(fp):
    font, size_ft, width, bold, italic, underline, bg, arrow_int, box = fp
    size_mm = ft_to_mm(size_ft)

    base = "{} {}".format(font, size_label_mm_with_inch(size_mm))

    if INCLUDE_WIDTH_IN_NAME_IF_NOT_1 and abs(width - 1.0) > 0.001:
        wtxt = "{:.2f}".format(width).rstrip("0").rstrip(".")
        base += " W{}".format(wtxt)

    if box == 1:
        base += " w/Box"
    if bold == 1:
        base += " B"
    if italic == 1:
        base += " I"
    if underline == 1:
        base += " L"

    try:
        aid = DB.ElementId(arrow_int) if arrow_int != -1 else DB.ElementId.InvalidElementId
    except:
        aid = DB.ElementId.InvalidElementId

    a_name = arrow_name_from_id(aid)
    if aid != DB.ElementId.InvalidElementId:
        if not is_default_arrow(aid, a_name):
            code = arrow_code_from_name(a_name)
            if code:
                base += " " + code

    if bg == 1:
        base += " (T)"

    return sanitize_name(norm(base))

def _is_close(a, b, tol=0.02):
    try:
        return abs(float(a) - float(b)) <= float(tol)
    except:
        return False

def _get_bg_int(tt):
    try:
        p = tt.get_Parameter(DB.BuiltInParameter.TEXT_BACKGROUND)
        return p.AsInteger() if p else None
    except:
        return None

def _get_width(tt):
    try:
        p = tt.get_Parameter(DB.BuiltInParameter.TEXT_WIDTH_SCALE)
        return float(p.AsDouble()) if p else None
    except:
        return None

def _get_font(tt):
    try:
        p = tt.get_Parameter(DB.BuiltInParameter.TEXT_FONT)
        return norm(p.AsString() or "") if p else ""
    except:
        return ""

def _get_size_mm(tt):
    try:
        p = tt.get_Parameter(DB.BuiltInParameter.TEXT_SIZE)
        return ft_to_mm(p.AsDouble()) if p else None
    except:
        return None

# ============================================================
# OFFICE STANDARD CREATION
# ============================================================

def find_existing_office_type(size_mm, require_name=None, strict_clean=True):
    """
    Superset finder. If strict_clean=True, requires a "clean" office type:
      - Font Arial
      - Width 1.0
      - Background Opaque (0)
      - Bold/Italic/Underline = 0
      - Show Border = 0
      - Arrowhead default A30
      - Size within SNAP_TOL_MM
    """
    target = float(size_mm)

    def _get_int(tt, bip):
        try:
            p = tt.get_Parameter(bip)
            return p.AsInteger() if p else 0
        except:
            return 0

    def _get_show_border(tt):
        for pname in ["Show Border", "Mostrar borde", "Border", "Borde"]:
            try:
                p = tt.LookupParameter(pname)
                if p and p.StorageType == DB.StorageType.Integer:
                    return p.AsInteger()
            except:
                pass
        return 0

    def _is_default_arrow_of_type(tt):
        try:
            aid = get_arrow_id(tt)
            an = arrow_name_from_id(aid)
            return is_default_arrow(aid, an)
        except:
            return False

    for tt in DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType):
        try:
            if require_name and norm(get_type_name(tt)) != norm(require_name):
                continue
            if norm(_get_font(tt)).lower() != norm(OFFICE_STANDARD_FONT).lower():
                continue
            if not _is_close(_get_width(tt), 1.0, tol=0.01):
                continue
            if _get_bg_int(tt) != 0:
                continue

            if strict_clean:
                if _get_int(tt, DB.BuiltInParameter.TEXT_STYLE_BOLD) != 0:
                    continue
                if _get_int(tt, DB.BuiltInParameter.TEXT_STYLE_ITALIC) != 0:
                    continue
                if _get_int(tt, DB.BuiltInParameter.TEXT_STYLE_UNDERLINE) != 0:
                    continue
                if _get_show_border(tt) != 0:
                    continue
                if not _is_default_arrow_of_type(tt):
                    continue

            smm = _get_size_mm(tt)
            if smm is None:
                continue
            if abs(float(smm) - target) > SNAP_TOL_MM:
                continue
            return tt
        except:
            continue
    return None

def standard_name(size_mm):
    return sanitize_name("{} {:.2f}mm".format(OFFICE_STANDARD_FONT, float(size_mm)))

def find_base_type_for_duplication(types):
    # 1) Ideal: Arial + Arrow 30 Degree
    for tt in types:
        n = get_type_name(tt)
        if is_excluded_name(n):
            continue
        try:
            p = tt.get_Parameter(DB.BuiltInParameter.TEXT_FONT)
            f = p.AsString() if p else ""
        except:
            f = ""
        if norm(f).lower() == "arial" and type_has_default_arrow(tt):
            return tt

    # 2) Cualquier fuente, pero Arrow 30 Degree
    for tt in types:
        n = get_type_name(tt)
        if is_excluded_name(n):
            continue
        if type_has_default_arrow(tt):
            return tt

    # 3) Si no hay base con Arrow 30 Degree, no crear standards
    return None

def create_office_standard_types():
    created = []
    failed = []

    types = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))
    existing_names = set(norm(get_type_name(t)) for t in types)

    base = find_base_type_for_duplication(types)
    if not base:
        failed.append(("ALL", "", "No base TextNoteType found with Arrow 30 Degree"))
        return created, failed

    for size_mm in OFFICE_STANDARD_SIZES_MM:
        new_name = standard_name(size_mm)
        legacy_name = sanitize_name("STANDARD {} {:.2f}mm".format(OFFICE_STANDARD_FONT, float(size_mm)))

        # 1) Si ya existe uno limpio de ese tamaño (con Arrow 30 Degree), no crear otro
        tt_clean_any = find_existing_office_type(size_mm, require_name=None, strict_clean=True)
        if tt_clean_any is not None:
            office_skip_log.append((float(size_mm), "exists-clean-any", get_type_name(tt_clean_any)))
            continue

        # 2) Si el legacy existe y es limpio, renombrar si se puede
        tt_legacy_clean = find_existing_office_type(size_mm, require_name=legacy_name, strict_clean=True)
        if tt_legacy_clean is not None:
            try:
                if norm(new_name) not in existing_names:
                    set_clr_name(tt_legacy_clean, new_name)
                    existing_names.discard(norm(legacy_name))
                    existing_names.add(norm(new_name))
                    office_skip_log.append((float(size_mm), "renamed-legacy->new", legacy_name))
                else:
                    office_skip_log.append((float(size_mm), "legacy-clean-exists-name-taken", legacy_name))
            except:
                office_skip_log.append((float(size_mm), "legacy-clean-exists-rename-failed", legacy_name))
            continue

        # 3) Si el nombre standard ya existe pero NO es limpio, NO tocarlo.
        #    Crear otro limpio con nombre único.
        if norm(new_name) in existing_names:
            office_skip_log.append((float(size_mm), "name-exists-but-not-clean", new_name))
            new_name = ensure_unique_name(new_name, existing_names)

        # 4) Si el legacy existe pero NO es limpio, tampoco tocarlo.
        #    Solo crear otro limpio con nombre único.
        if norm(legacy_name) in existing_names:
            office_skip_log.append((float(size_mm), "legacy-name-exists-but-not-clean", legacy_name))

        try:
            new_type = base.Duplicate(new_name)

            p_font = new_type.get_Parameter(DB.BuiltInParameter.TEXT_FONT)
            if p_font and not p_font.IsReadOnly:
                p_font.Set(OFFICE_STANDARD_FONT)

            p_size = new_type.get_Parameter(DB.BuiltInParameter.TEXT_SIZE)
            if p_size and not p_size.IsReadOnly:
                p_size.Set(mm_to_ft(size_mm))

            p_w = new_type.get_Parameter(DB.BuiltInParameter.TEXT_WIDTH_SCALE)
            if p_w and not p_w.IsReadOnly:
                p_w.Set(1.0)

            for bip in [DB.BuiltInParameter.TEXT_STYLE_BOLD,
                        DB.BuiltInParameter.TEXT_STYLE_ITALIC,
                        DB.BuiltInParameter.TEXT_STYLE_UNDERLINE]:
                p = new_type.get_Parameter(bip)
                if p and not p.IsReadOnly:
                    p.Set(0)

            p_bg = new_type.get_Parameter(DB.BuiltInParameter.TEXT_BACKGROUND)
            if p_bg and not p_bg.IsReadOnly:
                p_bg.Set(0)

            for pname in ["Show Border", "Mostrar borde", "Border", "Borde"]:
                p = new_type.LookupParameter(pname)
                if p and p.StorageType == DB.StorageType.Integer:
                    try:
                        p.Set(0)
                    except:
                        pass

            # Importante:
            # NO tocamos LEADER_ARROWHEAD.
            # El tipo nuevo hereda Arrow 30 Degree porque la base elegida ya lo tiene.

            created.append(new_name)
            existing_names.add(norm(new_name))
        except Exception as ex:
            failed.append((str(size_mm), new_name, str(ex).splitlines()[0]))

    return created, failed

# ============================================================
# UI
# ============================================================

modo = forms.CommandSwitchWindow.show(
    ["RUN: CLEAN AND STANDARDIZE", "ADD STANDARD FONTS", "RENAME ONLY", "CANCEL"],
    message=(
        "pyMenvic | Text Style Standardizer (v2.4.9)\n\n"
        "IMPORTANTE (medido): TEXT_BACKGROUND 1=Transparent, 0=Opaque\n"
        "(T) = Transparent\n"
        "STANDARD = siempre Arrow 30 Degree\n"
        "Leader Arrowhead NO se modifica (solo lectura).\n"
    ),
    title="pyMenvic"
)

if not modo or modo == "CANCEL":
    script.exit()

# ============================================================
# MAIN
# ============================================================

all_notes = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNote).WhereElementIsNotElementType())

count_font_fixed = 0
count_size_snapped = 0
count_groups_merged = 0
count_types_deleted = 0
count_notes_moved = 0
count_renamed = 0
count_rename_failed = 0
count_purged = 0

created_standard = []
failed_standard = []
office_skip_log = []
failed_rename = []
log_size = []

def ejecutar(rename_only=False):
    global count_font_fixed, count_size_snapped
    global count_groups_merged, count_types_deleted, count_notes_moved
    global count_renamed, count_rename_failed, count_purged
    global created_standard, failed_standard
    global failed_rename, log_size

    types_now = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))

    # A) Fix fonts
    for tt in types_now:
        name = get_type_name(tt)
        if is_excluded_name(name):
            continue
        changed, _ = fix_text_font(tt)
        if changed:
            count_font_fixed += 1

    # B) Snap text sizes
    if (not rename_only) and SNAP_TEXT_SIZES:
        for tt in types_now:
            name = get_type_name(tt)
            if is_excluded_name(name):
                continue
            changed, reason = snap_text_size_param(tt)
            if changed:
                count_size_snapped += 1
                log_size.append((tt.Id.IntegerValue, name, reason))

    # C) Merge duplicates
    if (not rename_only) and MERGE_DUPLICATES:
        types_now = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))
        fp_to_types = {}
        for tt in types_now:
            name = get_type_name(tt)
            if is_excluded_name(name):
                continue
            try:
                fp = get_footprint(tt)
                fp_to_types.setdefault(fp, []).append(tt)
            except:
                pass

        fp_canonical = {}
        for fp, tlist in fp_to_types.items():
            fp_canonical[fp] = sorted(tlist, key=lambda x: x.Id.IntegerValue)[0]

        for fp, tlist in fp_to_types.items():
            if len(tlist) <= 1:
                continue
            canonical = fp_canonical.get(fp)
            if not canonical:
                continue

            canon_id = canonical.Id
            for tt in tlist:
                if tt.Id == canon_id:
                    continue
                tid = tt.Id

                for n in all_notes:
                    try:
                        if n.GetTypeId() == tid:
                            n.ChangeTypeId(canon_id)
                            count_notes_moved += 1
                    except:
                        pass

                try:
                    doc.Delete(tid)
                    count_types_deleted += 1
                except:
                    pass

            count_groups_merged += 1

    # D) Rename
    types_after = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))
    existing_names = set(norm(get_type_name(t)) for t in types_after)

    for tt in types_after:
        old_name = norm(get_type_name(tt))
        if not old_name or is_excluded_name(old_name):
            continue

        try:
            fp = get_footprint(tt)
        except:
            continue

        desired = make_standard_name_from_fp(fp)
        if not desired:
            continue

        if low(desired) == low(old_name):
            continue

        new_name = ensure_unique_name(desired, existing_names)

        try:
            set_clr_name(tt, new_name)
            after = norm(get_type_name(tt))
            if after == new_name:
                count_renamed += 1
                existing_names.add(new_name)
            else:
                count_rename_failed += 1
                failed_rename.append((tt.Id.IntegerValue, old_name, new_name, "did-not-stick (after='{}')".format(after)))
        except Exception as ex:
            count_rename_failed += 1
            failed_rename.append((tt.Id.IntegerValue, old_name, new_name, "rename-failed: {}".format(str(ex).splitlines()[0])))

    if rename_only:
        return

    # E) Purge unused (except STANDARD)
    used_type_ids = set([n.GetTypeId() for n in all_notes])
    protected_office_names = set(norm(standard_name(mm)) for mm in OFFICE_STANDARD_SIZES_MM)
    protected_office_names |= set(norm(sanitize_name("STANDARD {} {:.2f}mm".format(OFFICE_STANDARD_FONT, float(mm))))
                                  for mm in OFFICE_STANDARD_SIZES_MM)

    for tt in list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType)):
        name_tt = get_type_name(tt)
        if is_excluded_name(name_tt):
            continue
        if norm(name_tt) in protected_office_names:
            continue
        if tt.Id not in used_type_ids:
            try:
                doc.Delete(tt.Id)
                count_purged += 1
            except:
                pass

    # F) Create office STANDARD
    if OFFICE_STANDARD_CREATE:
        c, f = create_office_standard_types()
        created_standard.extend(c)
        failed_standard.extend(f)

with revit.Transaction("pyMenvic | Text Style Standardizer (v2.4.9)"):
    if modo == "RUN: CLEAN AND STANDARDIZE":
        ejecutar(rename_only=False)
    elif modo == "ADD STANDARD FONTS":
        c, f = create_office_standard_types()
        created_standard.extend(c)
        failed_standard.extend(f)
    elif modo == "RENAME ONLY":
        ejecutar(rename_only=True)

# ============================================================
# REPORT
# ============================================================

output.print_md("# pyMenvic | Text Style Standardizer (v2.4.9) — {}".format(modo))
output.print_md("**(T)=Transparent** | TEXT_BACKGROUND: 1=Transparent, 0=Opaque")
output.print_md("**STANDARD:** Arial + Opaque + W1 + sin B/I/L + sin Box + **Arrow 30 Degree**")
output.print_md("## Resumen")
output.print_md("- 🔤 Fuentes ajustadas: **{}**".format(count_font_fixed))
output.print_md("- 📏 Text Size snapped: **{}**".format(count_size_snapped))
output.print_md("- 🔁 Grupos merged: **{}**".format(count_groups_merged))
output.print_md("- 🏷 Tipos renombrados: **{}**".format(count_renamed))
output.print_md("- 🧹 Tipos purgados: **{}**".format(count_purged))
output.print_md("- 📐 Standards creados: **{}**".format(len(created_standard)))
output.print_md("---")

if office_skip_log and (created_standard or count_renamed):
    output.print_md("## Office standards verificados")
    output.print_md("| Size(mm) | Motivo | Match |")
    output.print_md("|---:|---|---|")
    for (mm, reason, match) in office_skip_log[:80]:
        output.print_md("| {:.2f} | {} | {} |".format(mm, reason, match))

if failed_standard:
    output.print_md("## Fallos relevantes")
    output.print_md("| Elemento | Acción | Motivo |")
    output.print_md("|---|---|---|")
    for itm, act, why in failed_standard[:80]:
        output.print_md("| {} | Create Standard | {} |".format(itm, why))

if failed_rename:
    output.print_md("## Renames con fallo")
    output.print_md("| Elemento | Acción | Motivo |")
    output.print_md("|---|---|---|")
    for rid, oldn, newn, why in failed_rename[:80]:
        output.print_md("| {} / {} | Rename -> {} | {} |".format(rid, oldn, newn, why))

if (count_size_snapped + count_font_fixed + count_groups_merged +
    count_renamed + count_purged + len(created_standard)) == 0:
    output.print_md("### ✅ Todo ya estaba estandarizado.")
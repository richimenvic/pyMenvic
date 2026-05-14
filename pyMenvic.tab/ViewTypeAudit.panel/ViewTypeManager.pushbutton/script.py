# -*- coding: utf-8 -*-


__title__ = "View Type Standardizer"
__author__ = "Ricardo J. Mendieta"

import os
import sys
import re
import clr

from pyrevit import revit, DB, script, forms

try:
    from lib.core.branding import get_logo_path
except ImportError:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    while True:
        if os.path.basename(current_dir).lower() == "pymenvic.extension":
            lib_dir = os.path.join(current_dir, "lib")
            if lib_dir not in sys.path:
                sys.path.append(lib_dir)
            break
        parent_dir = os.path.dirname(current_dir)
        if parent_dir == current_dir:
            break
        current_dir = parent_dir
    from core.branding import get_logo_path


clr.AddReference("System")
clr.AddReference("System.Core")
clr.AddReference("System.Windows")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")

from System import Uri, UriKind
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption


doc = revit.doc
output = script.get_output()



XAML_FILE = script.get_bundle_file("ui.xaml")
TXT_FILE = script.get_bundle_file("view_type_standards.txt")
LOGO_FILE = get_logo_path()


# ==================================================
# HELPERS
# ==================================================

def safe_str(value):
    try:
        if value is None:
            return ""
        return str(value)
    except:
        try:
            return unicode(value)
        except:
            return ""


def normalize_text(value):
    return safe_str(value).strip()


def normalize_key(value):
    return normalize_text(value).lower()


def first_line_error(ex):
    try:
        return safe_str(ex).splitlines()[0]
    except:
        return "Unknown error"


def get_type_name(elem):
    try:
        return DB.Element.Name.GetValue(elem)
    except:
        try:
            return elem.Name
        except:
            return "<Unnamed>"


def get_view_family_name(vft):
    try:
        return safe_str(vft.ViewFamily)
    except:
        return "Unknown"


def is_protected_name(name):
    name = normalize_text(name)
    if not name:
        return True
    if name.startswith("<") and name.endswith(">"):
        return True
    return False


def clear_collection(collection):
    while collection.Count > 0:
        collection.RemoveAt(0)


def sort_alpha(values):
    return sorted(values, key=lambda x: normalize_key(x))


def get_combobox_text(combo):
    try:
        if combo.SelectedItem is not None:
            return normalize_text(combo.SelectedItem)
    except:
        pass
    try:
        return normalize_text(combo.Text)
    except:
        return ""


def make_brush(hex_color):
    hex_color = hex_color.replace("#", "")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return SolidColorBrush(Color.FromRgb(r, g, b))


def set_combo_items(column, items):
    try:
        column.ItemsSource = items
    except:
        pass


def refresh_grid(grid):
    try:
        grid.Items.Refresh()
    except:
        try:
            source = grid.ItemsSource
            grid.ItemsSource = None
            grid.ItemsSource = source
        except:
            pass


def load_logo_if_available(window):
    try:
        logo_path = LOGO_FILE
        if not logo_path or not os.path.exists(logo_path):
            return

        bitmap = BitmapImage()
        bitmap.BeginInit()
        bitmap.CacheOption = BitmapCacheOption.OnLoad
        bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
        bitmap.EndInit()
        window.logoImage.Source = bitmap
    except:
        pass


def strip_accents(text):
    replacements = {
        u"Á": "A", u"À": "A", u"Ä": "A", u"Â": "A",
        u"É": "E", u"È": "E", u"Ë": "E", u"Ê": "E",
        u"Í": "I", u"Ì": "I", u"Ï": "I", u"Î": "I",
        u"Ó": "O", u"Ò": "O", u"Ö": "O", u"Ô": "O",
        u"Ú": "U", u"Ù": "U", u"Ü": "U", u"Û": "U",
        u"á": "a", u"à": "a", u"ä": "a", u"â": "a",
        u"é": "e", u"è": "e", u"ë": "e", u"ê": "e",
        u"í": "i", u"ì": "i", u"ï": "i", u"î": "i",
        u"ó": "o", u"ò": "o", u"ö": "o", u"ô": "o",
        u"ú": "u", u"ù": "u", u"ü": "u", u"û": "u",
        u"Ñ": "N", u"ñ": "n"
    }
    text = safe_str(text)
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    return text


def normalize_display_name(text):
    text = strip_accents(text)
    text = text.replace("_", " ")
    text = text.replace("/", " ")
    text = re.sub(r"\s+", " ", text).strip().upper()

    replacements = [
        ("APOBACION", "APROBACION"),
        ("SENALITICA", "SENALETICA"),
        ("SEÑALETICA", "SENALETICA"),
        ("SEÑALETICA", "SENALETICA"),
        ("SEÑALÉTICA", "SENALETICA"),
        ("SEÑALETICA", "SENALETICA"),
        ("ELEVACION", "ELEVACIONES"),
        ("ELEVATIONS", "ELEVACIONES"),
        ("SECTIONS", "SECCIONES"),
        ("SECTION", "SECCIONES"),
        ("PLANS", "PLANOS"),
        ("DETAILS", "DETALLES"),
        ("AMPLIADO", "AMPLIADOS"),
        ("ACABADOS FINALES", "ACABADOS"),
        ("BORDE DE LA LOSA", "BORDE DE LOSA"),
        ("EXTERIOR", "EXTERIORES"),
        ("INTERIOR", "INTERIORES"),
        ("DETALLES", "DETALLE"),
        ("ESCALERAS", "ESCALERA"),
        ("ISOMETRICO", "ISOMETRICOS"),
        ("MOBILIARIOS", "MOBILIARIO"),
        ("CUARTO TECNICO", "CUARTOS TECNICOS"),
        ("CUARTOS TECNICO", "CUARTOS TECNICOS"),
        ("ELECTRICOS", "ELECTRICOS"),
        ("MECANICOS", "MECANICOS"),
    ]
    for src, dst in replacements:
        text = text.replace(src, dst)

    text = re.sub(r"\s+", " ", text).strip()
    return text


# ==================================================
# DEFAULT TXT
# ==================================================

def ensure_txt_exists():
    if os.path.exists(TXT_FILE):
        return

    default_text = """# Discipline|ViewFamily|OldTypeName|NewTypeName
Architecture|FloorPlan|WORKING|00 WORKING
Architecture|Elevation|00 ELEVACIONES - WORKING|00 WORKING
Architecture|FloorPlan|01 PLANOS|01 PLANOS - GENERALES
Architecture|FloorPlan|01 PLANOS - ACOTADOS|02 PLANOS - ACOTADOS
Architecture|FloorPlan|01 PLANOS - ALCALDIA|07 PLANOS - APROBACION MUNICIPAL
Architecture|FloorPlan|01 PLANOS - BORDE DE LA LOSA|03 PLANOS - BORDE DE LOSA
Architecture|FloorPlan|02 PLANOS - AMOBLADOS|04 PLANOS - AMOBLADOS
Architecture|FloorPlan|03 PLANOS - CUBIERTA|05 PLANOS - CUBIERTA
Architecture|FloorPlan|04 PLANOS - AMPLIADOS|06 PLANOS - AMPLIADOS
Architecture|FloorPlan|05 PLANOS - ACABADOS FINALES|08 PLANOS - ACABADOS
Architecture|FloorPlan|06 PLANOS PARA SENALETICA|10 PLANOS - SENALETICA
Architecture|CeilingPlan|01 PLANOS DEL CIELO REFLEJADO|01 PLANOS - CIELO REFLEJADO
Architecture|CeilingPlan|02 PLANOS DE CIELO REFLEJADO - AMPLIADO|02 PLANOS - CIELO REFLEJADO - AMPLIADOS
Architecture|Elevation|01 ELEVACIONES - EXTERIOR|01 ELEVACIONES - EXTERIORES
Architecture|Elevation|02 ELEVACIONES - INTERIOR|02 ELEVACIONES - INTERIORES
Architecture|Elevation|03 ELEVACIONES - DETALLES|03 ELEVACIONES - DETALLE
Architecture|Elevation|01 ELEVACIONES - ALCALDIA|04 ELEVACIONES - APROBACION MUNICIPAL
Architecture|Section|01 SECCIONES - EDIFICIO COMPLETO|01 SECCIONES - GENERALES
Architecture|Section|02 SECCIONES - CORTE DE BORDE|02 SECCIONES - DE BORDE
Architecture|Section|03 SECCIONES - DETALLES|03 SECCIONES - DE DETALLE
Architecture|Section|04 SECCIONES - ESCALERAS|04 SECCIONES - DE ESCALERA
Architecture|Section|01 SECCIONES - ALCALDIA|05 SECCIONES - APROBACION MUNICIPAL
Architecture|Drafting|Detail|01 DETALLES - GENERALES
Architecture|Drafting|06 DETALLES|01 DETALLES - GENERALES
Architecture|Drafting|06 DETALLES PLANTA|09 DETALLES - PLANTA
Architecture|Drafting|02 DETALLE MUEBLES|02 DETALLES - DE MOBILIARIO
Architecture|Drafting|A_06_DETAIL|03 DETALLES - ESPECIALES
Architecture|ThreeDimensional|3D View|03 3D - GENERAL
"""
    f = open(TXT_FILE, "w")
    try:
        f.write(default_text)
    finally:
        f.close()


# ==================================================
# DATA CLASSES
# ==================================================

class StandardRuleRow(object):
    def __init__(self, discipline, family_name, old_name, new_name):
        self.Discipline = normalize_text(discipline)
        self.ViewFamilyName = normalize_text(family_name)
        self.OldTypeName = normalize_text(old_name)
        self.NewTypeName = normalize_text(new_name)


class ProjectTypeRow(object):
    def __init__(self, element, family_name, current_name):
        self.Element = element
        self.ElementId = element.Id.IntegerValue
        self.ViewFamilyName = normalize_text(family_name)
        self.CurrentTypeName = normalize_text(current_name)
        self.NewTypeName = ""
        self.AddRule = False
        self.Include = False
        self.Status = "UNMATCHED"
        self.StatusBg = make_brush("#8E4B4B")
        self.StatusFg = make_brush("#FFFFFF")

    def set_status(self, status):
        self.Status = status
        color_map = {
            "MATCHED": "#416172",
            "NORMALIZED": "#6A5A7A",
            "NUMBERED": "#8A6D3B",
            "UNMATCHED": "#7A4C4C",
            "NO CHANGE": "#5A5A5A",
            "RENAMED": "#4E7A56",
            "SKIPPED": "#666666",
            "ALREADY EXISTS": "#7A5A4C",
            "DUPLICATE TARGET": "#7A5A4C",
            "FAILED": "#A94442",
        }
        self.StatusBg = make_brush(color_map.get(status, "#7A4C4C"))
        self.StatusFg = make_brush("#FFFFFF")


# ==================================================
# STANDARDS IO
# ==================================================

def read_txt_standards():
    ensure_txt_exists()

    discipline_map = {}
    errors = []

    f = open(TXT_FILE, "r")
    try:
        lines = f.readlines()
    finally:
        f.close()

    line_number = 0
    for raw_line in lines:
        line_number += 1
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        parts = [normalize_text(x) for x in line.split("|")]
        if len(parts) != 4:
            errors.append("Line {} invalid. Expected 4 parts.".format(line_number))
            continue

        discipline, family, old_name, new_name = parts
        if not discipline or not family or not old_name or not new_name:
            errors.append("Line {} invalid. Empty values.".format(line_number))
            continue

        if discipline not in discipline_map:
            discipline_map[discipline] = []

        discipline_map[discipline].append(StandardRuleRow(discipline, family, old_name, new_name))

    for key in discipline_map.keys():
        discipline_map[key] = sorted(
            discipline_map[key],
            key=lambda x: (normalize_key(x.ViewFamilyName), normalize_key(x.OldTypeName))
        )

    return discipline_map, errors


def write_txt_standards(discipline_map):
    disciplines = sort_alpha(discipline_map.keys())
    lines = ["# Discipline|ViewFamily|OldTypeName|NewTypeName"]

    first_block = True
    for discipline in disciplines:
        rows = discipline_map.get(discipline, [])
        rows = sorted(rows, key=lambda x: (normalize_key(x.ViewFamilyName), normalize_key(x.OldTypeName)))

        if not first_block:
            lines.append("")
        first_block = False

        for row in rows:
            discipline_name = normalize_text(row.Discipline)
            family = normalize_text(row.ViewFamilyName)
            old_name = normalize_text(row.OldTypeName)
            new_name = normalize_text(row.NewTypeName)

            if not discipline_name or not family or not old_name or not new_name:
                continue

            lines.append("{}|{}|{}|{}".format(discipline_name, family, old_name, new_name))

    f = open(TXT_FILE, "w")
    try:
        f.write("\n".join(lines))
    finally:
        f.close()


# ==================================================
# PROJECT SCAN
# ==================================================

def collect_project_view_types():
    rows = []
    failures = []

    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
    for vft in collector:
        try:
            rows.append(ProjectTypeRow(vft, get_view_family_name(vft), get_type_name(vft)))
        except Exception as ex:
            failures.append({
                "element": "ViewFamilyType",
                "action": "Scan",
                "reason": first_line_error(ex)
            })

    rows = sorted(rows, key=lambda x: (normalize_key(x.ViewFamilyName), normalize_key(x.CurrentTypeName)))
    return rows, failures


# ==================================================
# MATCH + SMART SUGGESTIONS
# ==================================================

def build_lookup(rule_rows):
    lookup = {}
    for row in rule_rows:
        family_key = normalize_key(row.ViewFamilyName)
        old_key = normalize_key(normalize_display_name(row.OldTypeName))
        new_name = normalize_text(row.NewTypeName)
        if family_key and old_key and new_name:
            lookup[(family_key, old_key)] = new_name
    return lookup


def family_category_label(view_family_name):
    vf = normalize_key(view_family_name)
    mapping = {
        "floorplan": "PLANOS",
        "structuralplan": "PLANOS",
        "areaplan": "PLANOS",
        "ceilingplan": "PLANOS",
        "elevation": "ELEVACIONES",
        "section": "SECCIONES",
        "detail": "DETALLES",
        "drafting": "DETALLES",
        "threedimensional": "3D",
    }
    return mapping.get(vf, "")


def parse_numbered_standard_name(name):
    text = normalize_display_name(name)
    match = re.match(r"^(\d{2})\s+(.+)$", text)
    if not match:
        return None, None, None

    num = match.group(1)
    remainder = match.group(2).strip()
    if " - " in remainder:
        category = remainder.split(" - ", 1)[0].strip()
        subcat = remainder.split(" - ", 1)[1].strip()
    else:
        category = remainder
        subcat = ""
    return num, category, subcat


def normalize_subcategory_from_text(text):
    text = normalize_display_name(text)
    text = text.replace(" - ", " ")
    text = re.sub(r"^\d{2}\s+", "", text).strip()
    return text


def infer_subcategory(view_family_name, current_name):
    family_key = normalize_key(view_family_name)
    text = normalize_display_name(current_name)

    keyword_sets = {
        "elevation": [
            ("APROBACION MUNICIPAL", ["ALCALDIA", "APROBACION MUNICIPAL", "MUNICIPAL"]),
            ("EXTERIORES", ["EXTERIOR", "EXTERIORES"]),
            ("INTERIORES", ["INTERIOR", "INTERIORES"]),
            ("DETALLE", ["DETALLE", "DETALLES"]),
            ("GENERALES", ["GENERAL", "GENERALES"]),
        ],
        "section": [
            ("APROBACION MUNICIPAL", ["ALCALDIA", "APROBACION MUNICIPAL", "MUNICIPAL"]),
            ("GENERALES", ["EDIFICIO COMPLETO", "GENERAL", "GENERALES"]),
            ("DE BORDE", ["CORTE DE BORDE", "BORDE"]),
            ("DE DETALLE", ["DETALLE", "DETALLES"]),
            ("DE ESCALERA", ["ESCALERA", "ESCALERAS"]),
            ("CIMENTACION", ["CIMENTACION"]),
            ("LOSAS", ["LOSA", "LOSAS"]),
            ("VIGAS", ["VIGA", "VIGAS"]),
            ("DUCTOS", ["DUCTO", "DUCTOS"]),
            ("EQUIPOS", ["EQUIPO", "EQUIPOS"]),
            ("CUARTOS TECNICOS", ["CUARTO TECNICO", "CUARTOS TECNICOS"]),
            ("CANALIZACIONES", ["CANALIZACION", "CANALIZACIONES"]),
        ],
        "floorplan": [
            ("WORKING", ["WORKING"]),
            ("GENERALES", ["PLANOS", "GENERAL", "GENERALES"]),
            ("ACOTADOS", ["ACOTADO", "ACOTADOS"]),
            ("APROBACION MUNICIPAL", ["ALCALDIA", "APROBACION MUNICIPAL", "MUNICIPAL"]),
            ("BORDE DE LOSA", ["BORDE DE LOSA", "BORDE DE LA LOSA", "BORDE"]),
            ("AMOBLADOS", ["AMOBLADO", "AMOBLADOS"]),
            ("CUBIERTA", ["CUBIERTA"]),
            ("AMPLIADOS", ["AMPLIADO", "AMPLIADOS"]),
            ("ACABADOS", ["ACABADO", "ACABADOS", "ACABADOS FINALES"]),
            ("SENALETICA", ["SENALETICA", "SENALITICA"]),
            ("AGUA FRIA", ["AGUA FRIA"]),
            ("AGUA CALIENTE", ["AGUA CALIENTE"]),
            ("RETORNO AGUA CALIENTE", ["RETORNO", "AGUA CALIENTE RETORNO"]),
            ("SANITARIO", ["SANITARIO"]),
            ("PLUVIAL", ["PLUVIAL"]),
            ("PROTECCION CONTRA INCENDIO", ["PROTECCION CONTRA INCENDIO", "PCI", "INCENDIO"]),
            ("TANQUES", ["TANQUE", "TANQUES"]),
            ("ISOMETRICOS", ["ISOMETRICO", "ISOMETRICOS"]),
            ("ILUMINACION", ["ILUMINACION"]),
            ("TOMACORRIENTES", ["TOMACORRIENTE", "TOMACORRIENTES"]),
            ("FUERZA", ["FUERZA"]),
            ("VOZ Y DATA", ["VOZ Y DATA"]),
            ("SISTEMAS ESPECIALES", ["SISTEMA ESPECIAL", "SISTEMAS ESPECIALES"]),
            ("ALARMAS", ["ALARMA", "ALARMAS"]),
            ("PUESTA A TIERRA", ["TIERRA", "PUESTA A TIERRA"]),
            ("HVAC", ["HVAC"]),
            ("SUMINISTRO DE AIRE", ["SUMINISTRO"]),
            ("RETORNO DE AIRE", ["RETORNO DE AIRE"]),
            ("EXTRACCION", ["EXTRACCION"]),
            ("EQUIPOS", ["EQUIPO", "EQUIPOS"]),
            ("CONTROLES", ["CONTROL", "CONTROLES"]),
            ("CIMENTACION", ["CIMENTACION"]),
            ("LOSAS", ["LOSA", "LOSAS"]),
            ("VIGAS", ["VIGA", "VIGAS"]),
            ("COLUMNAS", ["COLUMNA", "COLUMNAS"]),
            ("REFUERZO", ["REFUERZO"]),
        ],
        "structuralplan": [
            ("WORKING", ["WORKING"]),
            ("GENERALES", ["PLANOS", "GENERAL", "GENERALES"]),
            ("ACOTADOS", ["ACOTADO", "ACOTADOS"]),
            ("APROBACION MUNICIPAL", ["ALCALDIA", "APROBACION MUNICIPAL", "MUNICIPAL"]),
            ("CIMENTACION", ["CIMENTACION"]),
            ("LOSAS", ["LOSA", "LOSAS"]),
            ("VIGAS", ["VIGA", "VIGAS"]),
            ("COLUMNAS", ["COLUMNA", "COLUMNAS"]),
            ("REFUERZO", ["REFUERZO"]),
        ],
        "ceilingplan": [
            ("CIELO REFLEJADO - AMPLIADOS", ["AMPLIADO", "AMPLIADOS"]),
            ("CIELO REFLEJADO", ["CIELO REFLEJADO", "REFLEJADO"]),
            ("ILUMINACION", ["ILUMINACION"]),
        ],
        "detail": [
            ("GENERALES", ["DETAIL", "DETALLE", "DETALLES"]),
            ("PLANTA", ["PLANTA"]),
            ("DE MOBILIARIO", ["MUEBLE", "MUEBLES", "MOBILIARIO"]),
            ("ESPECIALES", ["ESPECIAL", "A 06 DETAIL"]),
            ("REFUERZO", ["REFUERZO"]),
            ("CONEXIONES", ["CONEXION", "CONEXIONES"]),
            ("SANITARIO", ["SANITARIO"]),
            ("AGUA FRIA", ["AGUA FRIA"]),
            ("AGUA CALIENTE", ["AGUA CALIENTE"]),
            ("PLUVIAL", ["PLUVIAL"]),
            ("PROTECCION CONTRA INCENDIO", ["PCI", "INCENDIO"]),
            ("ILUMINACION", ["ILUMINACION"]),
            ("TABLEROS", ["TABLERO", "TABLEROS"]),
            ("CANALIZACIONES", ["CANALIZACION", "CANALIZACIONES"]),
            ("PUESTA A TIERRA", ["TIERRA"]),
            ("DUCTOS", ["DUCTO", "DUCTOS"]),
            ("EQUIPOS", ["EQUIPO", "EQUIPOS"]),
            ("SOPORTES", ["SOPORTE", "SOPORTES"]),
        ],
        "drafting": [
            ("GENERALES", ["DETAIL", "DETALLE", "DETALLES"]),
            ("PLANTA", ["PLANTA"]),
            ("DE MOBILIARIO", ["MUEBLE", "MUEBLES", "MOBILIARIO"]),
            ("ESPECIALES", ["ESPECIAL", "A 06 DETAIL"]),
        ],
        "threedimensional": [
            ("GENERAL", ["3D VIEW", "GENERAL"]),
            ("ESTRUCTURAL", ["STRUCTURAL"]),
            ("PLOMERIA", ["PLUMBING"]),
            ("ELECTRICO", ["ELECTRICAL"]),
            ("MECANICO", ["MECHANICAL"]),
        ],
    }

    family_rules = keyword_sets.get(family_key, [])
    for subcat, needles in family_rules:
        for needle in needles:
            if needle in text:
                return subcat

    num, category, subcat = parse_numbered_standard_name(text)
    if subcat:
        return normalize_subcategory_from_text(subcat)

    # Secondary fallback: if the family category word exists inside the
    # current name, keep everything after it as the inferred subcategory.
    family_category = family_category_label(view_family_name)
    if family_category and family_category in text:
        try:
            tail = text.split(family_category, 1)[1].strip()
            tail = tail.lstrip('-').strip()
            tail = normalize_subcategory_from_text(tail)
            if tail:
                return tail
        except:
            pass

    return ""


def build_standard_subcategory_numbers(rule_rows):
    numbers = {}
    for row in rule_rows:
        family_key = normalize_key(row.ViewFamilyName)
        num, category, subcat = parse_numbered_standard_name(row.NewTypeName)
        if not num or not subcat:
            continue
        subcat_key = normalize_key(normalize_subcategory_from_text(subcat))
        numbers[(family_key, subcat_key)] = num
    return numbers


def collect_used_numbers(rule_rows, project_rows):
    used = {}

    def add_name(family_name, name):
        family_key = normalize_key(family_name)
        num, category, subcat = parse_numbered_standard_name(name)
        if not num:
            return
        if family_key not in used:
            used[family_key] = set()
        try:
            used[family_key].add(int(num))
        except:
            pass

    for row in rule_rows:
        add_name(row.ViewFamilyName, row.NewTypeName)

    for row in project_rows:
        add_name(row.ViewFamilyName, row.CurrentTypeName)
        add_name(row.ViewFamilyName, row.NewTypeName)

    return used


def next_free_number(used_numbers, family_name):
    family_key = normalize_key(family_name)
    values = used_numbers.get(family_key, set())
    candidate = 1
    while candidate in values:
        candidate += 1
    return "%02d" % candidate


def build_suggested_name(family_name, subcategory, number_string):
    category = family_category_label(family_name)
    if not category or not subcategory:
        return ""
    return "{} {} - {}".format(number_string, category, subcategory)


def is_structured_name_valid_for_family(family_name, current_name):
    text = normalize_display_name(current_name)
    if not text:
        return False
    if text == "00 WORKING":
        return True

    num, category, subcat = parse_numbered_standard_name(text)
    family_category = family_category_label(family_name)

    if num and family_category and category == family_category and subcat:
        return True

    # Ceiling plans still use PLANOS - ... in this standard.
    if normalize_key(family_name) == "ceilingplan" and num and category == "PLANOS" and subcat:
        return True

    return False


def apply_rules_to_project(rule_rows, project_rows):
    lookup = build_lookup(rule_rows)
    subcat_standard_numbers = build_standard_subcategory_numbers(rule_rows)
    used_numbers = collect_used_numbers(rule_rows, project_rows)

    matched = 0
    normalized = 0
    numbered = 0
    unmatched = 0

    for row in project_rows:
        row.Include = False
        row.AddRule = False
        row.NewTypeName = ""

        current_normalized = normalize_display_name(row.CurrentTypeName)
        key = (normalize_key(row.ViewFamilyName), normalize_key(current_normalized))

        if key in lookup:
            row.NewTypeName = lookup[key]
            row.Include = True
            if normalize_key(current_normalized) == normalize_key(normalize_display_name(row.NewTypeName)):
                row.set_status("NO CHANGE")
            else:
                row.set_status("MATCHED")
            matched += 1
            continue

        # Respect names that already look properly standardized for the family.
        if is_structured_name_valid_for_family(row.ViewFamilyName, row.CurrentTypeName):
            row.NewTypeName = current_normalized
            row.set_status("NORMALIZED")
            normalized += 1
            continue

        subcat = infer_subcategory(row.ViewFamilyName, row.CurrentTypeName)
        family_key = normalize_key(row.ViewFamilyName)

        if subcat:
            subcat_key = normalize_key(subcat)
            standard_num = subcat_standard_numbers.get((family_key, subcat_key))
            if standard_num:
                candidate_name = build_suggested_name(row.ViewFamilyName, subcat, standard_num)
                if candidate_name and normalize_key(candidate_name) != normalize_key(current_normalized):
                    row.NewTypeName = candidate_name
                    row.set_status("NUMBERED")
                    numbered += 1
                    continue
                elif current_normalized:
                    row.NewTypeName = current_normalized
                    row.set_status("NORMALIZED")
                    normalized += 1
                    continue

            next_num = next_free_number(used_numbers, row.ViewFamilyName)
            candidate_name = build_suggested_name(row.ViewFamilyName, subcat, next_num)
            if candidate_name and normalize_key(candidate_name) != normalize_key(current_normalized):
                row.NewTypeName = candidate_name
                try:
                    used_numbers.setdefault(family_key, set()).add(int(next_num))
                except:
                    pass
                row.set_status("NUMBERED")
                numbered += 1
                continue

        if current_normalized:
            row.NewTypeName = current_normalized
            row.set_status("NORMALIZED")
            normalized += 1
        else:
            row.NewTypeName = ""
            row.set_status("UNMATCHED")
            unmatched += 1

    return matched, normalized, numbered, unmatched

# ==================================================
# RENAME
# ==================================================

def build_existing_name_lookup(project_rows):
    lookup = {}
    for row in project_rows:
        lookup[(normalize_key(row.ViewFamilyName), normalize_key(row.CurrentTypeName))] = row.ElementId
    return lookup


def get_duplicate_targets(project_rows):
    counts = {}
    duplicates = set()

    for row in project_rows:
        if not row.Include:
            continue
        new_name = normalize_key(row.NewTypeName)
        family = normalize_key(row.ViewFamilyName)
        if not new_name:
            continue
        key = (family, new_name)
        counts[key] = counts.get(key, 0) + 1

    for key, count in counts.items():
        if count > 1:
            duplicates.add(key)
    return duplicates


def rename_selected_view_types(project_rows):
    renamed = []
    skipped = []
    errors = []

    existing_lookup = build_existing_name_lookup(project_rows)
    duplicate_targets = get_duplicate_targets(project_rows)

    for row in project_rows:
        if not row.Include:
            continue

        family_name = normalize_text(row.ViewFamilyName)
        current_name = normalize_text(row.CurrentTypeName)
        new_name = normalize_text(row.NewTypeName)

        family_key = normalize_key(family_name)
        current_key = normalize_key(current_name)
        new_key = normalize_key(new_name)

        if is_protected_name(current_name):
            row.set_status("SKIPPED")
            skipped.append({"family": family_name, "old": current_name, "new": new_name, "reason": "Protected/system name"})
            continue

        if not new_name:
            row.set_status("SKIPPED")
            skipped.append({"family": family_name, "old": current_name, "new": new_name, "reason": "Empty new name"})
            continue

        if is_protected_name(new_name):
            row.set_status("SKIPPED")
            skipped.append({"family": family_name, "old": current_name, "new": new_name, "reason": "Invalid target"})
            continue

        if current_key == new_key:
            row.set_status("NO CHANGE")
            skipped.append({"family": family_name, "old": current_name, "new": new_name, "reason": "Name unchanged"})
            continue

        if (family_key, new_key) in duplicate_targets:
            row.set_status("DUPLICATE TARGET")
            skipped.append({"family": family_name, "old": current_name, "new": new_name, "reason": "Duplicate target"})
            continue

        if (family_key, new_key) in existing_lookup and existing_lookup[(family_key, new_key)] != row.ElementId:
            row.set_status("ALREADY EXISTS")
            skipped.append({"family": family_name, "old": current_name, "new": new_name, "reason": "Name already exists"})
            continue

        t = DB.Transaction(doc, "Rename View Type")
        try:
            t.Start()
            row.Element.Name = new_name
            t.Commit()
            row.CurrentTypeName = new_name
            row.set_status("RENAMED")
            renamed.append({"family": family_name, "old": current_name, "new": new_name})
            existing_lookup.pop((family_key, current_key), None)
            existing_lookup[(family_key, new_key)] = row.ElementId
        except Exception as ex:
            try:
                if t.HasStarted():
                    t.RollBack()
            except:
                pass
            row.set_status("FAILED")
            errors.append({"family": family_name, "old": current_name, "new": new_name, "reason": first_line_error(ex)})

    return renamed, skipped, errors


# ==================================================
# WINDOW
# ==================================================

class ViewTypeStandardizerWindow(forms.WPFWindow):
    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)

        self.standard_rows = ObservableCollection[object]()
        self.standard_rows_master = []
        self.project_rows_view = ObservableCollection[object]()
        self.project_rows_master = []
        self.discipline_map = {}
        self.txt_errors = []
        self.scan_errors = []
        self.project_family_names = []
        self.discipline_display_map = {}
        self.discipline_reverse_display_map = {}
        self.loaded_discipline_name = None

        self.StandardGrid.ItemsSource = self.standard_rows
        self.ProjectGrid.ItemsSource = self.project_rows_view

        load_logo_if_available(self)

        self._wire_events()
        self._load_project_rows()
        self._reload_txt_and_combo()
        self._load_default_discipline()
        self._load_filter_values(status_default="UNMATCHED")
        self._set_standard_family_column_items()
        self._apply_project_filters()
        self._refresh_summary()

    def _wire_events(self):
        self.LoadDisciplineButton.Click += self.on_load_discipline
        self.DisciplineCombo.SelectionChanged += self.on_discipline_changed
        self.ReloadTxtButton.Click += self.on_reload_txt
        self.SaveTxtButton.Click += self.on_save_txt
        self.AddRuleButton.Click += self.on_add_rule
        self.RemoveRuleButton.Click += self.on_remove_rule
        self.MatchRulesButton.Click += self.on_match_rules
        self.RefreshProjectButton.Click += self.on_refresh_project
        self.AddCheckedToRulesButton.Click += self.on_add_checked_to_rules
        self.ApplyButton.Click += self.on_apply
        self.CloseButton.Click += self.on_close
        self.ProjectFamilyFilter.SelectionChanged += self.on_filter_changed
        self.ProjectStatusFilter.SelectionChanged += self.on_filter_changed
        self.ProjectSearchBox.TextChanged += self.on_filter_changed
        self.CheckVisibleAddRuleButton.Click += self.on_check_visible_add_rule
        self.CheckVisibleUseButton.Click += self.on_check_visible_use
        self.UncheckVisibleButton.Click += self.on_uncheck_visible

    def _actual_to_display_discipline(self, name):
        return normalize_text(name).upper()

    def _display_to_actual_discipline(self, value):
        value = normalize_text(value)
        if value in self.discipline_reverse_display_map:
            return self.discipline_reverse_display_map[value]
        return value

    def _get_selected_discipline_actual(self):
        display_value = get_combobox_text(self.DisciplineCombo)
        return self._display_to_actual_discipline(display_value)

    def _set_selected_discipline_actual(self, actual_value):
        display_value = self._actual_to_display_discipline(actual_value)
        try:
            self.DisciplineCombo.SelectedItem = display_value
        except:
            pass

    def _load_project_rows(self):
        rows, errors = collect_project_view_types()
        self.project_rows_master = rows
        self.scan_errors = errors
        found = set()
        for row in rows:
            found.add(normalize_text(row.ViewFamilyName))
        self.project_family_names = sort_alpha(list(found))

    def _reload_txt_and_combo(self):
        self.discipline_map, self.txt_errors = read_txt_standards()
        self.discipline_display_map = {}
        self.discipline_reverse_display_map = {}

        display_items = []
        for actual_name in sort_alpha(self.discipline_map.keys()):
            display_name = self._actual_to_display_discipline(actual_name)
            self.discipline_display_map[actual_name] = display_name
            self.discipline_reverse_display_map[display_name] = actual_name
            display_items.append(display_name)

        self.DisciplineCombo.ItemsSource = display_items
        if self.DisciplineCombo.Items.Count > 0 and self.DisciplineCombo.SelectedIndex < 0:
            self.DisciplineCombo.SelectedIndex = 0

    def _load_default_discipline(self):
        if "Architecture" in self.discipline_map:
            self._set_selected_discipline_actual("Architecture")
            self.load_discipline_into_grid("Architecture")
            self.match_rules_to_project(show_alert=False)
        elif len(self.discipline_map.keys()) > 0:
            first_name = sort_alpha(self.discipline_map.keys())[0]
            self._set_selected_discipline_actual(first_name)
            self.load_discipline_into_grid(first_name)
            self.match_rules_to_project(show_alert=False)

    def _load_filter_values(self, status_default="UNMATCHED"):
        current_family = get_combobox_text(self.ProjectFamilyFilter) or "ALL"
        current_status = get_combobox_text(self.ProjectStatusFilter) or status_default

        families = ["ALL"] + list(self.project_family_names)
        statuses = ["ALL", "UNMATCHED", "NORMALIZED", "NUMBERED", "MATCHED", "NO CHANGE", "RENAMED"]

        self.ProjectFamilyFilter.ItemsSource = families
        self.ProjectStatusFilter.ItemsSource = statuses

        if current_family in families:
            self.ProjectFamilyFilter.SelectedItem = current_family
        elif self.ProjectFamilyFilter.Items.Count > 0:
            self.ProjectFamilyFilter.SelectedItem = "ALL"

        if current_status in statuses:
            self.ProjectStatusFilter.SelectedItem = current_status
        elif status_default in statuses:
            self.ProjectStatusFilter.SelectedItem = status_default
        else:
            self.ProjectStatusFilter.SelectedItem = "ALL"

    def _set_standard_family_column_items(self):
        set_combo_items(self.StandardFamilyColumn, self.project_family_names)

    def _sort_standard_rows_master(self):
        def standard_sort_key(row):
            num, category, subcat = parse_numbered_standard_name(row.NewTypeName)
            order_num = num if num is not None else 999
            return (normalize_key(row.ViewFamilyName), order_num, normalize_key(row.NewTypeName), normalize_key(row.OldTypeName))

        self.standard_rows_master = sorted(self.standard_rows_master, key=standard_sort_key)

    def _apply_standard_filters(self):
        family_filter = get_combobox_text(self.ProjectFamilyFilter)

        clear_collection(self.standard_rows)
        self._sort_standard_rows_master()

        visible_count = 0
        for row in self.standard_rows_master:
            if family_filter and family_filter != "ALL":
                if normalize_key(row.ViewFamilyName) != normalize_key(family_filter):
                    continue
            self.standard_rows.Add(row)
            visible_count += 1

        discipline_name = self._get_selected_discipline_actual() or get_combobox_text(self.DisciplineCombo)
        if family_filter and family_filter != "ALL":
            if discipline_name:
                self.StandardRulesCountText.Text = "Showing {} {} rules for {}".format(visible_count, discipline_name, family_filter)
            else:
                self.StandardRulesCountText.Text = "Showing {} rules for {}".format(visible_count, family_filter)
        else:
            if discipline_name:
                self.StandardRulesCountText.Text = "Showing {} {} rules".format(visible_count, discipline_name)
            else:
                self.StandardRulesCountText.Text = "Showing {} rules".format(visible_count)

    def _apply_project_filters(self):
        family_filter = get_combobox_text(self.ProjectFamilyFilter)
        status_filter = get_combobox_text(self.ProjectStatusFilter)
        search_text = normalize_key(self.ProjectSearchBox.Text)

        clear_collection(self.project_rows_view)
        visible_count = 0
        total_count = len(self.project_rows_master)

        for row in self.project_rows_master:
            if family_filter and family_filter != "ALL":
                if normalize_key(row.ViewFamilyName) != normalize_key(family_filter):
                    continue
            if status_filter and status_filter != "ALL":
                if normalize_key(row.Status) != normalize_key(status_filter):
                    continue
            if search_text:
                haystack = "{} {} {}".format(row.CurrentTypeName, row.NewTypeName, row.ViewFamilyName)
                if search_text not in normalize_key(haystack):
                    continue
            self.project_rows_view.Add(row)
            visible_count += 1

        self.ProjectTypesCountText.Text = "Showing {} of {} project types".format(visible_count, total_count)
        self._apply_standard_filters()

    def _refresh_summary(self):
        total_rules = self.standard_rows.Count
        total_types = len(self.project_rows_master)
        selected_count = 0
        matched_count = 0
        renamed_count = 0

        for row in self.project_rows_master:
            if row.Include:
                selected_count += 1
            if row.Status == "MATCHED":
                matched_count += 1
            if row.Status == "RENAMED":
                renamed_count += 1

        self.TotalRulesText.Text = safe_str(total_rules)
        self.TotalTypesText.Text = safe_str(total_types)
        self.SelectedTypesText.Text = safe_str(selected_count)
        self.MatchedTypesText.Text = safe_str(matched_count)
        self.RenamedTypesText.Text = safe_str(renamed_count)

    def load_discipline_into_grid(self, discipline_name):
        discipline_name = normalize_text(discipline_name)
        self.loaded_discipline_name = discipline_name
        self.standard_rows_master = []
        rows = self.discipline_map.get(discipline_name, [])
        for row in rows:
            self.standard_rows_master.append(StandardRuleRow(row.Discipline, row.ViewFamilyName, row.OldTypeName, row.NewTypeName))
        self._set_selected_discipline_actual(discipline_name)
        self._apply_standard_filters()
        self._refresh_summary()
        refresh_grid(self.StandardGrid)

    def save_current_grid_to_memory(self):
        discipline_name = self.loaded_discipline_name or self._get_selected_discipline_actual()
        if not discipline_name:
            forms.alert("Select a discipline first.")
            return False

        rows = []
        self._sort_standard_rows_master()
        for row in self.standard_rows_master:
            family = normalize_text(row.ViewFamilyName)
            old_name = normalize_text(row.OldTypeName)
            new_name = normalize_text(row.NewTypeName)
            if not family and not old_name and not new_name:
                continue
            if not family or not old_name or not new_name:
                continue
            rows.append(StandardRuleRow(discipline_name, family, old_name, new_name))

        self.discipline_map[discipline_name] = rows
        self._reload_txt_and_combo()
        self._set_selected_discipline_actual(discipline_name)
        self._apply_standard_filters()
        return True

    def rule_exists(self, discipline_name, family_name, old_name):
        rows = self.discipline_map.get(discipline_name, [])
        for row in rows:
            if normalize_key(row.ViewFamilyName) == normalize_key(family_name) and normalize_key(row.OldTypeName) == normalize_key(old_name):
                return True
        return False

    def match_rules_to_project(self, show_alert=True):
        rules = [row for row in self.standard_rows_master]
        matched, normalized, numbered, unmatched = apply_rules_to_project(rules, self.project_rows_master)
        self._load_filter_values(status_default="UNMATCHED")
        self._apply_project_filters()
        self._refresh_summary()

        if show_alert:
            forms.alert(
                "Analysis completed.\n\nMatched: {}\nNormalized: {}\nNumbered: {}\nUnmatched: {}".format(matched, normalized, numbered, unmatched),
                title="View Type Standardizer"
            )

    def print_report(self, renamed, skipped, errors):
        output.print_md("# MENVIC | VIEW TYPE STANDARDIZER")
        output.print_md("")
        output.print_md("## Summary")
        output.print_md("")
        output.print_md("* Discipline: {}".format(self._get_selected_discipline_actual()))
        output.print_md("* Standards file: {}".format(TXT_FILE))
        output.print_md("* Total rules: {}".format(self.standard_rows.Count))
        output.print_md("* Total project view types: {}".format(len(self.project_rows_master)))
        output.print_md("* Renamed: {}".format(len(renamed)))
        output.print_md("* Skipped: {}".format(len(skipped)))
        output.print_md("* Failures: {}".format(len(errors)))
        output.print_md("")

        if renamed:
            output.print_md("## Renamed")
            output.print_md("")
            output.print_md("| View Family | Old Name | New Name |")
            output.print_md("|---|---|---|")
            for item in renamed:
                output.print_md("| {} | {} | {} |".format(item["family"], item["old"], item["new"]))
            output.print_md("")

    # EVENTS

    def on_load_discipline(self, sender, args):
        discipline_name = self._get_selected_discipline_actual()
        if not discipline_name:
            forms.alert("Select a discipline first.")
            return

        if discipline_name not in self.discipline_map:
            forms.alert("Discipline was not found in TXT.")
            return

        self.load_discipline_into_grid(discipline_name)

    def on_reload_txt(self, sender, args):
        current_name = self._get_selected_discipline_actual()
        self._reload_txt_and_combo()
        if current_name and current_name in self.discipline_map:
            self._set_selected_discipline_actual(current_name)
            self.load_discipline_into_grid(current_name)
        self._refresh_summary()
        forms.alert("TXT reloaded.")

    def on_save_txt(self, sender, args):
        if not self.save_current_grid_to_memory():
            return
        try:
            write_txt_standards(self.discipline_map)
            forms.alert("TXT saved successfully.")
        except Exception as ex:
            forms.alert("Failed to save TXT.\n\n{}".format(first_line_error(ex)))

    def on_add_rule(self, sender, args):
        discipline_name = self._get_selected_discipline_actual() or "Architecture"
        default_family = self.project_family_names[0] if self.project_family_names else "FloorPlan"
        old_name = ""
        new_name = ""

        try:
            selected_project_row = self.ProjectGrid.SelectedItem
        except:
            selected_project_row = None

        if selected_project_row is not None:
            default_family = normalize_text(selected_project_row.ViewFamilyName) or default_family
            old_name = normalize_text(selected_project_row.CurrentTypeName)
            new_name = normalize_text(selected_project_row.NewTypeName)

        new_row = StandardRuleRow(discipline_name, default_family, old_name, new_name)
        self.standard_rows_master.append(new_row)
        self._apply_standard_filters()
        self.StandardGrid.ScrollIntoView(new_row)
        self._refresh_summary()

    def on_remove_rule(self, sender, args):
        selected_items = self.StandardGrid.SelectedItems
        if selected_items is None or selected_items.Count == 0:
            forms.alert("Select one or more rule rows to remove.")
            return

        to_remove = [item for item in selected_items]
        for item in to_remove:
            try:
                self.standard_rows_master.remove(item)
            except:
                pass
        self._apply_standard_filters()
        self._refresh_summary()

    def on_match_rules(self, sender, args):
        self.match_rules_to_project(show_alert=True)

    def on_refresh_project(self, sender, args):
        current_status = get_combobox_text(self.ProjectStatusFilter) or "ALL"
        self._load_project_rows()
        self._load_filter_values(status_default=current_status)
        self._set_standard_family_column_items()
        self._apply_project_filters()
        self._refresh_summary()
        forms.alert("Project View Types refreshed.")

    def on_add_checked_to_rules(self, sender, args):
        discipline_name = self._get_selected_discipline_actual()
        if not discipline_name:
            forms.alert("Select a discipline first.")
            return

        added = 0
        skipped = 0
        for row in self.project_rows_master:
            if not row.AddRule:
                continue

            family = normalize_text(row.ViewFamilyName)
            old_name = normalize_text(row.CurrentTypeName)
            new_name = normalize_text(row.NewTypeName)

            if not new_name:
                skipped += 1
                continue

            if self.rule_exists(discipline_name, family, old_name):
                skipped += 1
                row.AddRule = False
                continue

            self.standard_rows_master.append(StandardRuleRow(discipline_name, family, old_name, new_name))
            row.AddRule = False
            added += 1

        self._apply_standard_filters()
        self.save_current_grid_to_memory()
        refresh_grid(self.StandardGrid)
        refresh_grid(self.ProjectGrid)
        self._refresh_summary()
        forms.alert("Rules added: {}\nSkipped: {}".format(added, skipped), title="View Type Standardizer")

    def on_apply(self, sender, args):
        renamed, skipped, errors = rename_selected_view_types(self.project_rows_master)
        self._apply_project_filters()
        self._refresh_summary()
        self.print_report(renamed, skipped, errors)
        forms.alert(
            "Apply finished.\n\nRenamed: {}\nSkipped: {}\nErrors: {}".format(len(renamed), len(skipped), len(errors)),
            title="View Type Standardizer"
        )

    def on_discipline_changed(self, sender, args):
        discipline_name = self._get_selected_discipline_actual() or get_combobox_text(self.DisciplineCombo)
        if not discipline_name:
            return
        if discipline_name not in self.discipline_map:
            return
        self.load_discipline_into_grid(discipline_name)
        self._apply_project_filters()
        self._refresh_summary()

    def on_filter_changed(self, sender, args):
        self._apply_project_filters()

    def on_check_visible_add_rule(self, sender, args):
        for row in self.project_rows_view:
            row.AddRule = True
        refresh_grid(self.ProjectGrid)
        self._refresh_summary()

    def on_check_visible_use(self, sender, args):
        for row in self.project_rows_view:
            row.Include = True
        refresh_grid(self.ProjectGrid)
        self._refresh_summary()

    def on_uncheck_visible(self, sender, args):
        for row in self.project_rows_view:
            row.AddRule = False
            row.Include = False
        refresh_grid(self.ProjectGrid)
        self._refresh_summary()

    def on_close(self, sender, args):
        self.Close()


# ==================================================
# CLEANUP
# ==================================================

def main():
    if not os.path.exists(XAML_FILE):
        forms.alert("ui.xaml was not found next to script.py", exitscript=True)
        return

    window = ViewTypeStandardizerWindow(XAML_FILE)
    window.ShowDialog()


if __name__ == "__main__":
    main()

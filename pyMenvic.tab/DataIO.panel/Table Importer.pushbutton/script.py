# -*- coding: utf-8 -*-
__title__ = "Table Importer"
__author__ = "Menvic"

import os
import sys
import shutil
import re

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System")
clr.AddReference("RevitAPI")

from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Window, Visibility, MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Forms import OpenFileDialog, DialogResult
import System
from System import Uri, Int64, Int32, Convert
from System.Windows.Controls import CheckBox, DataGridRow, ListBoxItem, StackPanel, TextBlock, DataGridEditingUnit
from System.Windows.Input import MouseButtonEventHandler
from System.Windows.Media import VisualTreeHelper
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Line,
    TextNote,
    TextNoteType,
    Transaction,
    ViewDrafting,
    ViewFamily,
    ViewFamilyType,
    XYZ,
)

from pyrevit import script

script_dir = os.path.dirname(__file__)
if script_dir not in sys.path:
    sys.path.append(script_dir)
importers_dir = os.path.join(script_dir, "importers")
if importers_dir not in sys.path:
    sys.path.append(importers_dir)

TOOL_VERSION = "MVP 0.4.2"
TOOL_VERSION_LABEL = "MVP 0.4.2"
DEBUG_OUTPUT = False
USE_WRAPPED_TEXT_NOTES = False

from importers import import_row_to_revit
from models import TableEntry
from storage import load_entries, save_entries, get_storage_path
from excel_reader import (
    get_excel_worksheets,
    get_last_modified,
    get_file_name_without_extension,
    get_excel_regions,
    clean_hidden_unicode,
    read_excel_table_values,
)

USED_RANGE_DISPLAY = u"Full Worksheet Used Range"
DEFAULT_IMPORT_TYPES = ["Excel Link"]
DEFAULT_VIEW_TYPES = ["Drafting View", "Legend View (Not implemented)", "Schedule View"]
DEFAULT_SCALES = ["1", "2", "5", "10", "20", "25", "50", "75", "100"]
DEFAULT_DPI = ["72", "96", "150", "200", "300", "600"]
ROW_READY_STATUS = "Ready to Update"
EXCEL_FILTER = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
MAX_TABLE_ROWS = 400
MAX_TABLE_COLUMNS = 60
TABLE_COLUMN_WIDTH = 0.8
TABLE_ROW_HEIGHT = 0.25
TABLE_SCALE = 1.0
EXCEL_WIDTH_SCALE = 0.064
EXCEL_ROW_HEIGHT_SCALE = 1.0
TABLE_TEXT_SIZE_MM = 1.6
TEXT_SIZE_SCALE = 1.0
MIN_TEXT_SIZE_FT = TABLE_TEXT_SIZE_MM / 304.8
MAX_TEXT_SIZE_FT = 0.012
APPROX_TEXT_CHAR_WIDTH_FACTOR = 1.60
MAX_CELL_TEXT_CHARS = 200
MIN_COL_WIDTH_FT = 0.25
MAX_COL_WIDTH_FT = 6.80
ADAPTIVE_MIN_TABLE_WIDTH_FT = 8.0
ADAPTIVE_MAX_TABLE_WIDTH_FT = 30.0
ADAPTIVE_TARGET_WIDTH_PER_COLUMN_FT = 1.05
ADAPTIVE_CONTENT_WEIGHT = 0.58
ADAPTIVE_EXCEL_WEIGHT = 0.42
MIN_ROW_HEIGHT_FT = 0.16
MAX_ROW_HEIGHT_FT = 0.48
MIN_BODY_ROW_HEIGHT_FT = 0.18
BODY_ROW_HEIGHT_SCALE = 1.12
TEXT_PADDING_X = 0.028
TEXT_PADDING_Y = 0.024
APPROX_CHARS_PER_FOOT = 18
ROW_HEIGHT_SCALE = EXCEL_ROW_HEIGHT_SCALE
TABLE_TEXT_OFFSET_X = TEXT_PADDING_X
TABLE_TEXT_OFFSET_Y = TEXT_PADDING_Y
TABLE_MIN_COLUMN_WIDTH = MIN_COL_WIDTH_FT
TABLE_MAX_COLUMN_WIDTH = MAX_COL_WIDTH_FT
TABLE_CHAR_WIDTH = 0.043
TABLE_MIN_ROW_HEIGHT = MIN_ROW_HEIGHT_FT
TABLE_MAX_ROW_HEIGHT = MAX_ROW_HEIGHT_FT
TABLE_LINE_HEIGHT = 0.145
TABLE_TEXT_TYPE_NAME = "Arial 1.60mm"
TABLE_TEXT_SIZE = TABLE_TEXT_SIZE_MM / 304.8
FALLBACK_TABLE_FONT = "Arial"
TEXT_NOTE_WIDTH_WARNING_PRINTED = False
TABLE_IMPORTER_TOOL_NAME = "pyMENVIC_TABLE_IMPORTER"
TABLE_IMPORTER_CREATED_BY = "Table Importer"
TABLE_IMPORTER_TAG_PREFIX = "pyMENVIC_TABLE_IMPORTER|"
STANDARD_TABLE_TEXT_TYPE_NAME = "Arial 1.60mm"
STANDARD_TABLE_TEXT_SIZE_MM = 1.60
STANDARD_TEXT_SIZE_TOL_MM = 0.06
TABLE_TEXT_STYLE_WARNING = False
TEXT_TYPE_CACHE = {}
EXCEL_POINT_TO_MM = 0.3527777778
MIN_EXCEL_TEXT_SIZE_MM = 1.20
MAX_EXCEL_TEXT_SIZE_MM = 3.20


def safe_unicode(value):
    if value is None:
        return u""
    try:
        if isinstance(value, unicode):
            return value
    except Exception:
        pass
    try:
        if isinstance(value, str):
            for enc in ("utf-8", "cp1252", "latin-1"):
                try:
                    return value.decode(enc)
                except Exception:
                    pass
            return value.decode("utf-8", "replace")
    except Exception:
        pass
    try:
        return unicode(value)
    except Exception:
        try:
            return unicode(value.ToString())
        except Exception:
            return u""


def clean_display_text(value):
    return clean_hidden_unicode(safe_unicode(value))


def normalize_row_status(value):
    status = clean_display_text(value).strip()
    if status in ("To Update", "Ready", "Not Created", "OK", "Created", "Legacy Reset"):
        return ROW_READY_STATUS
    if status in ("Error",):
        return "Failed"
    if status in (ROW_READY_STATUS, "Updating...", "Updated", "Skipped", "Failed", "Missing File", "Invalid Region"):
        return status
    if not status:
        return ROW_READY_STATUS
    return status


def is_ready_status(value):
    return normalize_row_status(value) == ROW_READY_STATUS


def normalize_status_label(text):
    label = clean_display_text(text).strip()
    while label.lower().startswith("status:"):
        label = label[7:].strip()
    if not label:
        label = "Ready"
    return label


def normalize_region_name_for_view(region_display, worksheet_name):
    region = clean_display_text(region_display).strip()
    worksheet = clean_display_text(worksheet_name).strip()
    lowered = region.lower()
    if not region or lowered == USED_RANGE_DISPLAY.lower():
        return worksheet or u"Used Range"

    for prefix in (u"Name ", u"Table ", u"Print Area "):
        if region.lower().startswith(prefix.lower()):
            region = region[len(prefix):].strip()
            break

    try:
        match = re.search(r"\s*:\s*\$?[A-Za-z]{1,3}\$?\d+\s*:\s*\$?[A-Za-z]{1,3}\$?\d+\s*$", region)
        if match:
            region = region[:match.start()].strip()
    except Exception:
        pass

    if not region:
        region = worksheet or u"Used Range"
    return region


def sanitize_revit_view_name(value):
    name = clean_display_text(value).strip()
    cleaned = []
    invalid_chars = set(u"{}[]|;<>?`~")
    for ch in name:
        if ch in invalid_chars:
            cleaned.append(u" ")
        elif ord(ch) < 32:
            cleaned.append(u" ")
        else:
            cleaned.append(ch)
    name = u" ".join(u"".join(cleaned).split())
    if not name:
        name = u"Imported Excel Table"
    return name


def make_unique_name(base_name, existing_names):
    base = sanitize_revit_view_name(base_name)
    existing = {}
    try:
        for item in existing_names or []:
            existing[clean_display_text(item).strip().lower()] = True
    except Exception:
        pass
    if base.lower() not in existing:
        existing[base.lower()] = True
        return base
    index = 2
    while True:
        candidate = u"%s %s" % (base, index)
        if candidate.lower() not in existing:
            existing[candidate.lower()] = True
            return candidate
        index += 1


def get_existing_revit_view_names(doc=None):
    names = []
    if doc is None:
        doc = get_revit_document()
    if doc is None:
        return names
    try:
        for view in FilteredElementCollector(doc).OfClass(DB.View):
            try:
                if not view.IsTemplate:
                    names.append(clean_display_text(view.Name))
            except Exception:
                pass
    except Exception:
        pass
    return names


def get_default_view_name(file_path, worksheet_name, region_display, existing_names):
    base = normalize_region_name_for_view(region_display, worksheet_name)
    return make_unique_name(base, existing_names)


def ensure_table_entry_uid(entry):
    try:
        uid = clean_display_text(getattr(entry, "TableEntryUid", "")).strip()
        if uid:
            return uid
    except Exception:
        uid = u""
    try:
        uid = safe_unicode(System.Guid.NewGuid().ToString("N"))
    except Exception:
        uid = safe_unicode(System.Guid.NewGuid())
    try:
        entry.TableEntryUid = uid
    except Exception:
        pass
    return uid


def _clean_tag_piece(value):
    text = clean_display_text(value).replace(u"|", u"/")
    return u" ".join(text.replace(u"\r", u" ").replace(u"\n", u" ").split())


def get_table_importer_tag_value(entry):
    uid = ensure_table_entry_uid(entry)
    source = _clean_tag_piece(getattr(entry, "FilePath", ""))
    worksheet = _clean_tag_piece(getattr(entry, "Worksheet", ""))
    region = _clean_tag_piece(getattr(entry, "Region", ""))
    return u"%s%s|tool_name=%s|table_entry_uid=%s|source_file=%s|worksheet=%s|region=%s|created_by=%s" % (
        TABLE_IMPORTER_TAG_PREFIX,
        uid,
        TABLE_IMPORTER_TOOL_NAME,
        uid,
        source,
        worksheet,
        region,
        TABLE_IMPORTER_CREATED_BY,
    )


def get_comments_parameter(element):
    if element is None:
        return None
    try:
        param = element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if param is not None:
            return param
    except Exception:
        pass
    try:
        return element.LookupParameter("Comments")
    except Exception:
        return None


def read_table_importer_tag(element):
    param = get_comments_parameter(element)
    if param is None:
        return u""
    try:
        return safe_unicode(param.AsString())
    except Exception:
        return u""


def apply_table_importer_tag(element, entry):
    if element is None or entry is None:
        return False
    value = get_table_importer_tag_value(entry)
    param = get_comments_parameter(element)
    if param is None:
        if DEBUG_OUTPUT:
            print("Table Importer: element has no writable Comments parameter for internal tag.")
        return False
    try:
        if param.IsReadOnly:
            if DEBUG_OUTPUT:
                print("Table Importer: element Comments parameter is read-only; tag not written.")
            return False
    except Exception:
        pass
    try:
        param.Set(value)
        return True
    except Exception as ex:
        if DEBUG_OUTPUT:
            print("Table Importer: could not write internal tag: %s" % safe_unicode(ex))
    return False


def element_has_table_importer_tag(element, entry):
    tag = read_table_importer_tag(element)
    if not tag:
        return False
    try:
        uid = ensure_table_entry_uid(entry)
    except Exception:
        uid = u""
    if uid:
        return tag.startswith(TABLE_IMPORTER_TAG_PREFIX + uid)
    return tag.startswith(TABLE_IMPORTER_TAG_PREFIX)


def safe_bool(value):
    try:
        if isinstance(value, bool):
            return value
    except Exception:
        pass
    text = safe_unicode(value).strip().lower()
    return text in (u"1", u"true", u"yes", u"si", u"sí", u"y")


def load_xaml(xaml_path):
    stream = FileStream(xaml_path, FileMode.Open, FileAccess.Read)
    try:
        return XamlReader.Load(stream)
    finally:
        stream.Close()


def get_storage_folder():
    try:
        return os.path.dirname(get_storage_path())
    except Exception:
        appdata = os.getenv("APPDATA") or script_dir
        return os.path.join(appdata, "pyMenvic", "TableImporter")


def resolve_entry_path(file_path):
    path = safe_unicode(file_path)
    if not path:
        return ""
    if os.path.isabs(path):
        return path
    return os.path.abspath(os.path.join(get_storage_folder(), path))


def find_visual_parent(child, parent_type):
    try:
        parent = VisualTreeHelper.GetParent(child)
        while parent is not None:
            if isinstance(parent, parent_type):
                return parent
            parent = VisualTreeHelper.GetParent(parent)
    except Exception:
        pass
    return None


def open_with_windows_shell(path):
    if not path:
        return False
    try:
        os.startfile(path)
        return True
    except Exception:
        pass
    try:
        psi = System.Diagnostics.ProcessStartInfo()
        psi.FileName = path
        psi.UseShellExecute = True
        System.Diagnostics.Process.Start(psi)
        return True
    except Exception:
        return False


def same_path(a, b):
    try:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))
    except Exception:
        return safe_unicode(a).lower() == safe_unicode(b).lower()


def get_combo_text(combo):
    try:
        item = combo.SelectedItem
        if item is None:
            return u""
        return safe_unicode(item)
    except Exception:
        return u""


def fill_combo(combo, values, selected=None):
    combo.Items.Clear()
    for value in values:
        combo.Items.Add(safe_unicode(value))
    if selected:
        for i in range(combo.Items.Count):
            if safe_unicode(combo.Items[i]) == safe_unicode(selected):
                combo.SelectedIndex = i
                return
    if combo.Items.Count:
        combo.SelectedIndex = 0

def get_shared_logo_path():
    extension_dir = os.path.abspath(os.path.join(script_dir, "..", "..", ".."))
    return os.path.join(extension_dir, "_resources", "logos", "menvic_logo.png")


def load_logo_into_image(image_control):
    try:
        if image_control is None:
            return
        logo_path = get_shared_logo_path()
        if not os.path.exists(logo_path):
            image_control.Source = None
            return
        bitmap = BitmapImage()
        bitmap.BeginInit()
        bitmap.UriSource = Uri(logo_path, System.UriKind.Absolute)
        bitmap.CacheOption = BitmapCacheOption.OnLoad
        bitmap.EndInit()
        image_control.Source = bitmap
    except Exception:
        try:
            image_control.Source = None
        except Exception:
            pass


def get_revit_document():
    try:
        return __revit__.ActiveUIDocument.Document
    except Exception:
        return None


def get_revit_uidocument():
    try:
        return __revit__.ActiveUIDocument
    except Exception:
        return None


def get_element_id_value(element_id):
    if element_id is None:
        return None
    try:
        return element_id.IntegerValue
    except Exception:
        pass
    try:
        return element_id.Value
    except Exception:
        pass
    try:
        return int(str(element_id))
    except Exception:
        print("Table Importer: could not convert ElementId value: %s" % safe_unicode(element_id))
        return None


def make_element_id(value):
    if value is None:
        return None

    try:
        numeric_value = int(value)
    except Exception:
        return None

    try:
        return DB.ElementId(Convert.ToInt64(numeric_value))
    except Exception:
        pass

    try:
        return DB.ElementId(Int64(numeric_value))
    except Exception:
        pass

    try:
        return DB.ElementId(Int32(numeric_value))
    except Exception:
        pass

    return None


def get_default_drafting_view_type(doc):
    collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)
    for view_type in collector:
        try:
            if view_type.ViewFamily == ViewFamily.Drafting:
                return view_type
        except Exception:
            pass
    return None


def get_default_text_note_type(doc):
    try:
        note_type = FilteredElementCollector(doc).OfClass(TextNoteType).FirstElement()
        return note_type
    except Exception:
        return None


def get_table_text_size():
    try:
        text_size = float(TABLE_TEXT_SIZE) * float(TEXT_SIZE_SCALE)
    except Exception:
        text_size = TABLE_TEXT_SIZE
    if text_size < MIN_TEXT_SIZE_FT:
        text_size = MIN_TEXT_SIZE_FT
    if text_size > MAX_TEXT_SIZE_FT:
        text_size = MAX_TEXT_SIZE_FT
    return text_size


def set_text_note_type_size(note_type, text_size):
    if note_type is None:
        return False
    try:
        size_param = note_type.get_Parameter(DB.BuiltInParameter.TEXT_SIZE)
        if size_param is not None and not size_param.IsReadOnly:
            size_param.Set(float(text_size))
            return True
    except Exception as ex:
        if DEBUG_OUTPUT:
            print("Table Importer: could not set table text size: %s" % safe_unicode(ex))
    return False


def set_text_note_type_font(note_type, font_name):
    if note_type is None:
        return False
    font_value = clean_display_text(font_name)
    if not font_value:
        font_value = FALLBACK_TABLE_FONT
    params = []
    try:
        text_font_param = getattr(DB.BuiltInParameter, "TEXT_FONT")
        params.append(note_type.get_Parameter(text_font_param))
    except Exception:
        pass
    try:
        params.append(note_type.LookupParameter("Text Font"))
    except Exception:
        pass
    try:
        params.append(note_type.LookupParameter("Font"))
    except Exception:
        pass
    for param in params:
        try:
            if param is not None and not param.IsReadOnly:
                param.Set(font_value)
                return True
        except Exception:
            pass
    if DEBUG_OUTPUT:
        print("Table Importer: could not set table font '%s' on %s." % (safe_unicode(font_value), safe_unicode(TABLE_TEXT_TYPE_NAME)))
    return False


def get_table_text_font(table_data):
    try:
        font_name = clean_display_text(getattr(table_data, "dominant_region_font", None))
        if font_name:
            return font_name
    except Exception:
        pass
    return FALLBACK_TABLE_FONT


def get_text_type_name(note_type):
    if note_type is None:
        return u""
    try:
        p = note_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            value = safe_unicode(p.AsString())
            if value:
                return clean_display_text(value)
    except Exception:
        pass
    try:
        return clean_display_text(note_type.Name)
    except Exception:
        return u""


def get_text_type_font(note_type):
    if note_type is None:
        return u""
    try:
        p = note_type.get_Parameter(DB.BuiltInParameter.TEXT_FONT)
        if p:
            return clean_display_text(p.AsString())
    except Exception:
        pass
    try:
        p = note_type.LookupParameter("Text Font")
        if p:
            return clean_display_text(p.AsString())
    except Exception:
        pass
    return u""


def get_text_type_size_mm(note_type):
    try:
        p = note_type.get_Parameter(DB.BuiltInParameter.TEXT_SIZE)
        if p:
            return float(p.AsDouble()) * 304.8
    except Exception:
        pass
    return None


def _get_yes_no_type_param(note_type, builtin_name, lookup_names):
    if note_type is None:
        return None
    try:
        builtin = getattr(DB.BuiltInParameter, builtin_name)
        p = note_type.get_Parameter(builtin)
        if p is not None:
            return p
    except Exception:
        pass
    for name in lookup_names or []:
        try:
            p = note_type.LookupParameter(name)
            if p is not None:
                return p
        except Exception:
            pass
    return None


def get_text_type_bold(note_type):
    p = _get_yes_no_type_param(note_type, "TEXT_STYLE_BOLD", ["Bold", "Negrita"])
    try:
        return bool(p.AsInteger()) if p is not None else False
    except Exception:
        return False


def get_text_type_italic(note_type):
    p = _get_yes_no_type_param(note_type, "TEXT_STYLE_ITALIC", ["Italic", "Cursive", "Italics"])
    try:
        return bool(p.AsInteger()) if p is not None else False
    except Exception:
        return False


def set_text_note_type_yes_no(note_type, builtin_name, lookup_names, value):
    p = _get_yes_no_type_param(note_type, builtin_name, lookup_names)
    try:
        if p is not None and not p.IsReadOnly:
            p.Set(1 if value else 0)
            return True
    except Exception:
        pass
    return False


def set_text_note_type_bold(note_type, value):
    return set_text_note_type_yes_no(note_type, "TEXT_STYLE_BOLD", ["Bold", "Negrita"], value)


def set_text_note_type_italic(note_type, value):
    return set_text_note_type_yes_no(note_type, "TEXT_STYLE_ITALIC", ["Italic", "Cursive", "Italics"], value)


def _normalize_font_name(font_name):
    text = clean_display_text(font_name).strip()
    return text if text else FALLBACK_TABLE_FONT


def _excel_points_to_text_mm(size_pt):
    try:
        value = float(size_pt) * EXCEL_POINT_TO_MM
    except Exception:
        value = STANDARD_TABLE_TEXT_SIZE_MM
    if value < MIN_EXCEL_TEXT_SIZE_MM:
        value = MIN_EXCEL_TEXT_SIZE_MM
    if value > MAX_EXCEL_TEXT_SIZE_MM:
        value = MAX_EXCEL_TEXT_SIZE_MM
    return value


def _make_excel_text_type_name(font_name, size_mm, bold=False, italic=False):
    parts = [u"Excel", _normalize_font_name(font_name), (u"%.2fmm" % float(size_mm))]
    if bold:
        parts.append(u"Bold")
    if italic:
        parts.append(u"Italic")
    return clean_display_text(u" ".join(parts))


def _text_type_matches(note_type, font_name, size_mm, bold=False, italic=False):
    try:
        if get_text_type_font(note_type).strip().lower() != _normalize_font_name(font_name).lower():
            return False
        existing_size = get_text_type_size_mm(note_type)
        if existing_size is None or abs(float(existing_size) - float(size_mm)) > 0.08:
            return False
        if bool(get_text_type_bold(note_type)) != bool(bold):
            return False
        if bool(get_text_type_italic(note_type)) != bool(italic):
            return False
        return True
    except Exception:
        return False


def _get_or_create_excel_text_type(doc, font_name, size_mm, bold=False, italic=False):
    """Find or create a Revit TextNoteType that matches the Excel cell style."""
    font_name = _normalize_font_name(font_name)
    try:
        size_mm = float(size_mm)
    except Exception:
        size_mm = STANDARD_TABLE_TEXT_SIZE_MM
    if size_mm < MIN_EXCEL_TEXT_SIZE_MM:
        size_mm = MIN_EXCEL_TEXT_SIZE_MM
    if size_mm > MAX_EXCEL_TEXT_SIZE_MM:
        size_mm = MAX_EXCEL_TEXT_SIZE_MM

    key = (font_name.lower(), round(size_mm, 2), bool(bold), bool(italic))
    try:
        cached_id = TEXT_TYPE_CACHE.get(key)
        if cached_id is not None:
            cached = doc.GetElement(cached_id)
            if cached is not None:
                return cached
    except Exception:
        pass

    all_types = []
    try:
        all_types = list(FilteredElementCollector(doc).OfClass(TextNoteType))
    except Exception:
        all_types = []

    for note_type in all_types:
        if _text_type_matches(note_type, font_name, size_mm, bold, italic):
            try:
                TEXT_TYPE_CACHE[key] = note_type.Id
            except Exception:
                pass
            return note_type

    base_type = get_table_text_note_type(doc, font_name)
    if base_type is None:
        return get_default_text_note_type(doc)

    new_type = None
    type_name = _make_excel_text_type_name(font_name, size_mm, bold, italic)
    try:
        existing_names = set([get_text_type_name(t).strip().lower() for t in all_types])
        final_name = type_name
        index = 2
        while final_name.strip().lower() in existing_names:
            final_name = u"%s %s" % (type_name, index)
            index += 1
        new_id = base_type.Duplicate(final_name)
        new_type = doc.GetElement(new_id)
        set_text_note_type_font(new_type, font_name)
        set_text_note_type_size(new_type, float(size_mm) / 304.8)
        set_text_note_type_bold(new_type, bold)
        set_text_note_type_italic(new_type, italic)
        TEXT_TYPE_CACHE[key] = new_type.Id
        return new_type
    except Exception as ex:
        if DEBUG_OUTPUT:
            print("Table Importer: could not create Excel text style '%s': %s" % (safe_unicode(type_name), safe_unicode(ex)))
    return base_type


def _get_cell_style(table_data, row_index, col_index):
    try:
        styles = getattr(table_data, "cell_styles", None)
        if not styles:
            return None
        return styles[row_index][col_index]
    except Exception:
        return None


def get_text_note_type_for_cell(doc, table_data, row_index, col_index, default_note_type):
    style = _get_cell_style(table_data, row_index, col_index)
    if not style:
        return default_note_type
    font_name = _normalize_font_name(style.get("font_name"))
    size_pt = style.get("size_pt")
    if not font_name and not size_pt:
        return default_note_type
    size_mm = _excel_points_to_text_mm(size_pt) if size_pt else STANDARD_TABLE_TEXT_SIZE_MM
    return _get_or_create_excel_text_type(doc, font_name, size_mm, bool(style.get("bold")), bool(style.get("italic")))


def _normalize_excel_alignment(value, fallback="left"):
    text = clean_display_text(value).strip().lower()
    if not text or text == "none":
        return fallback
    if text in ("center", "centre", "centercontinuous", "distributed"):
        return "center"
    if text in ("right",):
        return "right"
    if text in ("top",):
        return "top"
    if text in ("middle", "center", "centre"):
        return "middle"
    if text in ("bottom",):
        return "bottom"
    return fallback


def _text_role_stats(value):
    text = clean_display_text(value).strip()
    alpha = 0
    digit = 0
    lower = 0
    upper = 0
    for ch in text:
        try:
            if ch.isalpha():
                alpha += 1
                if ch.islower():
                    lower += 1
                if ch.isupper():
                    upper += 1
            elif ch.isdigit():
                digit += 1
        except Exception:
            pass
    words = [word for word in re.split(r"\s+", text) if word]
    return text, alpha, digit, lower, upper, words


def _is_short_code_like(value):
    text, alpha, digit, lower, upper, words = _text_role_stats(value)
    if not text:
        return False
    compact = text.replace(u" ", u"")
    if re.match(r"^[+-]?\d+([\.,/]\d+)?([\"']|mm|cm|m|a|v|w|kw|va|kva|awg)?$", compact.lower()):
        return True
    if re.match(r"^[A-Za-z]{1,3}\d{1,3}[A-Za-z]?$", compact):
        return True
    if re.match(r"^\d+[A-Za-z]{1,3}$", compact):
        return True
    if re.match(r"^AWG\d{1,3}$", compact.upper()):
        return True
    if digit > 0 and len(text) <= 14 and len(words) <= 3:
        return True
    if lower == 0 and len(compact) <= 10 and re.match(r"^[A-Z0-9\-_/\.\"'=+]+$", compact):
        return True
    return False


def _is_description_like(value, cell_width):
    text, alpha, digit, lower, upper, words = _text_role_stats(value)
    if not text:
        return False
    if _is_short_code_like(text):
        return False
    alnum = alpha + digit
    alpha_ratio = 0.0
    if alnum > 0:
        alpha_ratio = float(alpha) / float(alnum)
    multiple_words = len(words) >= 2
    try:
        wide_column = float(cell_width) >= 0.70
    except Exception:
        wide_column = False
    if alpha_ratio >= 0.60 and ((len(text) > 18) or (multiple_words and len(text) > 10)):
        return True
    if wide_column and multiple_words and alpha_ratio >= 0.50 and lower > 0:
        return True
    return False


def _cell_span_is_merged(row_index, col_index, top_left_map):
    try:
        span = _get_cell_span(row_index, col_index, top_left_map or {})
        return int(span[2]) > int(span[0]) or int(span[3]) > int(span[1])
    except Exception:
        return False


def _row_header_like(table_data, row_index):
    try:
        row = table_data[row_index]
    except Exception:
        return False
    non_empty = 0
    shortish = 0
    styled = 0
    for col_index in range(0, len(row)):
        value = clean_display_text(row[col_index]).strip()
        if not value:
            continue
        non_empty += 1
        if len(value) <= 18:
            shortish += 1
        style = _get_cell_style(table_data, row_index, col_index)
        try:
            if style and (style.get("bold") or _normalize_excel_alignment(style.get("horizontal"), "") == "center"):
                styled += 1
        except Exception:
            pass
    if non_empty <= 0:
        return False
    if non_empty <= 2 and styled > 0:
        return True
    return shortish >= max(1, non_empty - 1) and styled >= max(1, non_empty / 2)


def _cell_header_or_title_like(table_data, row_index, col_index, value, top_left_map):
    if _cell_span_is_merged(row_index, col_index, top_left_map):
        return True
    style = _get_cell_style(table_data, row_index, col_index)
    try:
        if style and style.get("bold") and _row_header_like(table_data, row_index):
            return True
    except Exception:
        pass
    text = clean_display_text(value).strip()
    if text and text.upper() == text and len(text) <= 28 and _row_header_like(table_data, row_index):
        return True
    return False


def _is_center_candidate(value):
    text = clean_display_text(value).strip()
    if not text:
        return False
    return _is_short_code_like(text)


def get_cell_horizontal_alignment(table_data, row_index, col_index, value, cell_width=None, top_left_map=None):
    style = _get_cell_style(table_data, row_index, col_index)
    is_description = _is_description_like(value, cell_width)
    is_header_or_title = _cell_header_or_title_like(table_data, row_index, col_index, value, top_left_map)
    try:
        if style and style.get("horizontal"):
            excel_alignment = _normalize_excel_alignment(style.get("horizontal"), "left")
            if is_description and not is_header_or_title:
                return "left"
            return excel_alignment
    except Exception:
        pass
    if is_description:
        return "left"
    if is_header_or_title:
        return "center"
    if _is_center_candidate(value):
        return "center"
    return "left"


def get_cell_vertical_alignment(table_data, row_index, col_index):
    style = _get_cell_style(table_data, row_index, col_index)
    try:
        if style and style.get("vertical"):
            return _normalize_excel_alignment(style.get("vertical"), "middle")
    except Exception:
        pass
    return "middle"


def _set_text_note_alignment_options(options, horizontal, vertical):
    try:
        if horizontal == "center":
            options.HorizontalAlignment = DB.HorizontalTextAlignment.Center
        elif horizontal == "right":
            options.HorizontalAlignment = DB.HorizontalTextAlignment.Right
        else:
            options.HorizontalAlignment = DB.HorizontalTextAlignment.Left
    except Exception:
        pass
    try:
        if vertical == "middle":
            options.VerticalAlignment = DB.VerticalTextAlignment.Middle
        elif vertical == "bottom":
            options.VerticalAlignment = DB.VerticalTextAlignment.Bottom
        else:
            options.VerticalAlignment = DB.VerticalTextAlignment.Top
    except Exception:
        pass
    return options


def get_text_point_for_alignment(x1, y1, x2, y2, horizontal, vertical, text_size):
    try:
        if horizontal == "center":
            x = (float(x1) + float(x2)) / 2.0
        elif horizontal == "right":
            x = float(x2) - TEXT_PADDING_X
        else:
            x = float(x1) + TEXT_PADDING_X
    except Exception:
        x = float(x1) + TEXT_PADDING_X
    try:
        if vertical == "middle":
            y = ((float(y1) + float(y2)) / 2.0) + (float(text_size) * 0.35)
        elif vertical == "bottom":
            y = float(y2) + TEXT_PADDING_Y + (float(text_size) * 0.35)
        else:
            y = float(y1) - TEXT_PADDING_Y
    except Exception:
        y = float(y1) - TEXT_PADDING_Y
    return XYZ(x, y, 0.0)


def is_bad_table_importer_text_type_name(type_name):
    """Return True for legacy/custom Table Importer text styles that should not be reused."""
    name = clean_display_text(type_name).strip().lower()
    if not name:
        return False
    if "menvic_table_text" in name:
        return True
    if "table importer" in name:
        return True
    return False


def is_allowed_table_text_type(note_type):
    try:
        return not is_bad_table_importer_text_type_name(get_text_type_name(note_type))
    except Exception:
        return False


def is_name_match(value, candidates):
    clean_value = clean_display_text(value).strip().lower()
    for candidate in candidates:
        if clean_value == clean_display_text(candidate).strip().lower():
            return True
    return False


def is_office_standard_table_text_type(note_type, min_size_mm=STANDARD_TABLE_TEXT_SIZE_MM):
    try:
        if not is_allowed_table_text_type(note_type):
            return False
        font_name = get_text_type_font(note_type).strip().lower()
        if font_name != FALLBACK_TABLE_FONT.lower():
            return False
        size_mm = get_text_type_size_mm(note_type)
        if size_mm is None:
            return False
        return abs(float(size_mm) - float(min_size_mm)) <= STANDARD_TEXT_SIZE_TOL_MM
    except Exception:
        return False


def get_table_text_note_type(doc, font_name=None):
    """
    Use the existing pyMENVIC office-standard text style.

    This tool must not create custom Table Importer text types. It looks for
    Arial 1.80mm first, then for an Arial TextNoteType at 1.80mm, then for the
    nearest Arial standard type >= 1.80mm. If no standard is found, it falls
    back to the first available TextNoteType without modifying it.
    """
    global TABLE_TEXT_STYLE_WARNING

    all_types = []
    try:
        all_types = list(FilteredElementCollector(doc).OfClass(TextNoteType))
    except Exception:
        all_types = []

    # 1) Exact/known office standard names. Never reuse legacy Table Importer types.
    standard_name_candidates = [
        STANDARD_TABLE_TEXT_TYPE_NAME,
        "Arial 1.8mm",
        "1.80mm Arial",
        "1.8mm Arial",
        "1.80 mm Arial",
        "1.8 mm Arial",
    ]
    for note_type in all_types:
        try:
            type_name = get_text_type_name(note_type)
            if is_allowed_table_text_type(note_type) and is_name_match(type_name, standard_name_candidates):
                return note_type
        except Exception:
            pass

    # 2) Any Arial 1.80mm type.
    for note_type in all_types:
        try:
            if is_office_standard_table_text_type(note_type, STANDARD_TABLE_TEXT_SIZE_MM):
                return note_type
        except Exception:
            pass

    # 3) Nearest Arial type >= 1.80mm.
    best_type = None
    best_size = None
    for note_type in all_types:
        try:
            if not is_allowed_table_text_type(note_type):
                continue
            if get_text_type_font(note_type).strip().lower() != FALLBACK_TABLE_FONT.lower():
                continue
            size_mm = get_text_type_size_mm(note_type)
            if size_mm is None:
                continue
            if size_mm + STANDARD_TEXT_SIZE_TOL_MM < STANDARD_TABLE_TEXT_SIZE_MM:
                continue
            if best_size is None or float(size_mm) < float(best_size):
                best_size = size_mm
                best_type = note_type
        except Exception:
            pass
    if best_type is not None:
        if DEBUG_OUTPUT:
            print("Table Importer: Arial 1.80mm not found. Using nearest Arial text style '%s'." % safe_unicode(get_text_type_name(best_type)))
        return best_type

    # 4) Fallback only. Do not modify or duplicate. Prefer a non-Table-Importer type.
    fallback = None
    for note_type in all_types:
        try:
            if is_allowed_table_text_type(note_type):
                fallback = note_type
                break
        except Exception:
            pass
    if fallback is None:
        fallback = get_default_text_note_type(doc)
    if fallback is not None and not TABLE_TEXT_STYLE_WARNING:
        TABLE_TEXT_STYLE_WARNING = True
        if DEBUG_OUTPUT:
            print("Table Importer: standard text style Arial 1.80mm not found. Using fallback text style '%s'." % safe_unicode(get_text_type_name(fallback)))
    return fallback


def get_text_note_type_size(note_type):
    if note_type is None:
        return get_table_text_size()
    try:
        size_param = note_type.get_Parameter(DB.BuiltInParameter.TEXT_SIZE)
        if size_param is not None:
            value = float(size_param.AsDouble())
            if value > 0:
                return value
    except Exception:
        pass
    return get_table_text_size()


def get_existing_view(doc, element_id):
    if element_id is None:
        return None
    try:
        return doc.GetElement(element_id)
    except Exception as ex:
        print("Table Importer: could not resolve ElementId '%s': %s" % (safe_unicode(element_id), safe_unicode(ex)))
        return None


def is_drafting_view(view):
    try:
        return view is not None and view.ViewType == DB.ViewType.DraftingView
    except Exception:
        pass
    try:
        view_type = safe_unicode(view.ViewType)
        clean_view_type = view_type.replace(" ", "").lower()
        return view is not None and clean_view_type == "draftingview"
    except Exception:
        pass
    return False


def normalize_view_type_name(value):
    text = clean_display_text(value).strip()
    if not text:
        return "Drafting View"
    lowered = text.lower()
    if "legend" in lowered:
        return "Legend View"
    if "schedule" in lowered:
        return "Schedule View"
    return "Drafting View"


def get_view_type_display_name(value):
    view_type = normalize_view_type_name(value)
    if view_type == "Legend View":
        return "Legend View (Not implemented)"
    if view_type == "Schedule View":
        return "Schedule View"
    return "Drafting View"


def get_or_create_drafting_view(entry):
    doc = get_revit_document()
    if doc is None:
        raise Exception("No active Revit document.")
    ensure_table_entry_uid(entry)

    view_type = get_default_drafting_view_type(doc)
    if view_type is None:
        raise Exception("No Drafting View type found in this project.")

    desired_name = clean_display_text(entry.ViewName).strip()
    if not desired_name:
        desired_name = get_default_view_name(entry.FilePath, entry.Worksheet, entry.Region, [])
    view_name = make_unique_name(desired_name, get_existing_revit_view_names(doc))
    view = ViewDrafting.Create(doc, view_type.Id)
    view.Name = view_name
    entry.ViewName = clean_display_text(view.Name)
    view_id_value = get_element_id_value(view.Id)
    if view_id_value is None:
        print("Table Importer: created view has an ElementId that could not be stored.")
        entry.RevitViewId = None
    else:
        entry.RevitViewId = safe_unicode(view_id_value)
    return view, True


def read_table_data_for_entry(entry):
    path = resolve_entry_path(entry.FilePath)
    if not path or not os.path.exists(path):
        raise Exception("Excel file not found: %s" % safe_unicode(path))

    table_data, row_count, column_count = read_excel_table_values(path, entry.Worksheet, entry.Region)
    if row_count <= 0 or column_count <= 0:
        raise Exception("No readable Excel data.")

    if row_count > MAX_TABLE_ROWS or column_count > MAX_TABLE_COLUMNS:
        raise Exception("Table is too large for MVP: %s row(s) x %s column(s)." % (row_count, column_count))

    return table_data, row_count, column_count


def _same_element_id(left, right):
    return get_element_id_value(left) == get_element_id_value(right)


def is_table_importer_drawn_element(element, view, entry):
    # Regeneration removes only tagged Table Importer table text and border
    # lines. Manual TextNotes, manual DetailCurves, symbols, dimensions, tags,
    # images, and other untagged view content are intentionally preserved.
    try:
        if element is None or not _same_element_id(element.OwnerViewId, view.Id):
            return False
    except Exception:
        return False

    is_drawn_class = False
    try:
        if isinstance(element, TextNote):
            is_drawn_class = True
    except Exception:
        pass

    try:
        if isinstance(element, DB.CurveElement):
            is_drawn_class = True
    except Exception:
        pass

    if not is_drawn_class:
        return False
    return element_has_table_importer_tag(element, entry)


def _is_legacy_untagged_drawn_element(element, view):
    try:
        if element is None or not _same_element_id(element.OwnerViewId, view.Id):
            return False
    except Exception:
        return False

    is_drawn_class = False
    try:
        if isinstance(element, TextNote):
            is_drawn_class = True
    except Exception:
        pass
    try:
        if isinstance(element, DB.CurveElement):
            is_drawn_class = True
    except Exception:
        pass
    if not is_drawn_class:
        return False
    return not read_table_importer_tag(element)


def _collect_legacy_untagged_ids(doc, view):
    ids = []
    text_count = 0
    curve_count = 0
    for element_class in (TextNote, DB.CurveElement):
        try:
            collector = FilteredElementCollector(doc, view.Id).OfClass(element_class)
            for element in collector:
                try:
                    if _is_legacy_untagged_drawn_element(element, view):
                        ids.append(element.Id)
                        try:
                            if isinstance(element, TextNote):
                                text_count += 1
                            else:
                                curve_count += 1
                        except Exception:
                            curve_count += 1
                except Exception:
                    pass
        except Exception as ex:
            if DEBUG_OUTPUT:
                print("Table Importer: legacy collection failed for '%s': %s" % (safe_unicode(element_class), safe_unicode(ex)))
    return ids, text_count, curve_count


def detect_legacy_untagged_table_content(view):
    doc = get_revit_document()
    if doc is None or view is None:
        return False, 0, 0
    ids, text_count, curve_count = _collect_legacy_untagged_ids(doc, view)
    return len(ids) > 0, text_count, curve_count


def clear_legacy_untagged_table_content(view):
    doc = get_revit_document()
    if doc is None:
        raise Exception("No active Revit document.")
    ids_to_delete, text_count, curve_count = _collect_legacy_untagged_ids(doc, view)
    deleted = 0
    for element_id in ids_to_delete:
        try:
            doc.Delete(element_id)
            deleted += 1
        except Exception as ex:
            print("Table Importer: could not delete legacy view element '%s': %s" % (safe_unicode(element_id), safe_unicode(ex)))
    if DEBUG_OUTPUT:
        print("Table Importer: legacy cleanup deleted %s TextNote(s) and %s CurveElement(s)." % (text_count, curve_count))
    return deleted, text_count, curve_count


def _collect_tagged_elements_of_class(doc, view, entry, element_class):
    ids = []
    try:
        collector = FilteredElementCollector(doc, view.Id).OfClass(element_class)
        for element in collector:
            try:
                if is_table_importer_drawn_element(element, view, entry):
                    ids.append(element.Id)
            except Exception:
                pass
    except Exception as ex:
        print("Table Importer: collection failed for '%s': %s" % (safe_unicode(element_class), safe_unicode(ex)))
    return ids


def is_table_importer_deletable_class(element):
    try:
        if isinstance(element, TextNote):
            return True
    except Exception:
        pass
    try:
        if isinstance(element, DB.CurveElement):
            return True
    except Exception:
        pass
    return False


def remember_created_element(entry, element):
    if entry is None or element is None:
        return
    try:
        element_id_value = get_element_id_value(element.Id)
        if element_id_value is None:
            return
        current = []
        try:
            current = list(getattr(entry, "CreatedElementIds", []) or [])
        except Exception:
            current = []
        text_id = safe_unicode(element_id_value)
        if text_id not in current:
            current.append(text_id)
            entry.CreatedElementIds = current
    except Exception:
        pass


def _collect_stored_created_element_ids(doc, view, entry):
    ids = []
    try:
        raw_ids = list(getattr(entry, "CreatedElementIds", []) or [])
    except Exception:
        raw_ids = []
    for raw_id in raw_ids:
        element_id = make_element_id(raw_id)
        if element_id is None:
            continue
        try:
            element = doc.GetElement(element_id)
        except Exception:
            element = None
        if element is None:
            continue
        try:
            if not _same_element_id(element.OwnerViewId, view.Id):
                continue
        except Exception:
            continue
        if not is_table_importer_deletable_class(element):
            continue
        ids.append(element.Id)
    return ids


def clear_table_importer_view(view, entry):
    doc = get_revit_document()
    if doc is None:
        raise Exception("No active Revit document.")

    ensure_table_entry_uid(entry)

    # Normal regeneration is intentionally strict:
    # - delete elements stored as created by this entry, or
    # - delete elements carrying the Table Importer internal tag for this entry.
    # Untagged/manual TextNotes, manual DetailCurves, symbols, dimensions, tags,
    # images, filled regions, and family instances are preserved.
    stored_ids = _collect_stored_created_element_ids(doc, view, entry)
    text_note_ids = _collect_tagged_elements_of_class(doc, view, entry, TextNote)
    curve_ids = _collect_tagged_elements_of_class(doc, view, entry, DB.CurveElement)

    ids_to_delete = []
    seen = {}
    for element_id in stored_ids + text_note_ids + curve_ids:
        try:
            key = safe_unicode(get_element_id_value(element_id))
            if key in seen:
                continue
            seen[key] = True
            ids_to_delete.append(element_id)
        except Exception:
            ids_to_delete.append(element_id)

    deleted = 0
    for element_id in ids_to_delete:
        try:
            doc.Delete(element_id)
            deleted += 1
        except Exception as ex:
            if DEBUG_OUTPUT:
                print("Table Importer: could not delete tagged/stored element '%s': %s" % (safe_unicode(element_id), safe_unicode(ex)))

    try:
        entry.CreatedElementIds = []
    except Exception:
        pass

    if DEBUG_OUTPUT:
        try:
            print("Table Importer: regeneration deleted %s stored/tagged element(s). Untagged manual content ignored." % deleted)
        except Exception:
            pass
    return deleted


def reset_table_view_legacy_content(view):
    deleted, text_count, curve_count = clear_legacy_untagged_table_content(view)
    return deleted


def update_existing_drafting_view(entry, view, cleanup_legacy=False):
    table_data, row_count, column_count = read_table_data_for_entry(entry)
    clear_table_importer_view(view, entry)
    # Legacy untagged content is never deleted during normal Apply.
    # It may be cleared only by a future explicit reset action.
    cleanup_legacy = False
    draw_table_in_view(view, table_data, entry)
    entry.Status = "Updated"
    return row_count, column_count


def import_to_legend_view(entry):
    raise Exception("Legend View import is not implemented yet.")


def is_schedule_view(view):
    try:
        if isinstance(view, DB.ViewSchedule):
            return True
    except Exception:
        pass
    try:
        return view is not None and view.ViewType == DB.ViewType.Schedule
    except Exception:
        pass
    try:
        view_type = safe_unicode(view.ViewType)
        return view is not None and view_type.replace(" ", "").lower() == "schedule"
    except Exception:
        pass
    return False


def _schedule_not_supported(reason):
    raise Exception("Schedule View header import is not supported in this Revit/API context. %s" % safe_unicode(reason))


def _get_schedule_category_ids(doc):
    category_ids = []
    candidates = []
    try:
        candidates.append(DB.BuiltInCategory.OST_GenericModel)
    except Exception:
        pass
    try:
        candidates.append(DB.BuiltInCategory.OST_Furniture)
    except Exception:
        pass
    try:
        candidates.append(DB.BuiltInCategory.OST_Doors)
    except Exception:
        pass
    try:
        candidates.append(DB.BuiltInCategory.OST_Walls)
    except Exception:
        pass
    try:
        candidates.append(DB.BuiltInCategory.OST_Rooms)
    except Exception:
        pass

    for built_in_category in candidates:
        try:
            category = doc.Settings.Categories.get_Item(built_in_category)
            if category is not None and category.Id is not None:
                category_ids.append(category.Id)
        except Exception:
            pass

    try:
        invalid_id = DB.ElementId.InvalidElementId
        if invalid_id is not None:
            category_ids.append(invalid_id)
    except Exception:
        pass
    return category_ids


def _existing_view_names_excluding(view, doc):
    names = []
    current_id = None
    try:
        current_id = get_element_id_value(view.Id)
    except Exception:
        current_id = None
    for candidate in FilteredElementCollector(doc).OfClass(DB.View):
        try:
            if candidate.IsTemplate:
                continue
            if current_id is not None and get_element_id_value(candidate.Id) == current_id:
                continue
            names.append(clean_display_text(candidate.Name))
        except Exception:
            pass
    return names


def _rename_schedule_view(schedule, entry, doc):
    desired_name = clean_display_text(entry.ViewName).strip()
    if not desired_name:
        desired_name = get_default_view_name(entry.FilePath, entry.Worksheet, entry.Region, [])
    desired_name = sanitize_revit_view_name(desired_name)
    try:
        current_name = clean_display_text(schedule.Name).strip()
    except Exception:
        current_name = u""
    if current_name.lower() != desired_name.lower():
        try:
            schedule.Name = make_unique_name(desired_name, _existing_view_names_excluding(schedule, doc))
        except Exception:
            schedule.Name = make_unique_name(desired_name, get_existing_revit_view_names(doc))
    try:
        entry.ViewName = clean_display_text(schedule.Name)
    except Exception:
        entry.ViewName = desired_name


def _create_schedule_view(doc, entry):
    desired_name = clean_display_text(entry.ViewName).strip()
    if not desired_name:
        desired_name = get_default_view_name(entry.FilePath, entry.Worksheet, entry.Region, [])
    desired_name = make_unique_name(desired_name, get_existing_revit_view_names(doc))

    category_ids = _get_schedule_category_ids(doc)
    last_error = None
    for category_id in category_ids:
        try:
            schedule = DB.ViewSchedule.CreateSchedule(doc, category_id)
            schedule.Name = desired_name
            entry.ViewName = clean_display_text(schedule.Name)
            schedule_id_value = get_element_id_value(schedule.Id)
            if schedule_id_value is None:
                entry.RevitViewId = None
            else:
                entry.RevitViewId = safe_unicode(schedule_id_value)
            return schedule
        except Exception as ex:
            last_error = ex
    _schedule_not_supported("Could not create a ViewSchedule: %s" % safe_unicode(last_error))


def _get_or_create_schedule_view(doc, entry):
    if entry.RevitViewId:
        element_id = make_element_id(entry.RevitViewId)
        if element_id is None:
            raise Exception("Invalid RevitViewId '%s'." % safe_unicode(entry.RevitViewId))
        view = doc.GetElement(element_id)
        if view is None:
            raise Exception("Missing Revit schedule for RevitViewId '%s'." % safe_unicode(entry.RevitViewId))
        if not is_schedule_view(view):
            try:
                revit_view_type = safe_unicode(view.ViewType)
            except Exception:
                revit_view_type = "<unknown>"
            raise Exception("Existing view is not a Schedule View; Revit view type is '%s'." % revit_view_type)
        _rename_schedule_view(view, entry, doc)
        return view, False

    schedule = _create_schedule_view(doc, entry)
    return schedule, True


def _get_section_count(section, count_name, first_name, last_name):
    try:
        return int(getattr(section, count_name))
    except Exception:
        pass
    try:
        first_value = int(getattr(section, first_name))
        last_value = int(getattr(section, last_name))
        return max(0, last_value - first_value + 1)
    except Exception:
        pass
    return 0


def _get_section_index(section, name, fallback):
    try:
        return int(getattr(section, name))
    except Exception:
        return fallback


def _get_header_section(schedule):
    try:
        table_data = schedule.GetTableData()
        return table_data.GetSectionData(DB.SectionType.Header)
    except Exception as ex:
        _schedule_not_supported("Could not access the schedule header section: %s" % safe_unicode(ex))


def _get_header_row_count(section):
    return _get_section_count(section, "NumberOfRows", "FirstRowNumber", "LastRowNumber")


def _get_header_col_count(section):
    return _get_section_count(section, "NumberOfColumns", "FirstColumnNumber", "LastColumnNumber")


def _insert_schedule_header_row(section):
    before = _get_header_row_count(section)
    insert_index = _get_section_index(section, "LastRowNumber", before - 1) + 1
    try:
        section.InsertRow(insert_index)
    except Exception:
        try:
            section.InsertRow(before)
        except Exception as ex:
            _schedule_not_supported("Could not add header row: %s" % safe_unicode(ex))
    after = _get_header_row_count(section)
    if after <= before:
        _schedule_not_supported("Header row count did not increase.")


def _insert_schedule_header_column(section):
    before = _get_header_col_count(section)
    insert_index = _get_section_index(section, "LastColumnNumber", before - 1) + 1
    try:
        section.InsertColumn(insert_index)
    except Exception:
        try:
            section.InsertColumn(before)
        except Exception as ex:
            _schedule_not_supported("Could not add header column: %s" % safe_unicode(ex))
    after = _get_header_col_count(section)
    if after <= before:
        _schedule_not_supported("Header column count did not increase.")


def _ensure_schedule_header_size(section, target_rows, target_cols):
    target_rows = max(1, int(target_rows))
    target_cols = max(1, int(target_cols))
    guard = 0
    while _get_header_row_count(section) < target_rows:
        guard += 1
        if guard > target_rows + 20:
            _schedule_not_supported("Could not size schedule header rows.")
        _insert_schedule_header_row(section)

    guard = 0
    while _get_header_col_count(section) < target_cols:
        guard += 1
        if guard > target_cols + 20:
            _schedule_not_supported("Could not size schedule header columns.")
        _insert_schedule_header_column(section)


def _set_schedule_header_cell_text(section, row_index, col_index, text):
    try:
        section.SetCellText(row_index, col_index, clean_display_text(text))
        return True
    except Exception:
        pass
    try:
        section.SetCellText(row_index, col_index, safe_unicode(text))
        return True
    except Exception:
        return False


def _clear_schedule_header_cells(section):
    first_row = _get_section_index(section, "FirstRowNumber", 0)
    first_col = _get_section_index(section, "FirstColumnNumber", 0)
    row_count = _get_header_row_count(section)
    col_count = _get_header_col_count(section)
    for row_index in range(first_row, first_row + row_count):
        for col_index in range(first_col, first_col + col_count):
            _set_schedule_header_cell_text(section, row_index, col_index, u"")


def _try_merge_schedule_header_cells(section, start_row, start_col, end_row, end_col):
    if end_row <= start_row and end_col <= start_col:
        return False
    try:
        merged_cell = DB.TableMergedCell(start_row, start_col, end_row, end_col)
        section.MergeCells(merged_cell)
        return True
    except Exception:
        pass
    try:
        merged_cell = DB.TableMergedCell(start_row, end_row, start_col, end_col)
        section.MergeCells(merged_cell)
        return True
    except Exception:
        pass
    try:
        section.MergeCells(start_row, start_col, end_row, end_col)
        return True
    except Exception as ex:
        if DEBUG_OUTPUT:
            print("Table Importer: schedule header merge skipped (%s,%s)-(%s,%s): %s" % (
                start_row,
                start_col,
                end_row,
                end_col,
                safe_unicode(ex),
            ))
    return False


def _apply_schedule_header_merges(section, table_data, first_row, first_col):
    try:
        top_left_map, covered_map = _build_merged_cell_maps(table_data)
    except Exception:
        return
    for key in top_left_map:
        try:
            min_row, min_col, max_row, max_col = top_left_map[key]
            _try_merge_schedule_header_cells(
                section,
                first_row + min_row,
                first_col + min_col,
                first_row + max_row,
                first_col + max_col,
            )
        except Exception:
            pass


def _write_table_data_to_schedule_header(schedule, table_data, row_count, column_count):
    section = _get_header_section(schedule)
    _ensure_schedule_header_size(section, row_count, column_count)
    first_row = _get_section_index(section, "FirstRowNumber", 0)
    first_col = _get_section_index(section, "FirstColumnNumber", 0)
    _clear_schedule_header_cells(section)

    try:
        top_left_map, covered_map = _build_merged_cell_maps(table_data)
    except Exception:
        top_left_map = {}
        covered_map = {}

    written = 0
    for row_index in range(row_count):
        for col_index in range(column_count):
            if _is_covered_non_top_left(row_index, col_index, top_left_map, covered_map):
                continue
            text = _get_cell_value(table_data, row_index, col_index)
            if _set_schedule_header_cell_text(section, first_row + row_index, first_col + col_index, text):
                written += 1

    if written <= 0 and row_count > 0 and column_count > 0:
        _schedule_not_supported("Could not write text into schedule header cells.")
    _apply_schedule_header_merges(section, table_data, first_row, first_col)


def import_to_schedule_view(entry):
    doc = get_revit_document()
    if doc is None:
        raise Exception("No active Revit document.")
    ensure_table_entry_uid(entry)
    table_data, row_count, column_count = read_table_data_for_entry(entry)

    transaction = Transaction(doc, "Import Table Importer Schedule View")
    transaction.Start()
    try:
        schedule, was_created = _get_or_create_schedule_view(doc, entry)
        _write_table_data_to_schedule_header(schedule, table_data, row_count, column_count)
        schedule_id_value = get_element_id_value(schedule.Id)
        if schedule_id_value is not None:
            entry.RevitViewId = safe_unicode(schedule_id_value)
        entry.Status = "Updated"
        transaction.Commit()
        transaction = None
        return "Created" if was_created else "Updated"
    except Exception as ex:
        try:
            if transaction is not None:
                transaction.RollBack()
        except Exception:
            pass
        message = safe_unicode(ex)
        if message.startswith("Schedule View header import is not supported"):
            raise
        if message.startswith("Invalid RevitViewId") or message.startswith("Missing Revit schedule") or "not a Schedule View" in message:
            raise
        _schedule_not_supported(message)


def import_to_drafting_view(entry, cleanup_legacy=False):
    doc = get_revit_document()
    if doc is None:
        raise Exception("No active Revit document.")
    transaction = None
    if entry.RevitViewId:
        element_id = make_element_id(entry.RevitViewId)
        if element_id is None:
            raise Exception("Invalid RevitViewId '%s'." % safe_unicode(entry.RevitViewId))
        view = doc.GetElement(element_id)
        if view is None:
            raise Exception("Missing Revit view for RevitViewId '%s'." % safe_unicode(entry.RevitViewId))
        if not is_drafting_view(view):
            try:
                revit_view_type = safe_unicode(view.ViewType)
            except Exception:
                revit_view_type = "<unknown>"
            raise Exception("Existing view is not a Drafting View; Revit view type is '%s'." % revit_view_type)
        transaction = Transaction(doc, "Update Table Importer Drafting View")
        transaction.Start()
        try:
            update_existing_drafting_view(entry, view, cleanup_legacy)
            transaction.Commit()
            transaction = None
            return "Updated"
        except Exception:
            try:
                if transaction is not None:
                    transaction.RollBack()
            except Exception:
                pass
            raise

    table_data, row_count, column_count = read_table_data_for_entry(entry)
    transaction = Transaction(doc, "Create Table Importer Drafting View")
    transaction.Start()
    try:
        view, was_created = get_or_create_drafting_view(entry)
        draw_table_in_view(view, table_data, entry)
        transaction.Commit()
        transaction = None
        return "Created" if was_created else "Updated"
    except Exception:
        try:
            if transaction is not None:
                transaction.RollBack()
        except Exception:
            pass
        raise


def _make_line(doc, view, x1, y1, x2, y2, entry=None):
    line = Line.CreateBound(XYZ(float(x1), float(y1), 0.0), XYZ(float(x2), float(y2), 0.0))
    detail_curve = doc.Create.NewDetailCurve(view, line)
    if entry is not None:
        apply_table_importer_tag(detail_curve, entry)
        remember_created_element(entry, detail_curve)
    return detail_curve


def _make_line_once(doc, view, drawn_lines, x1, y1, x2, y2, entry=None):
    try:
        key = (
            round(float(x1), 5),
            round(float(y1), 5),
            round(float(x2), 5),
            round(float(y2), 5),
        )
        reverse_key = (key[2], key[3], key[0], key[1])
        if key in drawn_lines or reverse_key in drawn_lines:
            return
        drawn_lines.add(key)
    except Exception:
        pass
    _make_line(doc, view, x1, y1, x2, y2, entry)


def _trim_note_text(text):
    value = clean_display_text(text)
    if len(value) > 250:
        return value[:247] + u"..."
    return value


def trim_text_to_cell_width(text, cell_width):
    value = clean_display_text(text)
    value = u" ".join(value.replace(u"\r", u" ").replace(u"\n", u" ").split())
    min_chars = 3
    try:
        max_chars = int(float(cell_width) * APPROX_CHARS_PER_FOOT)
    except Exception:
        max_chars = min_chars
    if max_chars < min_chars:
        max_chars = min_chars
    if len(value) > max_chars:
        trimmed = value[:max_chars - 3] + u"..."
        if DEBUG_OUTPUT:
            try:
                print("Table Importer: trimmed TextNote from %s to %s chars for %.3f ft cell." % (len(value), len(trimmed), float(cell_width)))
            except Exception:
                pass
        return trimmed
    return value


MIN_VISIBLE_CHARS = 4
# Physical character width in feet for Arial 1.6mm at model scale 1:1.
# Arial 1.6mm height = 1.6/304.8 ft.
# Average char width ≈ 55% of height.
# This is the PHYSICAL size independent of view scale.
_ARIAL_18MM_CHAR_WIDTH_FT = (1.6 / 304.8) * 0.55


def fit_text_to_cell(text, cell_width, text_size):
    # text_size is intentionally ignored here: the Revit TEXT_SIZE parameter
    # returns a value already multiplied by the view scale, which makes it
    # useless for estimating how many characters fit in a cell whose width is
    # expressed in model-space feet.  We use the known physical size of the
    # office standard text (Arial 1.8mm) directly.
    value = clean_display_text(text)
    value = u" ".join(value.replace(u"\r", u" ").replace(u"\n", u" ").split())
    if not value:
        return u""

    # Hard cap to avoid absurdly long strings.
    if len(value) > MAX_CELL_TEXT_CHARS:
        value = value[:MAX_CELL_TEXT_CHARS]

    try:
        available_width = max(float(cell_width) - 2.0 * float(TEXT_PADDING_X), 0.0)
    except Exception:
        available_width = 0.0
    if available_width <= 0:
        return u""

    try:
        char_w = _ARIAL_18MM_CHAR_WIDTH_FT
        max_chars = int(available_width / char_w)
    except Exception:
        max_chars = MAX_CELL_TEXT_CHARS

    # Clamp between MIN_VISIBLE_CHARS and MAX_CELL_TEXT_CHARS.
    if max_chars < MIN_VISIBLE_CHARS:
        max_chars = MIN_VISIBLE_CHARS
    if max_chars > MAX_CELL_TEXT_CHARS:
        max_chars = MAX_CELL_TEXT_CHARS

    # Text fits: keep as-is.
    if len(value) <= max_chars:
        return value

    # Cell too narrow for a meaningful stub.
    if max_chars <= 3:
        return u"..."

    return value[:max_chars - 3] + u"..."


def _debug_fit_text(original, fitted, cell_width, text_size, debug_count):
    if not DEBUG_OUTPUT:
        return
    try:
        if debug_count[0] >= 20:
            return
        debug_count[0] += 1
        estimated_width = float(len(clean_display_text(original))) * float(text_size) * float(APPROX_TEXT_CHAR_WIDTH_FACTOR)
        print("Table Importer text fit %s: original='%s' fitted='%s' cell_width=%.3f text_size=%.4f estimated_width=%.3f" % (
            debug_count[0],
            safe_unicode(original),
            safe_unicode(fitted),
            float(cell_width),
            float(text_size),
            estimated_width,
        ))
    except Exception:
        pass


def _get_cell_value(table_data, row_index, col_index):
    try:
        row = table_data[row_index]
        if col_index < len(row):
            return _trim_note_text(row[col_index])
    except Exception:
        pass
    return u""


def _ceil_div(value, divisor):
    try:
        return int((int(value) + int(divisor) - 1) / int(divisor))
    except Exception:
        return 1


def _get_merged_ranges(table_data):
    try:
        if not getattr(table_data, "merges_available", False):
            print("Table Importer: merged-cell data unavailable; drawing cells without merge metadata.")
            return []
        return getattr(table_data, "merged_ranges", []) or []
    except Exception:
        print("Table Importer: merged-cell data unavailable; drawing cells without merge metadata.")
        return []


def _build_merged_cell_maps(table_data):
    top_left_map = {}
    covered_map = {}
    merged_ranges = _get_merged_ranges(table_data)
    for item in merged_ranges:
        try:
            min_row = int(item.get("min_row", 0))
            min_col = int(item.get("min_col", 0))
            max_row = int(item.get("max_row", min_row))
            max_col = int(item.get("max_col", min_col))
            if max_row < min_row or max_col < min_col:
                continue
            top_left = (min_row, min_col)
            top_left_map[top_left] = (min_row, min_col, max_row, max_col)
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    covered_map[(r, c)] = top_left
        except Exception:
            pass
    return top_left_map, covered_map


def _get_cell_span(row_index, col_index, top_left_map):
    try:
        if (row_index, col_index) in top_left_map:
            return top_left_map[(row_index, col_index)]
    except Exception:
        pass
    return row_index, col_index, row_index, col_index


def _is_covered_non_top_left(row_index, col_index, top_left_map, covered_map):
    try:
        key = (row_index, col_index)
        return key in covered_map and key not in top_left_map
    except Exception:
        return False


def _sum_range(values, start_index, end_index):
    total = 0.0
    try:
        for index in range(start_index, end_index + 1):
            total += float(values[index])
    except Exception:
        pass
    return total

def _safe_float(value, fallback):
    try:
        return float(value)
    except Exception:
        return float(fallback)


def _clamp(value, min_value, max_value):
    value = _safe_float(value, min_value)
    if value < min_value:
        return min_value
    if value > max_value:
        return max_value
    return value


def _table_total_width(widths):
    total = 0.0
    try:
        for width in widths:
            total += float(width)
    except Exception:
        pass
    return total


def _merge_excel_and_content_column_widths(excel_widths, content_widths):
    # Generic strategy: Excel widths define proportions, content widths protect
    # readability. This avoids tuning the importer to one specific workbook.
    if not excel_widths:
        return list(content_widths or [])
    if not content_widths:
        return list(excel_widths or [])
    merged = []
    count = max(len(excel_widths), len(content_widths))
    for index in range(0, count):
        excel_w = excel_widths[index] if index < len(excel_widths) else TABLE_COLUMN_WIDTH
        content_w = content_widths[index] if index < len(content_widths) else TABLE_COLUMN_WIDTH
        excel_w = _clamp(excel_w, MIN_COL_WIDTH_FT, MAX_COL_WIDTH_FT)
        content_w = _clamp(content_w, MIN_COL_WIDTH_FT, MAX_COL_WIDTH_FT)
        # Weighted max: don't let Excel narrow columns crush long text, but
        # don't let a single long value explode the whole table either.
        weighted = (excel_w * ADAPTIVE_EXCEL_WEIGHT) + (content_w * ADAPTIVE_CONTENT_WEIGHT)
        merged.append(max(excel_w * 0.82, weighted))
    return merged


def _adaptive_column_widths(widths, cols):
    if not widths:
        return widths
    cleaned = []
    for width in widths:
        cleaned.append(_clamp(width, MIN_COL_WIDTH_FT, MAX_COL_WIDTH_FT))

    total = _table_total_width(cleaned)
    if total <= 0.0:
        return cleaned

    # Dynamic table target: wider tables are allowed to grow, but every import
    # remains inside a practical drafting-view width. This is generic and file-
    # independent: it only depends on number of columns and measured content.
    try:
        target = float(cols) * float(ADAPTIVE_TARGET_WIDTH_PER_COLUMN_FT)
    except Exception:
        target = ADAPTIVE_MIN_TABLE_WIDTH_FT
    target = _clamp(target, ADAPTIVE_MIN_TABLE_WIDTH_FT, ADAPTIVE_MAX_TABLE_WIDTH_FT)

    if total > target:
        scale = target / total
        scaled = []
        for width in cleaned:
            scaled.append(_clamp(float(width) * scale, MIN_COL_WIDTH_FT, MAX_COL_WIDTH_FT))
        cleaned = scaled
    return cleaned


def _calculate_column_widths(table_data, rows, cols, top_left_map, covered_map):
    widths = []
    for c in range(0, cols):
        max_len = 0
        for r in range(0, rows):
            if _is_covered_non_top_left(r, c, top_left_map, covered_map):
                continue
            value = _get_cell_value(table_data, r, c)
            if value:
                span = _get_cell_span(r, c, top_left_map)
                span_cols = max(1, int(span[3] - span[1] + 1))
                max_len = max(max_len, _ceil_div(len(value), span_cols))
        width = 0.32 + (float(min(max_len, 42)) * TABLE_CHAR_WIDTH)
        if width < TABLE_MIN_COLUMN_WIDTH:
            width = TABLE_MIN_COLUMN_WIDTH
        if width > TABLE_MAX_COLUMN_WIDTH:
            width = TABLE_MAX_COLUMN_WIDTH
        widths.append(width)
    return widths


def _calculate_row_heights(table_data, rows, cols, column_widths, top_left_map, covered_map):
    heights = []
    for r in range(0, rows):
        max_lines = 1
        for c in range(0, cols):
            if _is_covered_non_top_left(r, c, top_left_map, covered_map):
                continue
            value = _get_cell_value(table_data, r, c)
            if not value:
                continue
            span = _get_cell_span(r, c, top_left_map)
            span_width = _sum_range(column_widths, span[1], span[3])
            usable_width = span_width - (TABLE_TEXT_OFFSET_X * 2.0)
            chars_per_line = int(max(6.0, usable_width / TABLE_CHAR_WIDTH))
            lines = _ceil_div(len(value), chars_per_line)
            span_rows = max(1, int(span[2] - span[0] + 1))
            row_lines = _ceil_div(lines, span_rows)
            if row_lines > max_lines:
                max_lines = row_lines
        height = (TABLE_TEXT_OFFSET_Y * 2.0) + (float(max_lines) * TABLE_LINE_HEIGHT)
        if height < TABLE_MIN_ROW_HEIGHT:
            height = TABLE_MIN_ROW_HEIGHT
        if height > TABLE_MAX_ROW_HEIGHT:
            height = TABLE_MAX_ROW_HEIGHT
        heights.append(height)
    return heights


def _is_body_row_for_readability(table_data, row_index, top_left_map):
    if _row_header_like(table_data, row_index):
        return False
    try:
        row = table_data[row_index]
    except Exception:
        return True
    for col_index in range(0, len(row)):
        try:
            value = clean_display_text(row[col_index]).strip()
            if not value:
                continue
            if _cell_header_or_title_like(table_data, row_index, col_index, value, top_left_map):
                return False
        except Exception:
            pass
    return True


def _apply_readability_row_heights(table_data, row_heights, top_left_map):
    adjusted = []
    for index in range(0, len(row_heights)):
        try:
            height = float(row_heights[index])
        except Exception:
            height = TABLE_ROW_HEIGHT
        if _is_body_row_for_readability(table_data, index, top_left_map):
            try:
                height = max(height * BODY_ROW_HEIGHT_SCALE, MIN_BODY_ROW_HEIGHT_FT)
            except Exception:
                height = max(height, MIN_BODY_ROW_HEIGHT_FT)
        if height > MAX_ROW_HEIGHT_FT:
            height = MAX_ROW_HEIGHT_FT
        adjusted.append(height)
    return adjusted


def excel_column_width_to_feet(width):
    try:
        if width is None:
            feet = TABLE_COLUMN_WIDTH
        else:
            feet = float(width) * EXCEL_WIDTH_SCALE
    except Exception:
        feet = TABLE_COLUMN_WIDTH
    if feet < MIN_COL_WIDTH_FT:
        feet = MIN_COL_WIDTH_FT
    if feet > MAX_COL_WIDTH_FT:
        feet = MAX_COL_WIDTH_FT
    return feet


def excel_row_height_to_feet(height):
    try:
        if height is None:
            feet = TABLE_ROW_HEIGHT
        else:
            feet = (float(height) / 72.0) / 12.0
            feet = feet * EXCEL_ROW_HEIGHT_SCALE
    except Exception:
        feet = TABLE_ROW_HEIGHT
    if feet < MIN_ROW_HEIGHT_FT:
        feet = MIN_ROW_HEIGHT_FT
    if feet > MAX_ROW_HEIGHT_FT:
        feet = MAX_ROW_HEIGHT_FT
    return feet


def _has_dimension_values(values, expected_count):
    try:
        if values is None or len(values) < expected_count:
            return False
        for value in values[:expected_count]:
            if value is not None:
                return True
    except Exception:
        return False
    return False


def _debug_dimension_conversions(label, original_values, converted_values):
    if not DEBUG_OUTPUT:
        return
    try:
        count = min(5, len(converted_values))
        for index in range(0, count):
            original = None
            try:
                original = original_values[index]
            except Exception:
                pass
            print("Table Importer: %s %s Excel=%s Revit=%.3f ft" % (label, index + 1, safe_unicode(original), float(converted_values[index])))
    except Exception:
        pass


def _get_excel_column_widths(table_data, cols):
    try:
        raw_widths = getattr(table_data, "column_widths", None)
        if not _has_dimension_values(raw_widths, cols):
            return None
        widths = []
        for index in range(0, cols):
            widths.append(excel_column_width_to_feet(raw_widths[index]))
        _debug_dimension_conversions("column width", raw_widths, widths)
        return widths
    except Exception:
        return None


def _get_excel_row_heights(table_data, rows):
    try:
        raw_heights = getattr(table_data, "row_heights", None)
        if not _has_dimension_values(raw_heights, rows):
            return None
        heights = []
        for index in range(0, rows):
            heights.append(excel_row_height_to_feet(raw_heights[index]))
        _debug_dimension_conversions("row height", raw_heights, heights)
        return heights
    except Exception:
        return None


def _build_table_positions(origin_x, origin_y, column_widths, row_heights):
    x_positions = [origin_x]
    y_positions = [origin_y]
    for width in column_widths:
        x_positions.append(x_positions[-1] + width)
    for height in row_heights:
        y_positions.append(y_positions[-1] - height)
    return x_positions, y_positions


def _scale_dimension_values(values, fallback_value):
    scaled = []
    try:
        scale = float(TABLE_SCALE)
    except Exception:
        scale = 1.0
    for value in values:
        try:
            scaled.append(float(value) * scale)
        except Exception:
            scaled.append(float(fallback_value) * scale)
    return scaled


def _get_table_borders(table_data, rows, cols):
    try:
        if not getattr(table_data, "border_available", False):
            return None, False
        borders = getattr(table_data, "borders", None)
        if borders is None or len(borders) < rows:
            return None, False
        for r in range(0, rows):
            if borders[r] is None or len(borders[r]) < cols:
                return None, False
        return borders, True
    except Exception:
        return None, False


def _draw_full_grid(doc, view, x_positions, y_positions, top_left_map=None, entry=None):
    drawn_lines = set()
    rows = max(0, len(y_positions) - 1)
    cols = max(0, len(x_positions) - 1)
    for boundary_index in range(0, cols + 1):
        blocked_segments = _get_internal_merged_vertical_blocks(top_left_map, boundary_index, y_positions)
        _draw_vertical_segment_with_blocks(doc, view, drawn_lines, x_positions[boundary_index], y_positions[0], y_positions[rows], blocked_segments, entry)
    for boundary_index in range(0, rows + 1):
        blocked_segments = _get_internal_merged_horizontal_blocks(top_left_map, boundary_index, x_positions)
        _draw_horizontal_segment_with_blocks(doc, view, drawn_lines, x_positions[0], x_positions[cols], y_positions[boundary_index], blocked_segments, entry)


def _cell_border_value(borders, row_index, col_index, side):
    try:
        return bool(borders[row_index][col_index].get(side))
    except Exception:
        return False


def _horizontal_border_exists(borders, rows, cols, boundary_index, col_index):
    if col_index < 0 or col_index >= cols:
        return False
    if boundary_index <= 0:
        return _cell_border_value(borders, 0, col_index, "top")
    if boundary_index >= rows:
        return _cell_border_value(borders, rows - 1, col_index, "bottom")
    return (
        _cell_border_value(borders, boundary_index - 1, col_index, "bottom")
        or _cell_border_value(borders, boundary_index, col_index, "top")
    )


def _vertical_border_exists(borders, rows, cols, boundary_index, row_index):
    if row_index < 0 or row_index >= rows:
        return False
    if boundary_index <= 0:
        return _cell_border_value(borders, row_index, 0, "left")
    if boundary_index >= cols:
        return _cell_border_value(borders, row_index, cols - 1, "right")
    return (
        _cell_border_value(borders, row_index, boundary_index - 1, "right")
        or _cell_border_value(borders, row_index, boundary_index, "left")
    )


def _subtract_blocked_segments(start_value, end_value, blocked_segments):
    start_value = float(start_value)
    end_value = float(end_value)
    low_value = min(start_value, end_value)
    high_value = max(start_value, end_value)
    segments = [(low_value, high_value)]
    cleaned_blocks = []
    for block_start, block_end in blocked_segments or []:
        try:
            block_low = max(low_value, min(float(block_start), float(block_end)))
            block_high = min(high_value, max(float(block_start), float(block_end)))
            if block_high > block_low:
                cleaned_blocks.append((block_low, block_high))
        except Exception:
            pass
    cleaned_blocks.sort()
    for block_low, block_high in cleaned_blocks:
        next_segments = []
        for seg_low, seg_high in segments:
            if block_high <= seg_low or block_low >= seg_high:
                next_segments.append((seg_low, seg_high))
                continue
            if block_low > seg_low:
                next_segments.append((seg_low, block_low))
            if block_high < seg_high:
                next_segments.append((block_high, seg_high))
        segments = next_segments
    return segments


def _get_internal_merged_horizontal_blocks(top_left_map, boundary_index, x_positions):
    blocks = []
    try:
        for span in (top_left_map or {}).values():
            min_row, min_col, max_row, max_col = span
            if int(min_row) < int(boundary_index) and int(boundary_index) <= int(max_row):
                blocks.append((x_positions[int(min_col)], x_positions[int(max_col) + 1]))
    except Exception:
        pass
    return blocks


def _get_internal_merged_vertical_blocks(top_left_map, boundary_index, y_positions):
    blocks = []
    try:
        for span in (top_left_map or {}).values():
            min_row, min_col, max_row, max_col = span
            if int(min_col) < int(boundary_index) and int(boundary_index) <= int(max_col):
                blocks.append((y_positions[int(min_row)], y_positions[int(max_row) + 1]))
    except Exception:
        pass
    return blocks


def _draw_horizontal_segment_with_blocks(doc, view, drawn_lines, x1, x2, y, blocked_segments, entry=None):
    remaining_segments = _subtract_blocked_segments(x1, x2, blocked_segments)
    for seg_start, seg_end in remaining_segments:
        _make_line_once(doc, view, drawn_lines, seg_start, y, seg_end, y, entry)


def _draw_vertical_segment_with_blocks(doc, view, drawn_lines, x, y1, y2, blocked_segments, entry=None):
    remaining_segments = _subtract_blocked_segments(y1, y2, blocked_segments)
    for seg_start, seg_end in remaining_segments:
        _make_line_once(doc, view, drawn_lines, x, seg_end, x, seg_start, entry)


def _draw_horizontal_border_runs(doc, view, drawn_lines, borders, rows, cols, x_positions, y_positions, top_left_map=None, entry=None):
    for boundary_index in range(0, rows + 1):
        blocked_segments = _get_internal_merged_horizontal_blocks(top_left_map, boundary_index, x_positions)
        run_start = None
        for col_index in range(0, cols):
            if _horizontal_border_exists(borders, rows, cols, boundary_index, col_index):
                if run_start is None:
                    run_start = col_index
            else:
                if run_start is not None:
                    _draw_horizontal_segment_with_blocks(doc, view, drawn_lines, x_positions[run_start], x_positions[col_index], y_positions[boundary_index], blocked_segments, entry)
                    run_start = None
        if run_start is not None:
            _draw_horizontal_segment_with_blocks(doc, view, drawn_lines, x_positions[run_start], x_positions[cols], y_positions[boundary_index], blocked_segments, entry)


def _draw_vertical_border_runs(doc, view, drawn_lines, borders, rows, cols, x_positions, y_positions, top_left_map=None, entry=None):
    for boundary_index in range(0, cols + 1):
        blocked_segments = _get_internal_merged_vertical_blocks(top_left_map, boundary_index, y_positions)
        run_start = None
        for row_index in range(0, rows):
            if _vertical_border_exists(borders, rows, cols, boundary_index, row_index):
                if run_start is None:
                    run_start = row_index
            else:
                if run_start is not None:
                    _draw_vertical_segment_with_blocks(doc, view, drawn_lines, x_positions[boundary_index], y_positions[run_start], y_positions[row_index], blocked_segments, entry)
                    run_start = None
        if run_start is not None:
            _draw_vertical_segment_with_blocks(doc, view, drawn_lines, x_positions[boundary_index], y_positions[run_start], y_positions[rows], blocked_segments, entry)


def _draw_excel_borders(doc, view, borders, rows, cols, x_positions, y_positions, top_left_map=None, entry=None):
    drawn_lines = set()
    _draw_horizontal_border_runs(doc, view, drawn_lines, borders, rows, cols, x_positions, y_positions, top_left_map, entry)
    _draw_vertical_border_runs(doc, view, drawn_lines, borders, rows, cols, x_positions, y_positions, top_left_map, entry)


def _get_cell_rect(row_index, col_index, top_left_map, x_positions, y_positions):
    span = _get_cell_span(row_index, col_index, top_left_map)
    min_row, min_col, max_row, max_col = span
    return (
        x_positions[min_col],
        y_positions[min_row],
        x_positions[max_col + 1],
        y_positions[max_row + 1],
    )


def _print_text_width_warning_once(message):
    global TEXT_NOTE_WIDTH_WARNING_PRINTED
    try:
        if TEXT_NOTE_WIDTH_WARNING_PRINTED:
            return
        TEXT_NOTE_WIDTH_WARNING_PRINTED = True
        print("Table Importer: %s" % safe_unicode(message))
    except Exception:
        pass


def _set_text_note_width(text_note, width):
    try:
        text_note.Width = float(width)
        return True
    except Exception:
        pass
    try:
        text_note.SetBoxWidth(float(width))
        return True
    except Exception:
        pass
    return False


def create_text_note_in_cell(doc, view, point, width, text, text_note_type_id, horizontal="left", vertical="middle"):
    if not USE_WRAPPED_TEXT_NOTES:
        try:
            options = DB.TextNoteOptions(text_note_type_id)
            options = _set_text_note_alignment_options(options, horizontal, vertical)
            return TextNote.Create(doc, view.Id, point, text, options)
        except Exception:
            pass
        try:
            note = TextNote.Create(doc, view.Id, point, text, text_note_type_id)
            try:
                if horizontal == "center":
                    note.HorizontalAlignment = DB.HorizontalTextAlignment.Center
                elif horizontal == "right":
                    note.HorizontalAlignment = DB.HorizontalTextAlignment.Right
            except Exception:
                pass
            return note
        except Exception as ex:
            _print_text_width_warning_once("could not create point-based TextNote: %s" % safe_unicode(ex))
            raise

    try:
        constrained_width = float(width)
    except Exception:
        constrained_width = 0.10
    if constrained_width < 0.10:
        constrained_width = 0.10

    try:
        options = DB.TextNoteOptions(text_note_type_id)
        options = _set_text_note_alignment_options(options, horizontal, vertical)
        return TextNote.Create(doc, view.Id, point, constrained_width, text, options)
    except Exception:
        pass

    try:
        return TextNote.Create(doc, view.Id, point, constrained_width, text, text_note_type_id)
    except Exception:
        pass

    text_note = None
    try:
        text_note = TextNote.Create(doc, view.Id, point, text, text_note_type_id)
    except Exception as ex:
        _print_text_width_warning_once("could not create TextNote with constrained width: %s" % safe_unicode(ex))
        raise

    if not _set_text_note_width(text_note, constrained_width):
        _print_text_width_warning_once("Revit TextNote width could not be constrained in this API version.")
    return text_note

def _format_first_values(values, count):
    result = []
    try:
        limit = min(count, len(values))
        for index in range(0, limit):
            result.append("%.3f" % float(values[index]))
    except Exception:
        pass
    return ", ".join(result)


def _debug_geometry_report(rows, cols, column_widths, row_heights, width_source, height_source):
    if not DEBUG_OUTPUT:
        return
    try:
        total_width = 0.0
        for width in column_widths:
            total_width += float(width)
        total_height = 0.0
        for height in row_heights:
            total_height += float(height)
        print("Table Importer geometry: rows=%s cols=%s widths=%s heights=%s" % (rows, cols, width_source, height_source))
        print("Table Importer geometry: first 10 column widths ft: %s" % _format_first_values(column_widths, 10))
        print("Table Importer geometry: first 10 row heights ft: %s" % _format_first_values(row_heights, 10))
        print("Table Importer geometry: total width %.3f ft, total height %.3f ft" % (total_width, total_height))
    except Exception:
        pass


def draw_table_in_view(view, table_data, entry=None):
    doc = get_revit_document()
    if doc is None:
        raise Exception("No active Revit document.")

    rows = len(table_data)
    cols = 0
    for row in table_data:
        try:
            cols = max(cols, len(row))
        except Exception:
            pass
    if rows <= 0 or cols <= 0:
        raise Exception("Excel region has no readable cells.")

    origin_x = 0.0
    origin_y = 0.0
    if entry is not None:
        ensure_table_entry_uid(entry)

    top_left_map, covered_map = _build_merged_cell_maps(table_data)
    excel_column_widths = _get_excel_column_widths(table_data, cols)
    content_column_widths = _calculate_column_widths(table_data, rows, cols, top_left_map, covered_map)
    width_source = "Excel+content+adaptive"
    if excel_column_widths is None:
        width_source = "content+adaptive"
        column_widths = content_column_widths
    else:
        column_widths = _merge_excel_and_content_column_widths(excel_column_widths, content_column_widths)
    column_widths = _adaptive_column_widths(column_widths, cols)
    row_heights = _get_excel_row_heights(table_data, rows)
    height_source = "Excel"
    content_row_heights = _calculate_row_heights(table_data, rows, cols, column_widths, top_left_map, covered_map)
    if row_heights is None:
        height_source = "fallback"
        row_heights = content_row_heights
    else:
        # Excel row heights are often too small for Revit TextNotes. Keep the
        # Excel proportions, but never allow text to spill into adjacent rows.
        adjusted_heights = []
        for index in range(0, rows):
            try:
                adjusted_heights.append(max(float(row_heights[index]), float(content_row_heights[index]) * 0.90))
            except Exception:
                adjusted_heights.append(row_heights[index])
        row_heights = adjusted_heights
        height_source = "Excel+fit"
    row_heights = _apply_readability_row_heights(table_data, row_heights, top_left_map)
    column_widths = _scale_dimension_values(column_widths, TABLE_COLUMN_WIDTH)
    row_heights = _scale_dimension_values(row_heights, TABLE_ROW_HEIGHT)
    x_positions, y_positions = _build_table_positions(origin_x, origin_y, column_widths, row_heights)
    _debug_geometry_report(rows, cols, column_widths, row_heights, width_source, height_source)

    borders, border_available = _get_table_borders(table_data, rows, cols)
    if border_available:
        _draw_excel_borders(doc, view, borders, rows, cols, x_positions, y_positions, top_left_map, entry)
    else:
        _draw_full_grid(doc, view, x_positions, y_positions, top_left_map, entry)

    default_note_type = get_table_text_note_type(doc, get_table_text_font(table_data))
    debug_text_fit_count = [0]
    for r in range(0, rows):
        for c in range(0, cols):
            if _is_covered_non_top_left(r, c, top_left_map, covered_map):
                continue
            value = _get_cell_value(table_data, r, c)
            if not value:
                continue
            x1, y1, x2, y2 = _get_cell_rect(r, c, top_left_map, x_positions, y_positions)
            cell_width = x2 - x1
            note_type = get_text_note_type_for_cell(doc, table_data, r, c, default_note_type)
            note_type_id = note_type.Id if note_type is not None else DB.ElementId.InvalidElementId
            text_size = get_text_note_type_size(note_type)
            if USE_WRAPPED_TEXT_NOTES:
                display_value = clean_display_text(value)
            else:
                display_value = fit_text_to_cell(value, cell_width, text_size)
            horizontal = get_cell_horizontal_alignment(table_data, r, c, value, cell_width, top_left_map)
            vertical = get_cell_vertical_alignment(table_data, r, c)
            point = get_text_point_for_alignment(x1, y1, x2, y2, horizontal, vertical, text_size)
            _debug_fit_text(value, display_value, cell_width, text_size, debug_text_fit_count)
            text_note = create_text_note_in_cell(doc, view, point, max(cell_width - (2.0 * TEXT_PADDING_X), 0.10), display_value, note_type_id, horizontal, vertical)
            if entry is not None:
                apply_table_importer_tag(text_note, entry)
                remember_created_element(entry, text_note)


class AddTableDialog(object):
    def __init__(self, owner=None, existing_names=None):
        self.xaml_path = os.path.join(script_dir, "AddTableWindow.xaml")
        self.window = load_xaml(self.xaml_path)
        if owner:
            self.window.Owner = owner

        self.file_paths = []
        self.result_entries = []
        self.active_file_path = None
        self.existing_names = list(existing_names or [])

        self.FilePathTextBox = self.window.FindName("FilePathTextBox")
        self.FilesListBox = self.window.FindName("FilesListBox")
        self.NoFilesTextBlock = self.window.FindName("NoFilesTextBlock")
        self.BrowseButton = self.window.FindName("BrowseButton")
        self.WorksheetComboBox = self.window.FindName("WorksheetComboBox")
        self.RegionComboBox = self.window.FindName("RegionComboBox")
        self.ImportTypeComboBox = self.window.FindName("ImportTypeComboBox")
        self.ViewTypeComboBox = self.window.FindName("ViewTypeComboBox")
        self.ViewScaleComboBox = self.window.FindName("ViewScaleComboBox")
        self.DpiComboBox = self.window.FindName("DpiComboBox")
        self.BlackAndWhiteCheckBox = self.window.FindName("BlackAndWhiteCheckBox")
        self.AutoSyncCheckBox = self.window.FindName("AutoSyncCheckBox")
        self.CopiesTextBox = self.window.FindName("CopiesTextBox")
        self.StatusTextBlock = self.window.FindName("StatusTextBlock")
        self.OkButton = self.window.FindName("OkButton")
        self.CancelButton = self.window.FindName("CancelButton")
        self.LogoImage = self.window.FindName("LogoImage")
        self.VersionTextBlock = self.window.FindName("VersionTextBlock")
        if self.VersionTextBlock:
            self.VersionTextBlock.Text = TOOL_VERSION_LABEL

        self.setup_defaults()
        self.bind_events()

    def setup_defaults(self):
        fill_combo(self.RegionComboBox, [USED_RANGE_DISPLAY], USED_RANGE_DISPLAY)
        fill_combo(self.ImportTypeComboBox, DEFAULT_IMPORT_TYPES, "Excel Link")
        fill_combo(self.ViewTypeComboBox, DEFAULT_VIEW_TYPES, "Drafting View")
        fill_combo(self.ViewScaleComboBox, DEFAULT_SCALES, "1")
        fill_combo(self.DpiComboBox, DEFAULT_DPI, "150")
        self.BlackAndWhiteCheckBox.IsChecked = True
        self.AutoSyncCheckBox.IsChecked = False
        self.CopiesTextBox.Text = "1"
        load_logo_into_image(self.LogoImage)
        self.update_file_controls()
        self.OkButton.IsEnabled = False



    def is_valid_excel_file(self, file_path):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            return os.path.exists(file_path) and ext in (".xlsx", ".xlsm", ".xls")
        except Exception:
            return False

    def get_valid_file_paths(self):
        result = []
        for path in self.file_paths:
            if self.is_valid_excel_file(path):
                result.append(path)
        return result

    def update_file_controls(self):
        valid_count = len(self.get_valid_file_paths())
        self.OkButton.IsEnabled = valid_count > 0
        if self.NoFilesTextBlock is not None:
            self.NoFilesTextBlock.Visibility = Visibility.Visible if valid_count == 0 else Visibility.Collapsed

    def add_file_list_item(self, file_path):
        item = ListBoxItem()
        item.Tag = file_path
        item.ToolTip = file_path
        panel = StackPanel()
        name_text = TextBlock()
        name_text.Text = os.path.basename(file_path)
        name_text.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
        path_text = TextBlock()
        path_text.Text = file_path
        path_text.FontSize = 11
        path_text.Foreground = System.Windows.Media.Brushes.Gray
        path_text.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
        panel.Children.Add(name_text)
        panel.Children.Add(path_text)
        item.Content = panel
        self.FilesListBox.Items.Add(item)

    def get_selected_file_path(self):
        try:
            item = self.FilesListBox.SelectedItem
            if item is not None and item.Tag is not None:
                return safe_unicode(item.Tag)
        except Exception:
            pass
        return u""
    def bind_events(self):
        self.BrowseButton.Click += self.on_browse
        self.FilesListBox.SelectionChanged += self.on_file_selected
        self.WorksheetComboBox.SelectionChanged += self.on_worksheet_changed
        self.OkButton.Click += self.on_ok
        self.CancelButton.Click += self.on_cancel

    def on_browse(self, sender, args):
        dialog = OpenFileDialog()
        dialog.Title = "Select Excel files"
        dialog.Filter = EXCEL_FILTER
        dialog.Multiselect = True
        result = dialog.ShowDialog()
        if result == DialogResult.OK:
            self.file_paths = [safe_unicode(x) for x in dialog.FileNames]
            self.refresh_file_list()

    def refresh_file_list(self):
        self.FilesListBox.Items.Clear()
        valid_paths = self.get_valid_file_paths()
        for path in valid_paths:
            self.add_file_list_item(path)
        self.update_file_controls()
        if valid_paths:
            self.FilePathTextBox.Text = "%s valid file(s) selected" % len(valid_paths)
            self.FilesListBox.SelectedIndex = 0
            self.load_worksheets(valid_paths[0])
            self.StatusTextBlock.Text = "%s file(s) loaded." % len(valid_paths)
        else:
            self.FilePathTextBox.Text = ""
            fill_combo(self.WorksheetComboBox, [], None)
            fill_combo(self.RegionComboBox, [USED_RANGE_DISPLAY], USED_RANGE_DISPLAY)
            self.StatusTextBlock.Text = "No files selected."

    def on_file_selected(self, sender, args):
        try:
            path = self.get_selected_file_path()
            if path:
                self.load_worksheets(path)
        except Exception:
            pass

    def load_worksheets(self, file_path):
        self.active_file_path = file_path
        worksheets = get_excel_worksheets(file_path)
        if not worksheets:
            self.WorksheetComboBox.Items.Clear()
            fill_combo(self.RegionComboBox, [USED_RANGE_DISPLAY], USED_RANGE_DISPLAY)
            self.StatusTextBlock.Text = "No worksheets found or Excel could not be opened."
            return
        fill_combo(self.WorksheetComboBox, worksheets, worksheets[0])
        self.load_region_for_selected_worksheet()
        self.StatusTextBlock.Text = "%s worksheet(s) loaded for %s." % (len(worksheets), os.path.basename(file_path))

    def on_worksheet_changed(self, sender, args):
        self.load_region_for_selected_worksheet()

    def load_region_for_selected_worksheet(self):
        worksheet = get_combo_text(self.WorksheetComboBox)
        if not self.active_file_path or not worksheet:
            fill_combo(self.RegionComboBox, [USED_RANGE_DISPLAY], USED_RANGE_DISPLAY)
            return
        regions = get_excel_regions(self.active_file_path, worksheet)
        if not regions:
            regions = [USED_RANGE_DISPLAY]
        fill_combo(self.RegionComboBox, regions, regions[0])

    def get_copies(self):
        try:
            value = int(self.CopiesTextBox.Text)
            if value < 1:
                return 1
            return value
        except Exception:
            return 1

    def get_sheet_for_file(self, file_path, requested_sheet):
        sheets = get_excel_worksheets(file_path)
        if requested_sheet in sheets:
            return requested_sheet
        if sheets:
            return sheets[0]
        return u""

    def get_region_for_file(self, file_path, worksheet, requested_region):
        regions = get_excel_regions(file_path, worksheet)
        if not regions:
            regions = [USED_RANGE_DISPLAY]
        if requested_region in regions:
            return requested_region, regions
        return regions[0], regions

    def on_ok(self, sender, args):
        valid_paths = self.get_valid_file_paths()
        if not valid_paths:
            self.StatusTextBlock.Text = "Please select one or more Excel files."
            return

        requested_sheet = get_combo_text(self.WorksheetComboBox)
        requested_region = get_combo_text(self.RegionComboBox) or USED_RANGE_DISPLAY
        import_type = get_combo_text(self.ImportTypeComboBox) or "Excel Link"
        view_type = get_combo_text(self.ViewTypeComboBox) or "Drafting View"
        view_scale = get_combo_text(self.ViewScaleComboBox) or "1"
        dpi = get_combo_text(self.DpiComboBox) or "150"
        black_white = safe_bool(self.BlackAndWhiteCheckBox.IsChecked)
        auto_sync = safe_bool(self.AutoSyncCheckBox.IsChecked)
        copies = self.get_copies()

        entries = []
        for file_path in valid_paths:
            if not file_path or not os.path.exists(file_path):
                continue
            worksheet = self.get_sheet_for_file(file_path, requested_sheet)
            if not worksheet:
                continue
            region, regions = self.get_region_for_file(file_path, worksheet, requested_region)
            for index in range(copies):
                view_name = get_default_view_name(file_path, worksheet, region, self.existing_names)
                self.existing_names.append(view_name)
                entries.append(TableEntry(
                    selected=True,
                    status=ROW_READY_STATUS,
                    source=clean_display_text(os.path.basename(file_path)),
                    import_type=import_type,
                    view_name=view_name,
                    dpi=dpi,
                    auto_sync=auto_sync,
                    black_and_white=black_white,
                    last_modified=get_last_modified(file_path),
                    worksheet=worksheet,
                    region=region,
                    region_options=list(regions),
                    view_type=view_type,
                    view_scale=view_scale,
                    file_path=file_path,
                    path_mode="Absolute",
                    revit_view_id=None,
                ))

        if not entries:
            self.StatusTextBlock.Text = "No valid table rows could be created."
            return

        self.result_entries = entries
        self.window.DialogResult = True
        self.window.Close()

    def on_cancel(self, sender, args):
        self.result_entries = []
        self.window.DialogResult = False
        self.window.Close()

    def show_dialog(self):
        self.window.ShowDialog()
        return self.result_entries


class TableImporterWindow(object):
    def __init__(self):
        self.xaml_path = os.path.join(script_dir, "TableImporter.xaml")
        self.window = load_xaml(self.xaml_path)
        self.all_entries = []
        self.entries = ObservableCollection[object]()
        self.current_search = u""
        self.pending_batch_action = None
        self.pending_batch_targets = []
        self._is_updating_selection = False

        self.TablesDataGrid = self.window.FindName("TablesDataGrid")
        self.AddTablesButton = self.window.FindName("AddTablesButton")
        self.EmptyStateButton = self.window.FindName("EmptyStateButton")
        self.EmptyStatePanel = self.window.FindName("EmptyStatePanel")
        self.BatchActionsButton = self.window.FindName("BatchActionsButton")
        self.RefreshButton = self.window.FindName("RefreshButton")
        self.SearchTextBox = self.window.FindName("SearchTextBox")
        self.FooterStatusTextBlock = self.window.FindName("FooterStatusTextBlock")
        self.CompletedTextBlock = self.window.FindName("CompletedTextBlock")
        self.CloseButton = self.window.FindName("CloseButton")
        self.ApplyButton = self.window.FindName("ApplyButton")
        self.LogoImage = self.window.FindName("LogoImage")
        self.StatusTextBlock = self.window.FindName("StatusTextBlock")
        self.ProgressTextBlock = self.window.FindName("ProgressTextBlock")
        self.ApplyProgressBar = self.window.FindName("ApplyProgressBar")
        self.ProgressDetailTextBlock = self.window.FindName("ProgressDetailTextBlock")
        self.VersionTextBlock = self.window.FindName("VersionTextBlock")
        if self.VersionTextBlock:
            self.VersionTextBlock.Text = TOOL_VERSION_LABEL

        self.BatchActionsContextMenu = self.BatchActionsButton.ContextMenu
        self.RowActionsContextMenu = self.TablesDataGrid.ContextMenu
        self.bind_menu_items()
        self.load_logo_image()
        self.load_saved_entries()
        self.set_apply_progress("Ready", "", 100)
        self.set_progress_detail("Ready.")
        self.bind_events()
        # Auto-refresh file status on open so LastModified is current.
        try:
            self.on_refresh(None, None)
            self.FooterStatusTextBlock.Text = ""
            self.update_footer()
        except Exception:
            pass

    def bind_menu_items(self):
        names = [
            "UpdateViews", "DuplicateViews", "ReloadFrom", "ResetLegacyContent", "AbsolutePath", "RelativePath",
            "OpenFiles", "OpenFolders", "DeleteViews", "UnlinkView", "OpenView"
        ]
        for name in names:
            setattr(self, "Batch%sMenuItem" % name, self.window.FindName("Batch%sMenuItem" % name))
            setattr(self, "Row%sMenuItem" % name, self.window.FindName("Row%sMenuItem" % name))

    def load_logo_image(self):
        load_logo_into_image(self.LogoImage)

    def refresh_dispatcher(self):
        try:
            priority = System.Windows.Threading.DispatcherPriority.Background
            self.window.Dispatcher.Invoke(priority, System.Action(lambda: None))
            return
        except Exception:
            pass
        try:
            self.window.Dispatcher.Invoke(System.Action(lambda: None))
        except Exception:
            pass

    def set_apply_progress(self, status_text, progress_label, percent):
        try:
            percent_value = int(percent)
        except Exception:
            percent_value = 0
        if percent_value < 0:
            percent_value = 0
        if percent_value > 100:
            percent_value = 100
        try:
            if self.StatusTextBlock:
                self.StatusTextBlock.Text = normalize_status_label(status_text)
        except Exception:
            pass
        try:
            if self.ProgressTextBlock:
                self.ProgressTextBlock.Text = "%s%%" % percent_value
        except Exception:
            pass
        try:
            if self.ApplyProgressBar:
                self.ApplyProgressBar.Value = percent_value
        except Exception:
            pass
        self.refresh_dispatcher()

    def set_progress_detail(self, detail_text):
        try:
            if self.ProgressDetailTextBlock:
                self.ProgressDetailTextBlock.Text = safe_unicode(detail_text)
        except Exception:
            pass
        self.refresh_dispatcher()

    def format_progress_name(self, entry):
        name = safe_unicode(getattr(entry, "ViewName", ""))
        if not name:
            name = safe_unicode(getattr(entry, "SourceName", ""))
        if len(name) > 72:
            name = name[:69] + "..."
        return name

    def update_apply_progress(self, processed, total):
        try:
            if total <= 0:
                percent = 0
            else:
                percent = int((float(processed) / float(total)) * 100.0)
        except Exception:
            percent = 0
        self.set_apply_progress("Updating... %s/%s" % (processed, total), "Updating", percent)

    def bind_events(self):
        self.AddTablesButton.Click += self.on_add_tables
        self.EmptyStateButton.Click += self.on_add_tables
        self.RefreshButton.Click += self.on_refresh
        self.CloseButton.Click += self.on_close
        self.ApplyButton.Click += self.on_apply
        self.BatchActionsButton.Click += self.on_batch_actions
        self.SearchTextBox.TextChanged += self.on_search_changed
        self.TablesDataGrid.PreviewMouseLeftButtonDown += self.on_datagrid_single_click
        self.TablesDataGrid.CellEditEnding += self.on_cell_edit_ending
        self.TablesDataGrid.Drop += self.on_drop_files
        try:
            self.EmptyStatePanel.Drop += self.on_drop_files
            self.EmptyStateButton.Drop += self.on_drop_files
        except Exception:
            pass

        for prefix in ("Batch", "Row"):
            menu_map = [
                ("UpdateViewsMenuItem", self.on_update_views),
                ("DuplicateViewsMenuItem", self.on_duplicate_views),
                ("ReloadFromMenuItem", self.on_reload_from),
                ("ResetLegacyContentMenuItem", self.on_reset_legacy_content),
                ("AbsolutePathMenuItem", self.on_absolute_path),
                ("RelativePathMenuItem", self.on_relative_path),
                ("OpenFilesMenuItem", self.on_open_files),
                ("OpenFoldersMenuItem", self.on_open_folders),
                ("DeleteViewsMenuItem", self.on_delete_views),
                ("UnlinkViewMenuItem", self.on_unlink_view),
                ("OpenViewMenuItem", self.on_open_view),
            ]
            for suffix, handler in menu_map:
                try:
                    item = self.window.FindName(prefix + suffix)
                    if item is not None:
                        item.Click += handler
                except Exception:
                    pass

    def on_datagrid_single_click(self, sender, args):
        try:
            if find_visual_parent(args.OriginalSource, CheckBox) is not None:
                return
            row = find_visual_parent(args.OriginalSource, DataGridRow)
            if row is not None and not row.IsEditing:
                self.TablesDataGrid.BeginEdit()
        except Exception:
            pass

    def on_cell_edit_ending(self, sender, args):
        try:
            if args.EditAction.ToString() == "Commit":
                try:
                    header = safe_unicode(args.Column.Header)
                except Exception:
                    header = u""
                save_entries(self.all_entries)
                self.update_selection_state()
                if header == u"✓":
                    return
                self.FooterStatusTextBlock.Text = "Changes saved."
        except Exception:
            pass

    def load_saved_entries(self):
        self.all_entries = []
        known_names = get_existing_revit_view_names()
        for entry in load_entries():
            self.prepare_entry(entry)
            try:
                if not clean_display_text(entry.ViewName).strip():
                    entry.ViewName = get_default_view_name(entry.FilePath, entry.Worksheet, entry.Region, known_names)
                known_names.append(entry.ViewName)
            except Exception:
                pass
            self.all_entries.append(entry)
        self.apply_filter()

    def prepare_entry(self, entry):
        try:
            ensure_table_entry_uid(entry)
            if not entry.Source:
                entry.Source = clean_display_text(os.path.basename(resolve_entry_path(entry.FilePath)))
            else:
                entry.Source = clean_display_text(entry.Source)
            entry.ViewName = clean_display_text(entry.ViewName)
            entry.Status = normalize_row_status(entry.Status)
            entry.ViewType = get_view_type_display_name(entry.ViewType)
            try:
                entry.remove_PropertyChanged(self.on_entry_property_changed)
            except Exception:
                pass
            try:
                entry.add_PropertyChanged(self.on_entry_property_changed)
            except Exception:
                pass
        except Exception:
            pass
        self.populate_worksheet_options(entry)
        self.populate_region_options(entry)

    def on_entry_property_changed(self, sender, args):
        try:
            prop_name = safe_unicode(args.PropertyName)
        except Exception:
            prop_name = u""
        if prop_name == "Selected":
            self.update_selection_state()
        elif prop_name == "Worksheet":
            try:
                self.populate_region_options(sender)
                save_entries(self.all_entries)
            except Exception:
                pass

    def populate_worksheet_options(self, entry):
        try:
            path = resolve_entry_path(entry.FilePath)
            if path and os.path.exists(path):
                options = get_excel_worksheets(path)
            else:
                options = []
            current = safe_unicode(entry.Worksheet)
            if current and current not in options:
                options.insert(0, current)
            entry.WorksheetOptions = options
            if options and not entry.Worksheet:
                entry.Worksheet = options[0]
        except Exception:
            try:
                entry.WorksheetOptions = [entry.Worksheet] if entry.Worksheet else []
            except Exception:
                pass

    def populate_region_options(self, entry):
        try:
            path = resolve_entry_path(entry.FilePath)
            if path and os.path.exists(path) and entry.Worksheet:
                options = get_excel_regions(path, entry.Worksheet)
            else:
                options = []
            current = safe_unicode(entry.Region)
            if current and current not in options:
                options.insert(0, current)
            if not options:
                options = [USED_RANGE_DISPLAY]
            entry.RegionOptions = options
            if not entry.Region:
                entry.Region = options[0]
        except Exception:
            try:
                entry.RegionOptions = [entry.Region or USED_RANGE_DISPLAY]
            except Exception:
                pass

    def commit_pending_grid_edits(self):
        try:
            if self.TablesDataGrid:
                self.TablesDataGrid.CommitEdit(DataGridEditingUnit.Cell, True)
                self.TablesDataGrid.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass

    def save_current_entries(self):
        self.commit_pending_grid_edits()
        for entry in self.all_entries:
            try:
                ensure_table_entry_uid(entry)
            except Exception:
                pass
        save_entries(self.all_entries)

    def update_footer(self):
        total = len(self.all_entries)
        selected = len(self.get_checked_entries())
        drafting = 0
        legend = 0
        schedule = 0
        for entry in self.all_entries:
            try:
                view_type = normalize_view_type_name(entry.ViewType)
                if view_type == "Legend View":
                    legend += 1
                elif view_type == "Schedule View":
                    schedule += 1
                else:
                    drafting += 1
            except Exception:
                pass
        self.FooterStatusTextBlock.Text = "Total %s | Legends %s | Schedules %s | Drafting %s | Selected %s" % (total, legend, schedule, drafting, selected)
        linked = 0
        for entry in self.all_entries:
            try:
                if entry.RevitViewId:
                    linked += 1
            except Exception:
                pass
        pct = 0
        if total:
            pct = int((float(linked) / float(total)) * 100.0)
        if self.CompletedTextBlock:
            self.CompletedTextBlock.Text = "Completed %s%%" % pct
        self.update_empty_state()
        self.update_apply_button_state(selected)

    def get_checked_entries(self):
        result = []
        try:
            for entry in self.all_entries:
                if entry.Selected:
                    result.append(entry)
        except Exception:
            pass
        return result

    def get_grid_selected_entries(self):
        result = []
        try:
            for item in self.TablesDataGrid.SelectedItems:
                if item is not None:
                    result.append(item)
        except Exception:
            pass
        return result

    def get_action_target_entries(self):
        result = []
        seen = {}
        for source in (self.get_checked_entries(), self.get_grid_selected_entries()):
            for entry in source:
                key = id(entry)
                if key in seen:
                    continue
                seen[key] = True
                result.append(entry)
        return result

    def is_processable_entry(self, entry):
        try:
            if normalize_view_type_name(entry.ViewType) in ("Legend View", "Schedule View"):
                return True
            path = resolve_entry_path(entry.FilePath)
            if not path or not os.path.exists(path):
                return False
            if not clean_display_text(entry.Worksheet).strip():
                return False
            if not clean_display_text(entry.Region).strip():
                return False
            return True
        except Exception:
            return False

    def update_apply_button_state(self, selected_count=None):
        if selected_count is None:
            selected_count = len(self.get_checked_entries())
        enabled = False
        if selected_count > 0:
            for entry in self.get_checked_entries():
                if self.is_processable_entry(entry):
                    enabled = True
                    break
        try:
            if self.ApplyButton:
                self.ApplyButton.IsEnabled = enabled
        except Exception:
            pass

    def update_selection_state(self):
        if self._is_updating_selection:
            return
        self._is_updating_selection = True
        try:
            selected = len(self.get_checked_entries())
            self.update_apply_button_state(selected)
            if selected <= 0:
                self.FooterStatusTextBlock.Text = "No rows selected."
            else:
                self.FooterStatusTextBlock.Text = "%s row(s) selected. Press Apply to update views." % selected
            if DEBUG_OUTPUT:
                try:
                    print("Table Importer: selection changed, selected=%s." % selected)
                except Exception:
                    pass
        finally:
            self._is_updating_selection = False

    def update_empty_state(self):
        if self.EmptyStatePanel is None:
            return
        if self.entries.Count == 0:
            self.EmptyStatePanel.Visibility = Visibility.Visible
        else:
            self.EmptyStatePanel.Visibility = Visibility.Collapsed

    def matches_search(self, entry, text):
        if not text:
            return True
        values = [entry.ViewName, entry.Worksheet, entry.Region, entry.FilePath]
        for value in values:
            if text in safe_unicode(value).lower():
                return True
        return False

    def apply_filter(self):
        self.entries.Clear()
        text = safe_unicode(self.current_search).lower().strip()
        for entry in self.all_entries:
            if self.matches_search(entry, text):
                self.entries.Add(entry)
        self.TablesDataGrid.ItemsSource = self.entries
        self.TablesDataGrid.Items.Refresh()
        self.update_footer()

    def get_known_view_names(self):
        names = []
        for entry in self.all_entries:
            try:
                value = clean_display_text(entry.ViewName).strip()
                if value:
                    names.append(value)
            except Exception:
                pass
        try:
            names.extend(get_existing_revit_view_names())
        except Exception:
            pass
        return names

    def add_entries(self, new_entries):
        known_names = self.get_known_view_names()
        for entry in new_entries:
            try:
                if not clean_display_text(entry.ViewName).strip():
                    entry.ViewName = get_default_view_name(entry.FilePath, entry.Worksheet, entry.Region, known_names)
                    known_names.append(entry.ViewName)
            except Exception:
                pass
            self.prepare_entry(entry)
            self.all_entries.append(entry)
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "%s row(s) ready to update. Press Apply to update views." % len(new_entries)

    def on_add_tables(self, sender, args):
        dialog = AddTableDialog(owner=self.window, existing_names=self.get_known_view_names())
        entries = dialog.show_dialog()
        if entries:
            self.add_entries(entries)

    def on_drop_files(self, sender, args):
        try:
            if not args.Data.GetDataPresent(System.Windows.DataFormats.FileDrop):
                return
            paths = list(args.Data.GetData(System.Windows.DataFormats.FileDrop))
            excel_paths = []
            for path in paths:
                ext = os.path.splitext(path)[1].lower()
                if ext in (".xlsx", ".xlsm", ".xls"):
                    excel_paths.append(path)
            if not excel_paths:
                self.FooterStatusTextBlock.Text = "Drop Excel files only."
                return
            entries = []
            known_names = self.get_known_view_names()
            for path in excel_paths:
                sheets = get_excel_worksheets(path)
                if not sheets:
                    continue
                sheet = sheets[0]
                regions = get_excel_regions(path, sheet)
                region = regions[0] if regions else USED_RANGE_DISPLAY
                view_name = get_default_view_name(path, sheet, region, known_names)
                known_names.append(view_name)
                entries.append(TableEntry(
                    selected=True,
                    status=ROW_READY_STATUS,
                    source=clean_display_text(os.path.basename(path)),
                    import_type="Excel Link",
                    view_name=view_name,
                    dpi="150",
                    auto_sync=False,
                    black_and_white=True,
                    last_modified=get_last_modified(path),
                    worksheet=sheet,
                    region=region,
                    region_options=regions,
                    view_type="Drafting View",
                    view_scale="1",
                    file_path=path,
                    path_mode="Absolute",
                    revit_view_id=None,
                ))
            if entries:
                self.add_entries(entries)
        except Exception as ex:
            self.FooterStatusTextBlock.Text = "Drop failed: %s" % safe_unicode(ex)

    def on_refresh(self, sender, args):
        updated = 0
        missing = 0
        for entry in self.all_entries:
            path = resolve_entry_path(entry.FilePath)
            if not path or not os.path.exists(path):
                entry.Status = "Missing File"
                entry.LastModified = ""
                self.populate_worksheet_options(entry)
                self.populate_region_options(entry)
                missing += 1
                continue
            self.populate_worksheet_options(entry)
            self.populate_region_options(entry)
            old_date = entry.LastModified
            new_date = get_last_modified(path)
            entry.LastModified = new_date
            entry.Source = clean_display_text(os.path.basename(path))
            if old_date and new_date and old_date != new_date:
                entry.Status = ROW_READY_STATUS
                updated += 1
            elif entry.RevitViewId:
                entry.Status = "Updated"
            else:
                entry.Status = ROW_READY_STATUS
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "Refresh complete. %s modified, %s missing." % (updated, missing)

    def on_apply(self, sender, args):
        try:
            self.update_selected_views(True)
        finally:
            self.pending_batch_action = None
            self.pending_batch_targets = []

    def on_batch_actions(self, sender, args):
        try:
            self.BatchActionsContextMenu.PlacementTarget = self.BatchActionsButton
            self.BatchActionsContextMenu.IsOpen = True
        except Exception:
            self.FooterStatusTextBlock.Text = "Batch Actions menu could not be opened."

    def get_target_entries(self, allow_selected_rows=True):
        if self.pending_batch_action == "Update Views" and self.pending_batch_targets:
            return list(self.pending_batch_targets)
        if allow_selected_rows:
            return self.get_action_target_entries()
        return self.get_checked_entries()

    def require_targets(self):
        targets = self.get_target_entries(True)
        if not targets:
            self.FooterStatusTextBlock.Text = "No rows selected."
        return targets

    def on_update_views(self, sender, args):
        targets = self.get_action_target_entries()
        if not targets:
            self.FooterStatusTextBlock.Text = "No rows selected."
            return
        self.pending_batch_action = "Update Views"
        self.pending_batch_targets = list(targets)
        self.FooterStatusTextBlock.Text = "Update Views selected. Press Apply to run."
        self.set_progress_detail("Update Views selected. Press Apply to run.")

    def on_reset_legacy_content(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return

        doc = get_revit_document()
        if doc is None:
            self.FooterStatusTextBlock.Text = "No active Revit document."
            return

        reset_candidates = []
        text_total = 0
        curve_total = 0
        for entry in targets:
            try:
                if normalize_view_type_name(entry.ViewType) != "Drafting View":
                    continue
                if not entry.RevitViewId:
                    continue
                element_id = make_element_id(entry.RevitViewId)
                if element_id is None:
                    continue
                view = doc.GetElement(element_id)
                if view is None or not is_drafting_view(view):
                    continue
                has_legacy, text_count, curve_count = detect_legacy_untagged_table_content(view)
                if has_legacy:
                    reset_candidates.append((entry, view, text_count, curve_count))
                    text_total += text_count
                    curve_total += curve_count
            except Exception:
                pass

        if not reset_candidates:
            self.FooterStatusTextBlock.Text = "No legacy untagged table content found in selected views."
            self.set_progress_detail("No legacy content found.")
            return

        result = MessageBox.Show(
            "Reset legacy untagged table content in %s selected view(s)?\n\n"
            "This removes only untagged TextNotes and Detail Lines owned by those Drafting Views.\n"
            "Symbol families, detail items, annotations, dimensions, tags, images and filled regions are preserved.\n\n"
            "Use this once for views created before the Table Importer tagging system." % len(reset_candidates),
            "Reset Legacy Content",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
        )
        if result != MessageBoxResult.Yes:
            self.FooterStatusTextBlock.Text = "Legacy reset cancelled."
            self.set_progress_detail("Legacy reset cancelled.")
            return

        deleted_views = 0
        deleted_elements = 0
        transaction = None
        try:
            transaction = Transaction(doc, "Table Importer | Reset Legacy Content")
            transaction.Start()
            for entry, view, text_count, curve_count in reset_candidates:
                deleted, deleted_text, deleted_curve = clear_legacy_untagged_table_content(view)
                if deleted:
                    deleted_views += 1
                    deleted_elements += deleted_text + deleted_curve
                    try:
                        entry.Status = ROW_READY_STATUS
                    except Exception:
                        pass
            transaction.Commit()
            transaction = None
        except Exception as ex:
            try:
                if transaction is not None:
                    transaction.RollBack()
            except Exception:
                pass
            self.FooterStatusTextBlock.Text = "Legacy reset failed: %s" % safe_unicode(ex)
            self.set_progress_detail("Legacy reset failed.")
            return

        self.save_current_entries()
        self.TablesDataGrid.Items.Refresh()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "Legacy reset complete. %s view(s), %s element(s) removed. Apply again to regenerate clean tables." % (deleted_views, deleted_elements)
        self.set_progress_detail("Legacy reset complete. Apply again to regenerate clean tables.")

    def update_selected_views(self, close_on_success=True):
        targets = self.require_targets()
        if not targets:
            self.set_apply_progress("Ready", "Ready", 0)
            self.set_progress_detail("No rows selected.")
            self.FooterStatusTextBlock.Text = "No rows selected."
            if DEBUG_OUTPUT:
                try:
                    script.get_output().print_md("Skipped update: no selected rows.")
                except Exception:
                    print("Table Importer: skipped update because no rows were selected.")
            return

        processable_targets = []
        for entry in targets:
            if self.is_processable_entry(entry):
                processable_targets.append(entry)
            else:
                try:
                    if not resolve_entry_path(entry.FilePath) or not os.path.exists(resolve_entry_path(entry.FilePath)):
                        entry.Status = "Missing File"
                    elif not clean_display_text(entry.Region).strip():
                        entry.Status = "Invalid Region"
                    else:
                        entry.Status = "Skipped"
                except Exception:
                    entry.Status = "Skipped"
        if not processable_targets:
            self.set_apply_progress("Ready", "Ready", 0)
            self.set_progress_detail("No selected rows can be processed.")
            self.FooterStatusTextBlock.Text = "No selected rows can be processed."
            self.TablesDataGrid.Items.Refresh()
            self.update_apply_button_state()
            return
        targets = processable_targets

        doc = get_revit_document()
        if doc is None:
            self.set_apply_progress("Failed", "Failed", 0)
            self.set_progress_detail("No active Revit document.")
            self.FooterStatusTextBlock.Text = "No active Revit document."
            return

        created = 0
        updated = 0
        skipped = 0
        failed = 0
        processed = 0
        total = len(targets)
        output_holder = [None]
        skip_reasons = []
        skip_reason_texts = []
        success_messages = []
        self.set_apply_progress("Processing", "Processing", 0)
        self.set_progress_detail("Preparing selected rows...")

        def debug_message(message):
            if not DEBUG_OUTPUT:
                return
            try:
                if output_holder[0] is None:
                    output_holder[0] = script.get_output()
                output_holder[0].print_md(message)
            except Exception:
                print("Table Importer: %s" % safe_unicode(message))

        def row_detail(prefix, entry, reason=None):
            index = processed + 1
            name = self.format_progress_name(entry)
            if reason:
                self.set_progress_detail("%s %s/%s: %s - %s" % (prefix, index, total, reason, name))
            else:
                self.set_progress_detail("%s %s/%s: %s" % (prefix, index, total, name))

        def debug_skip(entry, reason):
            clean_reason = safe_unicode(reason)
            view_name = safe_unicode(getattr(entry, "ViewName", ""))
            message = "Skipped '%s': %s" % (view_name, clean_reason)
            skip_reasons.append(message)
            skip_reason_texts.append(clean_reason)
            try:
                entry.StatusDetail = clean_reason
            except Exception:
                pass
            row_detail("Skipping", entry, clean_reason)
            debug_message(message)

        def is_read_skip_reason(message):
            clean_message = safe_unicode(message).lower()
            if "too large" in clean_message:
                return True
            if "no readable excel data" in clean_message:
                return True
            if "excel file not found" in clean_message:
                return True
            return False

        def status_from_read_error(message):
            clean_message = safe_unicode(message).lower()
            if "excel file not found" in clean_message:
                return "Missing File"
            if "no readable excel data" in clean_message:
                return "Invalid Region"
            return "Skipped"

        legacy_view_count = 0
        legacy_text_count = 0
        legacy_curve_count = 0
        cleanup_legacy = False
        import_context = {
            "import_to_drafting_view": import_to_drafting_view,
            "read_table_data_for_entry": read_table_data_for_entry,
            "get_cell_value": _get_cell_value,
            "make_element_id": make_element_id,
            "get_element_id_value": get_element_id_value,
            "make_unique_name": make_unique_name,
            "get_existing_revit_view_names": get_existing_revit_view_names,
            "sanitize_revit_view_name": sanitize_revit_view_name,
            "clean_display_text": clean_display_text,
            "safe_unicode": safe_unicode,
            "get_revit_uidocument": get_revit_uidocument,
        }
        for candidate in targets:
            try:
                if normalize_view_type_name(candidate.ViewType) != "Drafting View":
                    continue
                if not candidate.RevitViewId:
                    continue
                element_id = make_element_id(candidate.RevitViewId)
                if element_id is None:
                    continue
                candidate_view = doc.GetElement(element_id)
                if candidate_view is None or not is_drafting_view(candidate_view):
                    continue
                has_legacy, text_count, curve_count = detect_legacy_untagged_table_content(candidate_view)
                if has_legacy:
                    legacy_view_count += 1
                    legacy_text_count += text_count
                    legacy_curve_count += curve_count
            except Exception as scan_ex:
                debug_message("Legacy scan skipped one row: %s" % safe_unicode(scan_ex))

        if legacy_view_count:
            warning_text = "Legacy untagged content was preserved in %s view(s). Use Reset Legacy Content only if this view was created with an older version." % legacy_view_count
            self.set_progress_detail(warning_text)
            self.FooterStatusTextBlock.Text = warning_text
            cleanup_legacy = False


        for entry in targets:
            transaction = None
            try:
                ensure_table_entry_uid(entry)
                entry.Status = "Updating..."
                self.TablesDataGrid.Items.Refresh()
                self.refresh_dispatcher()
                entry_import_type = safe_unicode(entry.ImportType)
                if entry_import_type == "Image":
                    entry.Status = "Skipped"
                    skipped += 1
                    debug_skip(entry, "row has invalid source/type '%s'; image import is not implemented" % entry_import_type)
                    continue

                entry_view_type = normalize_view_type_name(entry.ViewType)
                entry.ViewType = get_view_type_display_name(entry_view_type)
                try:
                    if entry_view_type == "Drafting View" and entry.RevitViewId:
                        row_detail("Updating", entry)
                    elif entry_view_type == "Drafting View":
                        row_detail("Creating", entry)
                    else:
                        row_detail("Updating", entry)
                    import_result = import_row_to_revit(entry, entry_view_type, None, doc, import_context, cleanup_legacy)
                    if import_result.status == "Updated":
                        entry.Status = "Created" if import_result.created else "Updated"
                        created += import_result.created
                        updated += import_result.updated
                        if import_result.message:
                            success_messages.append(safe_unicode(import_result.message))
                        if import_result.created:
                            row_detail("Created", entry)
                        else:
                            row_detail("Updated", entry)
                    elif import_result.status == "Skipped":
                        if safe_unicode(import_result.message).startswith("Missing Revit schedule"):
                            entry.Status = "Missing View"
                        else:
                            entry.Status = "Skipped"
                        skipped += max(1, import_result.skipped)
                        debug_skip(entry, import_result.message)
                    else:
                        entry.Status = "Error"
                        failed += max(1, import_result.failed)
                        self.set_progress_detail("Error %s/%s: %s" % (processed + 1, total, self.format_progress_name(entry)))
                        print("Table Importer: %s for '%s'." % (safe_unicode(import_result.message), safe_unicode(entry.ViewName)))
                except Exception as ex:
                    message = safe_unicode(ex)
                    if is_read_skip_reason(message):
                        entry.Status = status_from_read_error(message)
                        skipped += 1
                        debug_skip(entry, "no readable Excel data or file issue: %s" % message)
                    elif message.startswith("Invalid RevitViewId") or message.startswith("Missing Revit view") or message.startswith("Missing Revit schedule") or "not a Drafting View" in message or "not a Schedule View" in message or message.startswith("Schedule View header import is not supported"):
                        if message.startswith("Missing Revit schedule"):
                            entry.Status = "Missing View"
                        else:
                            entry.Status = "Skipped"
                        skipped += 1
                        debug_skip(entry, message)
                    else:
                        entry.Status = "Error"
                        failed += 1
                        self.set_progress_detail("Error %s/%s: %s" % (processed + 1, total, self.format_progress_name(entry)))
                        print("Table Importer: %s for '%s'." % (message, safe_unicode(entry.ViewName)))
                    continue

            except Exception as ex:
                failed += 1
                try:
                    entry.Status = "Error"
                except Exception:
                    pass
                try:
                    if transaction is not None:
                        transaction.RollBack()
                except Exception:
                    pass
                self.set_progress_detail("Error %s/%s: %s" % (processed + 1, total, self.format_progress_name(entry)))
                print("Table Importer: failed '%s': %s" % (safe_unicode(entry.ViewName), safe_unicode(ex)))
            finally:
                processed += 1
                self.update_apply_progress(processed, total)

        self.save_current_entries()
        self.TablesDataGrid.Items.Refresh()
        self.apply_filter()
        success_count = created + updated
        first_skip_reason = ""
        if skip_reason_texts:
            first_skip_reason = skip_reason_texts[0]
        only_one_skip_reason = False
        if first_skip_reason:
            only_one_skip_reason = True
            for skip_reason_text in skip_reason_texts:
                if skip_reason_text != first_skip_reason:
                    only_one_skip_reason = False
                    break
        summary_text = "Created %s view(s). Updated %s view(s). Skipped %s row(s). Failed %s row(s)." % (created, updated, skipped, failed)
        if skipped and success_count == 0 and failed == 0 and only_one_skip_reason:
            summary_text = "%s %s" % (summary_text, first_skip_reason)
        legacy_text = ""
        if legacy_view_count:
            legacy_text = "Legacy untagged content was preserved in %s view(s). Use Reset Legacy Content only if this view was created with an older version." % legacy_view_count
            if failed == 0 and skipped == 0:
                summary_text = "%s %s" % (summary_text, legacy_text)
        self.FooterStatusTextBlock.Text = summary_text
        if failed:
            self.set_apply_progress("Error", "", 100)
        else:
            self.set_apply_progress("Ready", "", 100)
        if legacy_view_count:
            if failed or skipped:
                self.set_progress_detail("Completed: %s created, %s updated, %s skipped, %s failed" % (created, updated, skipped, failed))
            else:
                self.set_progress_detail("Completed: %s created, %s updated, %s skipped, %s failed. %s" % (created, updated, skipped, failed, legacy_text))
        else:
            self.set_progress_detail("Completed: %s created, %s updated, %s skipped, %s failed" % (created, updated, skipped, failed))
        if skip_reasons:
            debug_message("Skip summary: %s skipped row(s). See messages above for reasons." % len(skip_reasons))

        # Keep the window open and report results in footer/status only.

    def duplicate_entry(self, entry):
        copied = TableEntry.from_dict(entry.to_dict())
        copied.Selected = True
        copied.Status = ROW_READY_STATUS
        copied.ViewName = clean_display_text("%s Copy" % safe_unicode(entry.ViewName))
        copied.RevitViewId = None
        copied.TableEntryUid = None
        ensure_table_entry_uid(copied)
        return copied

    def on_duplicate_views(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        self.FooterStatusTextBlock.Text = "Duplicate Views is not implemented yet."
        self.set_progress_detail("Duplicate Views is not implemented yet.")

    def on_reload_from(self, sender, args):
        targets = self.require_targets()
        if targets:
            self.FooterStatusTextBlock.Text = "Reload From placeholder. %s table(s) selected." % len(targets)

    def on_absolute_path(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        changed = 0
        base_folder = get_storage_folder()
        for entry in targets:
            try:
                path = safe_unicode(entry.FilePath)
                if path and not os.path.isabs(path):
                    entry.FilePath = os.path.abspath(os.path.join(base_folder, path))
                    changed += 1
                entry.PathMode = "Absolute"
            except Exception:
                pass
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "%s path(s) set to absolute." % changed

    def on_relative_path(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        changed = 0
        base_folder = get_storage_folder()
        for entry in targets:
            try:
                path = safe_unicode(entry.FilePath)
                if path and os.path.isabs(path):
                    entry.FilePath = os.path.relpath(path, base_folder)
                    changed += 1
                entry.PathMode = "Relative"
            except Exception:
                pass
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "%s path(s) set to relative." % changed

    def on_open_files(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        opened = 0
        missing = 0
        first_missing = ""
        for entry in targets:
            path = resolve_entry_path(entry.FilePath)
            if path and os.path.exists(path) and open_with_windows_shell(path):
                opened += 1
            else:
                missing += 1
                if not first_missing:
                    first_missing = path or safe_unicode(entry.FilePath)
        if first_missing:
            self.FooterStatusTextBlock.Text = "File not found: %s" % safe_unicode(first_missing)
        else:
            self.FooterStatusTextBlock.Text = "Opened %s file(s)." % opened

    def on_open_folders(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        opened_folders = []
        missing = 0
        first_missing = ""
        for entry in targets:
            path = resolve_entry_path(entry.FilePath)
            folder = os.path.dirname(path) if path else ""
            if folder and os.path.exists(folder):
                already = False
                for open_folder in opened_folders:
                    if same_path(open_folder, folder):
                        already = True
                        break
                if not already and open_with_windows_shell(folder):
                    opened_folders.append(folder)
            else:
                missing += 1
                if not first_missing:
                    first_missing = folder or path or safe_unicode(entry.FilePath)
        if first_missing:
            self.FooterStatusTextBlock.Text = "Folder not found: %s" % safe_unicode(first_missing)
        else:
            self.FooterStatusTextBlock.Text = "Opened %s folder(s)." % len(opened_folders)

    def on_delete_views(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        result = MessageBox.Show(
            "Remove selected row(s) from Table Importer?\n\nThis will not delete Revit views.",
            "Remove Rows",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
        )
        if result != MessageBoxResult.Yes:
            self.FooterStatusTextBlock.Text = "Remove rows cancelled."
            return
        removed = 0
        for entry in list(targets):
            try:
                if entry in self.all_entries:
                    self.all_entries.remove(entry)
                    removed += 1
            except Exception:
                pass
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "%s table row(s) removed." % removed

    def on_unlink_view(self, sender, args):
        targets = self.require_targets()
        if targets:
            self.FooterStatusTextBlock.Text = "Unlink View placeholder. %s table(s) selected." % len(targets)

    def on_open_view(self, sender, args):
        targets = self.require_targets()
        if targets:
            self.FooterStatusTextBlock.Text = "Open View placeholder. %s table(s) selected." % len(targets)

    def on_search_changed(self, sender, args):
        self.current_search = safe_unicode(self.SearchTextBox.Text)
        self.apply_filter()
        if self.current_search:
            self.FooterStatusTextBlock.Text = "Search: %s | %s visible." % (self.current_search, self.entries.Count)

    def on_close(self, sender, args):
        self.save_current_entries()
        self.window.Close()

    def show(self):
        self.window.ShowDialog()


if __name__ == "__main__":
    try:
        ui = TableImporterWindow()
        ui.show()
    except Exception as ex:
        output = script.get_output()
        output.print_md("## Table Importer Error")
        output.print_md("```")
        output.print_md(safe_unicode(ex))
        output.print_md("```")

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
from System.Windows.Controls import DataGridRow, ListBoxItem, StackPanel, TextBlock
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

TOOL_VERSION = "MVP 0.3.17"
DEBUG_OUTPUT = False
USE_WRAPPED_TEXT_NOTES = False

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
DEFAULT_IMPORT_TYPES = ["Excel Link", "Excel Import", "Image"]
DEFAULT_VIEW_TYPES = ["Drafting View", "Legend View"]
DEFAULT_SCALES = ["1", "2", "5", "10", "20", "25", "50", "75", "100"]
DEFAULT_DPI = ["72", "96", "150", "200", "300", "600"]
EXCEL_FILTER = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
MAX_TABLE_ROWS = 200
MAX_TABLE_COLUMNS = 30
TABLE_COLUMN_WIDTH = 0.8
TABLE_ROW_HEIGHT = 0.25
TABLE_SCALE = 1.0
EXCEL_WIDTH_SCALE = 0.055
EXCEL_ROW_HEIGHT_SCALE = 1.0
TABLE_TEXT_SIZE_MM = 1.8
TEXT_SIZE_SCALE = 1.0
MIN_TEXT_SIZE_FT = TABLE_TEXT_SIZE_MM / 304.8
MAX_TEXT_SIZE_FT = 0.012
APPROX_TEXT_CHAR_WIDTH_FACTOR = 0.70
MAX_CELL_TEXT_CHARS = 120
MIN_COL_WIDTH_FT = 0.25
MAX_COL_WIDTH_FT = 3.00
MIN_ROW_HEIGHT_FT = 0.18
MAX_ROW_HEIGHT_FT = 0.60
TEXT_PADDING_X = 0.015
TEXT_PADDING_Y = 0.015
APPROX_CHARS_PER_FOOT = 14
ROW_HEIGHT_SCALE = EXCEL_ROW_HEIGHT_SCALE
TABLE_TEXT_OFFSET_X = TEXT_PADDING_X
TABLE_TEXT_OFFSET_Y = TEXT_PADDING_Y
TABLE_MIN_COLUMN_WIDTH = MIN_COL_WIDTH_FT
TABLE_MAX_COLUMN_WIDTH = MAX_COL_WIDTH_FT
TABLE_CHAR_WIDTH = 0.045
TABLE_MIN_ROW_HEIGHT = MIN_ROW_HEIGHT_FT
TABLE_MAX_ROW_HEIGHT = MAX_ROW_HEIGHT_FT
TABLE_LINE_HEIGHT = 0.15
TABLE_TEXT_TYPE_NAME = "MENVIC_TABLE_TEXT_1.8mm"
TABLE_TEXT_SIZE = TABLE_TEXT_SIZE_MM / 304.8
FALLBACK_TABLE_FONT = "Arial"
TEXT_NOTE_WIDTH_WARNING_PRINTED = False
TABLE_IMPORTER_TOOL_NAME = "pyMENVIC_TABLE_IMPORTER"
TABLE_IMPORTER_CREATED_BY = "Table Importer"
TABLE_IMPORTER_TAG_PREFIX = "pyMENVIC_TABLE_IMPORTER|"


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


def get_table_text_note_type(doc, font_name=None):
    text_size = get_table_text_size()
    table_font = clean_display_text(font_name) or FALLBACK_TABLE_FONT
    try:
        for note_type in FilteredElementCollector(doc).OfClass(TextNoteType):
            try:
                if safe_unicode(note_type.Name) == TABLE_TEXT_TYPE_NAME:
                    set_text_note_type_size(note_type, text_size)
                    set_text_note_type_font(note_type, table_font)
                    return note_type
            except Exception:
                pass
    except Exception:
        pass

    default_type = get_default_text_note_type(doc)
    if default_type is None:
        return None

    try:
        new_type = default_type.Duplicate(TABLE_TEXT_TYPE_NAME)
        set_text_note_type_size(new_type, text_size)
        set_text_note_type_font(new_type, table_font)
        return new_type
    except Exception:
        return default_type


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


def clear_table_importer_view(view, entry):
    doc = get_revit_document()
    if doc is None:
        raise Exception("No active Revit document.")

    ensure_table_entry_uid(entry)
    text_note_ids = _collect_tagged_elements_of_class(doc, view, entry, TextNote)
    curve_ids = _collect_tagged_elements_of_class(doc, view, entry, DB.CurveElement)
    ids_to_delete = []
    ids_to_delete.extend(text_note_ids)
    ids_to_delete.extend(curve_ids)

    deleted = 0
    for element_id in ids_to_delete:
        try:
            doc.Delete(element_id)
            deleted += 1
        except Exception as ex:
            print("Table Importer: could not delete view element '%s': %s" % (safe_unicode(element_id), safe_unicode(ex)))
    if DEBUG_OUTPUT:
        try:
            print("Table Importer: regeneration deleted %s tagged TextNote element(s) and %s tagged CurveElement element(s)." % (len(text_note_ids), len(curve_ids)))
        except Exception:
            pass
    return deleted


def reset_table_view_legacy_content(view):
    deleted, text_count, curve_count = clear_legacy_untagged_table_content(view)
    return deleted


def update_existing_drafting_view(entry, view, cleanup_legacy=False):
    table_data, row_count, column_count = read_table_data_for_entry(entry)
    clear_table_importer_view(view, entry)
    if cleanup_legacy:
        clear_legacy_untagged_table_content(view)
    draw_table_in_view(view, table_data, entry)
    entry.Status = "Updated"
    return row_count, column_count


def _make_line(doc, view, x1, y1, x2, y2, entry=None):
    line = Line.CreateBound(XYZ(float(x1), float(y1), 0.0), XYZ(float(x2), float(y2), 0.0))
    detail_curve = doc.Create.NewDetailCurve(view, line)
    if entry is not None:
        apply_table_importer_tag(detail_curve, entry)
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


def fit_text_to_cell(text, cell_width, text_size):
    value = clean_display_text(text)
    value = u" ".join(value.replace(u"\r", u" ").replace(u"\n", u" ").split())
    if not value:
        return u""

    if len(value) > MAX_CELL_TEXT_CHARS:
        value = value[:MAX_CELL_TEXT_CHARS - 3] + u"..."

    try:
        available_width = float(cell_width) - (2.0 * float(TEXT_PADDING_X))
    except Exception:
        available_width = 0.0
    if available_width <= 0:
        return u"..."

    try:
        char_width = float(text_size) * float(APPROX_TEXT_CHAR_WIDTH_FACTOR)
    except Exception:
        char_width = 0.0
    if char_width <= 0:
        char_width = float(MIN_TEXT_SIZE_FT) * float(APPROX_TEXT_CHAR_WIDTH_FACTOR)

    estimated_width = float(len(value)) * char_width
    if estimated_width <= available_width:
        return value

    try:
        max_chars = int(available_width / char_width)
    except Exception:
        max_chars = 3
    if max_chars <= 3:
        return u"..."
    if max_chars > MAX_CELL_TEXT_CHARS:
        max_chars = MAX_CELL_TEXT_CHARS
    if max_chars < 6:
        keep = max(1, max_chars - 3)
        return value[:keep] + u"..."
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


def _draw_full_grid(doc, view, x_positions, y_positions, entry=None):
    drawn_lines = set()
    for x in x_positions:
        _make_line_once(doc, view, drawn_lines, x, y_positions[0], x, y_positions[-1], entry)
    for y in y_positions:
        _make_line_once(doc, view, drawn_lines, x_positions[0], y, x_positions[-1], y, entry)


def _draw_excel_borders(doc, view, borders, rows, cols, x_positions, y_positions, entry=None):
    drawn_lines = set()
    for r in range(0, rows):
        for c in range(0, cols):
            try:
                cell_border = borders[r][c]
            except Exception:
                cell_border = {}
            x1 = x_positions[c]
            x2 = x_positions[c + 1]
            y1 = y_positions[r]
            y2 = y_positions[r + 1]
            if cell_border.get("top"):
                _make_line_once(doc, view, drawn_lines, x1, y1, x2, y1, entry)
            if cell_border.get("right"):
                _make_line_once(doc, view, drawn_lines, x2, y1, x2, y2, entry)
            if cell_border.get("bottom"):
                _make_line_once(doc, view, drawn_lines, x1, y2, x2, y2, entry)
            if cell_border.get("left"):
                _make_line_once(doc, view, drawn_lines, x1, y1, x1, y2, entry)


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


def create_text_note_in_cell(doc, view, point, width, text, text_note_type_id):
    if not USE_WRAPPED_TEXT_NOTES:
        try:
            return TextNote.Create(doc, view.Id, point, text, text_note_type_id)
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
        return TextNote.Create(doc, view.Id, point, constrained_width, text, text_note_type_id)
    except Exception:
        pass

    try:
        options = DB.TextNoteOptions(text_note_type_id)
        return TextNote.Create(doc, view.Id, point, constrained_width, text, options)
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
    column_widths = _get_excel_column_widths(table_data, cols)
    width_source = "Excel"
    if column_widths is None:
        width_source = "fallback"
        column_widths = _calculate_column_widths(table_data, rows, cols, top_left_map, covered_map)
    row_heights = _get_excel_row_heights(table_data, rows)
    height_source = "Excel"
    if row_heights is None:
        height_source = "fallback"
        row_heights = _calculate_row_heights(table_data, rows, cols, column_widths, top_left_map, covered_map)
    column_widths = _scale_dimension_values(column_widths, TABLE_COLUMN_WIDTH)
    row_heights = _scale_dimension_values(row_heights, TABLE_ROW_HEIGHT)
    x_positions, y_positions = _build_table_positions(origin_x, origin_y, column_widths, row_heights)
    _debug_geometry_report(rows, cols, column_widths, row_heights, width_source, height_source)

    borders, border_available = _get_table_borders(table_data, rows, cols)
    if border_available:
        _draw_excel_borders(doc, view, borders, rows, cols, x_positions, y_positions, entry)
    else:
        _draw_full_grid(doc, view, x_positions, y_positions, entry)

    note_type = get_table_text_note_type(doc, get_table_text_font(table_data))
    note_type_id = note_type.Id if note_type is not None else DB.ElementId.InvalidElementId
    text_size = get_text_note_type_size(note_type)
    debug_text_fit_count = [0]
    for r in range(0, rows):
        for c in range(0, cols):
            if _is_covered_non_top_left(r, c, top_left_map, covered_map):
                continue
            value = _get_cell_value(table_data, r, c)
            if not value:
                continue
            x1, y1, x2, y2 = _get_cell_rect(r, c, top_left_map, x_positions, y_positions)
            point = XYZ(
                x1 + TEXT_PADDING_X,
                y1 - TEXT_PADDING_Y,
                0.0,
            )
            cell_width = x2 - x1
            display_value = fit_text_to_cell(value, cell_width, text_size)
            _debug_fit_text(value, display_value, cell_width, text_size, debug_text_fit_count)
            text_note = create_text_note_in_cell(doc, view, point, cell_width, display_value, note_type_id)
            if entry is not None:
                apply_table_importer_tag(text_note, entry)


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
            self.VersionTextBlock.Text = TOOL_VERSION

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
                    status="Not Created",
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
            self.VersionTextBlock.Text = TOOL_VERSION

        self.BatchActionsContextMenu = self.BatchActionsButton.ContextMenu
        self.RowActionsContextMenu = self.TablesDataGrid.ContextMenu
        self.bind_menu_items()
        self.load_logo_image()
        self.load_saved_entries()
        self.set_apply_progress("Status: Ready", "Ready", 100)
        self.set_progress_detail("Ready.")
        self.bind_events()

    def bind_menu_items(self):
        names = [
            "UpdateViews", "DuplicateViews", "ReloadFrom", "AbsolutePath", "RelativePath",
            "OpenFiles", "OpenFolders", "DeleteViews", "UnlinkView", "OpenView"
        ]
        for name in names:
            setattr(self, "Batch%sMenuItem" % name, self.window.FindName("Batch%sMenuItem" % name))
            setattr(self, "Row%sMenuItem" % name, self.window.FindName("Row%sMenuItem" % name))

    def load_logo_image(self):
        load_logo_into_image(self.LogoImage)

    def refresh_dispatcher(self):
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
                self.StatusTextBlock.Text = safe_unicode(status_text)
        except Exception:
            pass
        try:
            if self.ProgressTextBlock:
                self.ProgressTextBlock.Text = "%s   %s%%" % (safe_unicode(progress_label), percent_value)
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
        self.set_apply_progress("Status: Processing...", "Processing", percent)

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
            self.window.FindName(prefix + "UpdateViewsMenuItem").Click += self.on_update_views
            self.window.FindName(prefix + "DuplicateViewsMenuItem").Click += self.on_duplicate_views
            self.window.FindName(prefix + "ReloadFromMenuItem").Click += self.on_reload_from
            self.window.FindName(prefix + "AbsolutePathMenuItem").Click += self.on_absolute_path
            self.window.FindName(prefix + "RelativePathMenuItem").Click += self.on_relative_path
            self.window.FindName(prefix + "OpenFilesMenuItem").Click += self.on_open_files
            self.window.FindName(prefix + "OpenFoldersMenuItem").Click += self.on_open_folders
            self.window.FindName(prefix + "DeleteViewsMenuItem").Click += self.on_delete_views
            self.window.FindName(prefix + "UnlinkViewMenuItem").Click += self.on_unlink_view
            self.window.FindName(prefix + "OpenViewMenuItem").Click += self.on_open_view

    def on_datagrid_single_click(self, sender, args):
        try:
            row = find_visual_parent(args.OriginalSource, DataGridRow)
            if row is not None and not row.IsEditing:
                self.TablesDataGrid.BeginEdit()
        except Exception:
            pass

    def on_cell_edit_ending(self, sender, args):
        try:
            if args.EditAction.ToString() == "Commit":
                self.save_current_entries()
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
        except Exception:
            pass
        self.populate_region_options(entry)

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

    def save_current_entries(self):
        for entry in self.all_entries:
            try:
                ensure_table_entry_uid(entry)
            except Exception:
                pass
        save_entries(self.all_entries)

    def update_footer(self):
        total = len(self.all_entries)
        visible = self.entries.Count
        selected = len(self.get_target_entries(False))
        linked = 0
        drafting = 0
        legend = 0
        for entry in self.all_entries:
            try:
                if entry.RevitViewId:
                    linked += 1
                if entry.ViewType == "Legend View":
                    legend += 1
                else:
                    drafting += 1
            except Exception:
                pass
        self.FooterStatusTextBlock.Text = "Total %s | Visible %s | Selected %s | Linked %s | Drafting %s | Legends %s" % (total, visible, selected, linked, drafting, legend)
        pct = 0
        if total:
            pct = int((float(linked) / float(total)) * 100.0)
        self.CompletedTextBlock.Text = "Completed %s%%" % pct
        self.update_empty_state()

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
        self.FooterStatusTextBlock.Text = "%s table row(s) added." % len(new_entries)

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
                    status="Not Created",
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
                missing += 1
                continue
            self.populate_region_options(entry)
            old_date = entry.LastModified
            new_date = get_last_modified(path)
            entry.LastModified = new_date
            entry.Source = clean_display_text(os.path.basename(path))
            if old_date and new_date and old_date != new_date:
                entry.Status = "Modified"
                updated += 1
            elif entry.RevitViewId:
                entry.Status = "OK"
            else:
                entry.Status = "Not Created"
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "Refresh complete. %s modified, %s missing." % (updated, missing)

    def on_apply(self, sender, args):
        self.update_selected_views()

    def on_batch_actions(self, sender, args):
        try:
            self.BatchActionsContextMenu.PlacementTarget = self.BatchActionsButton
            self.BatchActionsContextMenu.IsOpen = True
        except Exception:
            self.FooterStatusTextBlock.Text = "Batch Actions menu could not be opened."

    def get_target_entries(self, allow_selected_rows=True):
        result = []
        try:
            for entry in self.all_entries:
                if entry.Selected:
                    result.append(entry)
        except Exception:
            pass
        if result or not allow_selected_rows:
            return result
        try:
            for item in self.TablesDataGrid.SelectedItems:
                if item not in result:
                    result.append(item)
        except Exception:
            pass
        return result

    def require_targets(self):
        targets = self.get_target_entries(True)
        if not targets:
            self.FooterStatusTextBlock.Text = "Select one or more tables first."
        return targets

    def on_update_views(self, sender, args):
        self.update_selected_views()

    def update_selected_views(self):
        targets = self.require_targets()
        if not targets:
            self.set_apply_progress("Status: Ready", "Ready", 0)
            self.set_progress_detail("No rows selected.")
            self.FooterStatusTextBlock.Text = "No rows selected."
            if DEBUG_OUTPUT:
                try:
                    script.get_output().print_md("Skipped update: no selected rows.")
                except Exception:
                    print("Table Importer: skipped update because no rows were selected.")
            return

        doc = get_revit_document()
        if doc is None:
            self.set_apply_progress("Status: Ready", "Ready", 0)
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
        self.set_apply_progress("Status: Processing...", "Processing", 0)
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
            view_name = safe_unicode(getattr(entry, "ViewName", ""))
            message = "Skipped '%s': %s" % (view_name, reason)
            skip_reasons.append(message)
            row_detail("Skipping", entry, reason)
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

        legacy_view_count = 0
        legacy_text_count = 0
        legacy_curve_count = 0
        cleanup_legacy = False
        for candidate in targets:
            try:
                if safe_unicode(candidate.ViewType) != "Drafting View":
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
            warning_text = "Warning: legacy untagged table content detected in %s view(s)." % legacy_view_count
            self.set_progress_detail(warning_text)
            self.FooterStatusTextBlock.Text = "%s Use Reset Legacy Content if needed." % warning_text
            result = MessageBox.Show(
                "Legacy untagged table content was detected in selected Table Importer views.\n\nThis usually comes from older MVP versions. Do you want to remove old untagged TextNotes and Detail Lines in these views one time before regenerating?\n\nManually placed symbols/families/detail items will be preserved.",
                "Reset Legacy Table Content",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
            )
            cleanup_legacy = result == MessageBoxResult.Yes
            if cleanup_legacy:
                self.set_progress_detail("Legacy cleanup enabled for %s view(s)." % legacy_view_count)
            else:
                self.set_progress_detail("%s Manual cleanup may be required." % warning_text)

        for entry in targets:
            transaction = None
            try:
                ensure_table_entry_uid(entry)
                entry_view_type = safe_unicode(entry.ViewType)
                if entry_view_type != "Drafting View":
                    entry.Status = "Skipped"
                    skipped += 1
                    debug_skip(entry, "unsupported view type '%s'" % entry_view_type)
                    continue

                entry_import_type = safe_unicode(entry.ImportType)
                if entry_import_type == "Image":
                    entry.Status = "Skipped"
                    skipped += 1
                    debug_skip(entry, "row has invalid source/type '%s'; image import is not implemented" % entry_import_type)
                    continue

                if entry.RevitViewId:
                    row_detail("Updating", entry)
                    debug_message("Resolving RevitViewId: %s" % entry.RevitViewId)
                    element_id = make_element_id(entry.RevitViewId)
                    debug_message("ElementId resolved: %s" % str(element_id))
                    if element_id is None:
                        entry.Status = "Skipped"
                        skipped += 1
                        debug_skip(entry, "invalid RevitViewId '%s'" % safe_unicode(entry.RevitViewId))
                        continue

                    view = doc.GetElement(element_id)
                    if view is None:
                        entry.Status = "Missing View"
                        skipped += 1
                        debug_skip(entry, "missing Revit view for RevitViewId '%s'" % safe_unicode(entry.RevitViewId))
                        continue
                    if not is_drafting_view(view):
                        entry.Status = "Skipped"
                        skipped += 1
                        try:
                            revit_view_type = safe_unicode(view.ViewType)
                        except Exception:
                            revit_view_type = "<unknown>"
                        debug_skip(entry, "existing view was not a Drafting View; Revit view type is '%s'" % revit_view_type)
                        continue

                    try:
                        debug_message("Updating existing Drafting View '%s' from row '%s'." % (safe_unicode(view.Name), safe_unicode(entry.ViewName)))
                        transaction = Transaction(doc, "Update Table Importer Drafting View")
                        transaction.Start()
                        update_existing_drafting_view(entry, view, cleanup_legacy)
                        transaction.Commit()
                        transaction = None
                        updated += 1
                        row_detail("Updated", entry)
                        debug_message("Updated existing Drafting View '%s'." % safe_unicode(view.Name))
                    except Exception as update_ex:
                        try:
                            if transaction is not None:
                                transaction.RollBack()
                        except Exception:
                            pass
                        message = safe_unicode(update_ex)
                        if is_read_skip_reason(message):
                            entry.Status = "Skipped"
                            skipped += 1
                            debug_skip(entry, "no readable Excel data or file issue during update: %s" % message)
                            continue
                        raise
                    continue

                try:
                    row_detail("Creating", entry)
                    table_data, row_count, column_count = read_table_data_for_entry(entry)
                except Exception as ex:
                    message = safe_unicode(ex)
                    if is_read_skip_reason(message):
                        entry.Status = "Skipped"
                        skipped += 1
                        debug_skip(entry, "no readable Excel data or file issue during create: %s" % message)
                    else:
                        entry.Status = "Error"
                        failed += 1
                        self.set_progress_detail("Error %s/%s: %s" % (processed + 1, total, self.format_progress_name(entry)))
                        print("Table Importer: %s for '%s'." % (message, safe_unicode(entry.ViewName)))
                    continue

                transaction = Transaction(doc, "Create Table Importer Drafting View")
                transaction.Start()
                view, was_created = get_or_create_drafting_view(entry)
                draw_table_in_view(view, table_data, entry)
                transaction.Commit()
                transaction = None

                if was_created:
                    entry.Status = "Created"
                    created += 1
                    row_detail("Created", entry)
                else:
                    entry.Status = "OK"
                    row_detail("Processed", entry)

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
        summary_text = "Created %s view(s). Updated %s view(s). Skipped %s row(s). Failed %s row(s)." % (created, updated, skipped, failed)
        if legacy_view_count and not cleanup_legacy:
            summary_text = "%s Warning: legacy untagged content remains in %s view(s)." % (summary_text, legacy_view_count)
        elif legacy_view_count and cleanup_legacy:
            summary_text = "%s Legacy cleanup applied to %s view(s)." % (summary_text, legacy_view_count)
        self.FooterStatusTextBlock.Text = summary_text
        self.set_apply_progress("Status: Ready", "Ready", 100)
        if legacy_view_count and not cleanup_legacy:
            self.set_progress_detail("Warning: legacy untagged table content remains in %s view(s)." % legacy_view_count)
        else:
            self.set_progress_detail("Completed: %s created, %s updated, %s skipped, %s failed" % (created, updated, skipped, failed))
        if skip_reasons:
            debug_message("Skip summary: %s skipped row(s). See messages above for reasons." % len(skip_reasons))

    def duplicate_entry(self, entry):
        copied = TableEntry.from_dict(entry.to_dict())
        copied.Selected = True
        copied.Status = "Not Created"
        copied.ViewName = clean_display_text("%s Copy" % safe_unicode(entry.ViewName))
        copied.RevitViewId = None
        copied.TableEntryUid = None
        ensure_table_entry_uid(copied)
        return copied

    def on_duplicate_views(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        new_entries = []
        for entry in targets:
            try:
                new_entries.append(self.duplicate_entry(entry))
            except Exception:
                pass
        for entry in new_entries:
            self.all_entries.append(entry)
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "%s duplicate table row(s) added." % len(new_entries)

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
        for entry in targets:
            path = resolve_entry_path(entry.FilePath)
            if path and os.path.exists(path) and open_with_windows_shell(path):
                opened += 1
            else:
                missing += 1
        self.FooterStatusTextBlock.Text = "Opened %s file(s). %s missing." % (opened, missing)

    def on_open_folders(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        opened_folders = []
        missing = 0
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
        self.FooterStatusTextBlock.Text = "Opened %s folder(s). %s missing." % (len(opened_folders), missing)

    def on_delete_views(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        result = MessageBox.Show(
            "Remove %s selected table row(s)?\n\nThis does not delete Revit views yet." % len(targets),
            "Delete Views",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
        )
        if result != MessageBoxResult.Yes:
            self.FooterStatusTextBlock.Text = "Delete cancelled."
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











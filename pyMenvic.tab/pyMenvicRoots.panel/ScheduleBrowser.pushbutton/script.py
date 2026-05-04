# -*- coding: utf-8 -*-

__title__ = "SheetLink - pyMENVIC"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
==========================================================
pyMENVIC | SHEETLINK
Revit + pyRevit

Descripción
-----------
Herramienta para explorar schedules del modelo, exportarlos a Excel
e importar cambios desde Excel a Revit, usando metadatos y columnas
técnicas de identificación cuando sea posible.

Capacidades
-----------
- Lista schedules disponibles en el proyecto
- Muestra campos visibles con origen y estado editable
- Exporta schedules a Excel con las hojas Data y Schema
- Escribe UniqueId y ElementId en columnas técnicas cuando hay una referencia fiable
- Permite previsualizar cambios antes de importar desde Excel
- Permite importar cambios desde Excel a Revit
- Detecta conflictos de unicidad en campos conocidos durante el preview y la importación

Funciones principales
---------------------
collect_schedules
    Recoge los schedules disponibles del modelo

get_schedule_parameters
    Analiza los campos visibles del schedule seleccionado

export_schedule_to_xlsx
    Exporta el schedule a Excel con formato y metadatos

run_import_preview
    Analiza el archivo Excel y muestra un resumen de cambios y conflictos

run_import_apply
    Importa cambios desde Excel a Revit cuando las referencias y validaciones lo permiten

Reglas importantes
------------------
- Compatible con IronPython 2.7
- Cambios mínimos sobre la lógica original
- No usa librerías externas fuera de Interop Excel
- El import usa UniqueId y ElementId como referencias técnicas principales
- El orden de las filas en Excel no se usa como referencia de importación

Autor
-----
Ricardo J. Mendieta
pyMENVIC – Ayudas para MENVIC ARQ
==========================================================
"""

import os
import sys
import time
import csv
import json
import subprocess
import io

from pyrevit import forms
from pyrevit import script
from Autodesk.Revit import DB
from Autodesk.Revit.UI import TaskDialog

import clr
clr.AddReference("System")
clr.AddReference("System.Data")
clr.AddReference("System.Windows.Forms")
from System import Array, Object, Type, Activator
from System.Data import DataTable
from System.Reflection import BindingFlags, Missing
from System.Runtime.InteropServices import Marshal
from Microsoft.Win32 import SaveFileDialog, OpenFileDialog
from System.Windows import Visibility, FontWeights, Style, Setter, DataTrigger
from System.Windows.Controls import DataGridCell
from System.Windows.Forms import Application
from System.Windows.Data import Binding, BindingMode, UpdateSourceTrigger
from System.Windows.Media import BrushConverter
logger = script.get_logger()


def import_excel_namespace():
    from Microsoft.Office.Interop import Excel as ExcelNamespace
    return ExcelNamespace


def load_excel_interop():
    interop_reference_names = [
        "Microsoft.Office.Interop.Excel",
    ]
    interop_search_roots = [
        r"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel",
        r"C:\Program Files\Microsoft Office\root\Office16",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16",
        r"C:\Program Files\Microsoft Office\Office16",
        r"C:\Program Files (x86)\Microsoft Office\Office16",
    ]

    for reference_name in interop_reference_names:
        try:
            clr.AddReference(reference_name)
            return import_excel_namespace()
        except Exception:
            pass

    for reference_name in interop_reference_names:
        try:
            clr.AddReferenceByName(reference_name)
            return import_excel_namespace()
        except Exception:
            pass

    for root_path in interop_search_roots:
        try:
            if not os.path.isdir(root_path):
                continue
        except Exception:
            continue

        candidate_paths = []
        direct_path = os.path.join(root_path, "Microsoft.Office.Interop.Excel.dll")
        if os.path.isfile(direct_path):
            candidate_paths.append(direct_path)

        try:
            for dirpath, dirnames, filenames in os.walk(root_path):
                for filename in filenames:
                    if filename.lower() == "microsoft.office.interop.excel.dll":
                        candidate_paths.append(os.path.join(dirpath, filename))
        except Exception:
            pass

        for candidate_path in candidate_paths:
            try:
                clr.AddReferenceToFileAndPath(candidate_path)
                return import_excel_namespace()
            except Exception:
                pass

    return None


Excel = load_excel_interop()


def get_excel_interop_members_text():
    if Excel is None:
        return ""

    try:
        members = [safe_text(x) for x in dir(Excel)]
        members = [x for x in members if x and not x.startswith("__")]
        members.sort()
        return ", ".join(members[:80])
    except Exception as ex:
        return "Could not inspect Excel namespace members. {}".format(safe_text(ex))


def unwrap_com_value(value):
    if isinstance(value, ComObjectWrapper):
        return value._com_object
    return value


def wrap_com_value(value):
    if value is None:
        return value

    if isinstance(value, ComObjectWrapper):
        return value

    try:
        type_name = value.GetType().FullName
        if type_name == "System.__ComObject":
            return ComObjectWrapper(value)
    except Exception:
        pass

    return value


def build_com_args(args):
    if args is None:
        args = []
    return Array[Object]([unwrap_com_value(arg) for arg in args])


def get_com_missing_value():
    try:
        return Missing.Value
    except Exception:
        pass

    try:
        return Type.Missing
    except Exception:
        return None


def get_excel_save_as_file_format(path_value):
    try:
        path_text = str(path_value).lower()
    except Exception:
        path_text = ""

    if path_text.endswith(".xlsx"):
        return 51
    if path_text.endswith(".xls"):
        return 56
    if path_text.endswith(".xlsm"):
        return 52

    return get_com_missing_value()


def build_excel_save_as_args(args):
    if not args:
        return build_com_args(args)

    filename = unwrap_com_value(args[0])
    missing = get_com_missing_value()
    file_format = get_excel_save_as_file_format(filename)

    save_as_args = [
        filename,
        file_format,
        missing,
        missing,
        missing,
        missing,
        missing,
        missing,
        missing,
        missing,
        missing,
        True
    ]
    return Array[Object](save_as_args)


EXCEL_COM_METHOD_NAMES = set([
    "Add",
    "Open",
    "Save",
    "SaveAs",
    "Close",
    "Quit",
    "Delete",
    "Clear",
    "ClearContents",
    "AutoFit",
    "Activate",
    "Select",
    "Copy",
    "Paste",
])


class ComMethodWrapper(object):
    def __init__(self, com_object, method_name):
        self._com_object = com_object
        self._method_name = method_name

    def __call__(self, *args):
        last_error = None
        method_name = self._method_name

        if method_name == "SaveAs":
            try:
                result = self._com_object.GetType().InvokeMember(
                    method_name,
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    None,
                    self._com_object,
                    build_excel_save_as_args(args)
                )
                return wrap_com_value(result)
            except Exception as ex:
                last_error = ex

        for flags in (BindingFlags.InvokeMethod, BindingFlags.GetProperty):
            try:
                result = self._com_object.GetType().InvokeMember(
                    method_name,
                    flags | BindingFlags.Public | BindingFlags.Instance,
                    None,
                    self._com_object,
                    build_com_args(args)
                )
                return wrap_com_value(result)
            except Exception as ex:
                last_error = ex

        details = ""
        try:
            if last_error is not None:
                details = " {}".format(str(last_error))
        except Exception:
            details = ""

        raise Exception("COM member '{}' could not be invoked.{}".format(method_name, details))

    def __getitem__(self, key):
        if isinstance(key, tuple):
            args = list(key)
        else:
            args = [key]

        flags = BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance
        result = self._com_object.GetType().InvokeMember(
            self._method_name,
            flags,
            None,
            self._com_object,
            build_com_args(args)
        )
        return wrap_com_value(result)


class ComObjectWrapper(object):
    def __init__(self, com_object):
        object.__setattr__(self, "_com_object", com_object)

    def _invoke(self, member_name, flags, args=None):
        result = self._com_object.GetType().InvokeMember(
            member_name,
            flags | BindingFlags.Public | BindingFlags.Instance,
            None,
            self._com_object,
            build_com_args(args)
        )
        return wrap_com_value(result)

    def __getattr__(self, name):
        if name in EXCEL_COM_METHOD_NAMES:
            return ComMethodWrapper(self._com_object, name)

        try:
            return self._invoke(name, BindingFlags.GetProperty)
        except Exception:
            return ComMethodWrapper(self._com_object, name)

    def __setattr__(self, name, value):
        if name == "_com_object":
            object.__setattr__(self, name, value)
            return

        self._invoke(name, BindingFlags.SetProperty, [value])

    def __call__(self, *args):
        for member_name, flags in (
            ("_Default", BindingFlags.InvokeMethod),
            ("_Default", BindingFlags.GetProperty),
            ("Item", BindingFlags.InvokeMethod),
            ("Item", BindingFlags.GetProperty),
        ):
            try:
                return self._invoke(member_name, flags, args)
            except Exception:
                pass

        raise Exception("COM object is not callable through _Default or Item.")

    def __getitem__(self, key):
        if isinstance(key, tuple):
            args = list(key)
        else:
            args = [key]

        for member_name in ("Item", "_Default"):
            try:
                return self._invoke(member_name, BindingFlags.GetProperty, args)
            except Exception:
                pass

        raise Exception("COM object does not expose an indexer for the requested arguments.")

    def __setitem__(self, key, value):
        if isinstance(key, tuple):
            args = list(key)
        else:
            args = [key]
        args.append(value)

        for member_name in ("Item", "_Default"):
            try:
                self._invoke(member_name, BindingFlags.SetProperty, args)
                return
            except Exception:
                pass

        raise Exception("COM object does not expose a writable indexer for the requested arguments.")


def create_excel_application():
    if Excel is None:
        raise Exception(
            "Microsoft.Office.Interop.Excel could not be loaded."
        )

    typed_error = ""

    try:
        if hasattr(Excel, "ApplicationClass"):
            app = Excel.ApplicationClass()
            app.Workbooks
            app.DisplayAlerts = False
            app.Visible = False
            return app
    except Exception as ex:
        typed_error = safe_text(ex)

    try:
        excel_type = Type.GetTypeFromProgID("Excel.Application")
        if excel_type is None:
            raise Exception("Excel.Application ProgID was not found. Microsoft Excel may not be installed.")

        com_app = Activator.CreateInstance(excel_type)
    except Exception as ex:
        raise Exception(
            "Excel.Application could not be created. {}".format(safe_text(ex))
        )

    for attr_name in ["_Application", "Application"]:
        if not hasattr(Excel, attr_name):
            continue

        try:
            target_type = getattr(Excel, attr_name)
            app = clr.Convert(com_app, target_type)
            app.Workbooks
            app.DisplayAlerts = False
            app.Visible = False
            return app
        except Exception:
            pass

    try:
        app = ComObjectWrapper(com_app)
        app.Workbooks
        app.DisplayAlerts = False
        app.Visible = False
        return app
    except Exception as ex:
        details = safe_text(ex)
        if typed_error:
            details = "{} | Typed attempt: {}".format(details, typed_error)
        raise Exception(
            "Microsoft.Office.Interop.Excel was loaded, but Excel Application could not be created. {}".format(details)
        )

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

XAML_FILENAME = "window.xaml"
THIS_DIR = os.path.dirname(__file__)
XAML_PATH = os.path.join(THIS_DIR, XAML_FILENAME)
PROJECT_STANDARDS_XAML_FILENAME = "project_standards_dialog.xaml"
PROJECT_STANDARDS_XAML_PATH = os.path.join(THIS_DIR, PROJECT_STANDARDS_XAML_FILENAME)

TEMP_FOLDER = os.path.expandvars("%temp%")
MAX_SAMPLE_ELEMENTS = 20
EXPORT_DELIMITER = ","

MODEL_CATEGORY_VISIBLE_NAMES = set([
    "hvac zones",
    "zonas hvac",
    "lines",
    "lineas",
    "líneas",
    "materials",
    "materiales",
    "pipe segments",
    "segmentos de tuberia",
    "segmentos de tubería",
    "sheets",
    "hojas",
    "views",
    "vistas",
])

PROJECT_STANDARDS_SECTIONS = [
    ("project_information", "Project Information"),
    ("project_parameters", "Project Parameters"),
    ("object_styles", "Object Styles"),
    ("line_styles", "Line Styles"),
    ("families", "Families"),
]

class ScheduleItem(object):
    def __init__(self, schedule):
        self.Schedule = schedule
        self.Name = get_schedule_name(schedule)


class ParameterItem(object):
    def __init__(self, name, origin, editable, metadata=None):
        self.Name = name
        self.Origin = origin
        self.Editable = editable
        self.Status = build_status(origin, editable)
        self.SymbolKind = get_parameter_symbol_kind(origin, editable)
        self.SymbolGlyph = get_parameter_symbol(origin, editable)
        self.DisplayName = u"{} {}".format(self.SymbolGlyph, self.Name)
        self.Metadata = metadata or {}


class CategoryItem(object):
    def __init__(self, category, element_count):
        self.Category = category
        self.Name = safe_text(getattr(category, "Name", ""))
        self.ElementCount = element_count


class ElementItem(object):
    def __init__(self, element):
        self.Element = element
        self.Name = get_element_display_name(element)
        self.ElementId = safe_text(get_element_id_value(element))
        try:
            self.UniqueId = safe_text(element.UniqueId)
        except Exception:
            self.UniqueId = ""


def safe_text(value):
    try:
        if value is None:
            return ""
        return str(value)
    except Exception:
        return ""


def normalize_text(value):
    text = safe_text(value)
    text = text.replace(u"\ufeff", "")
    text = text.replace(u"\xef\xbb\xbf", "")
    return text.strip()


def normalize_header(value):
    return normalize_text(value).strip().lower().replace(" ", "")


def build_status(origin, editable):
    origin_text = safe_text(origin).strip()
    editable_text = safe_text(editable).strip()

    if not origin_text:
        origin_text = "Special"

    if origin_text == "Instance":
        origin_text = "Inst"
    elif origin_text == "Type":
        origin_text = "Type"
    elif origin_text == "Mixed":
        origin_text = "Mixed"

    if editable_text == "Yes":
        return "{} / Editable".format(origin_text)
    if editable_text == "No":
        return "{} / Locked".format(origin_text)
    return "{} / Unknown".format(origin_text)


def is_editable_metadata_value(value):
    return normalize_text(value) == "Yes"


def get_parameter_symbol(origin, editable):
    editable_text = safe_text(editable).strip()
    origin_text = safe_text(origin).strip()

    if editable_text == "No":
        return u"\u25CF"

    if origin_text == "Type":
        return u"\u25B2"

    return u"\u25A0"


def get_parameter_symbol_kind(origin, editable):
    editable_text = safe_text(editable).strip()
    origin_text = safe_text(origin).strip()

    if editable_text == "No":
        return "ReadOnly"

    if origin_text == "Type":
        return "Type"

    return "Instance"


def sanitize_filename(text):
    if not text:
        return "Schedule"

    clean_chars = []
    for ch in safe_text(text):
        if ch.isalnum() or ch in (" ", "_", "-"):
            clean_chars.append(ch)

    result = "".join(clean_chars).strip()
    if not result:
        result = "Schedule"

    return result.replace(" ", "_")

def ask_output_xlsx_path_for_name(default_name):
    file_name = sanitize_filename(default_name)

    dialog = SaveFileDialog()
    dialog.Title = "Save Excel Export"
    dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
    dialog.FileName = "{}.xlsx".format(file_name)
    dialog.DefaultExt = ".xlsx"
    dialog.AddExtension = True
    dialog.OverwritePrompt = True

    result = dialog.ShowDialog()

    if result:
        return dialog.FileName

    return None

def ask_input_xlsx_path():
    dialog = OpenFileDialog()
    dialog.Title = "Select Excel File to Import"
    dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
    dialog.Multiselect = False

    result = dialog.ShowDialog()
    if result:
        return dialog.FileName

    return None

   
def get_schedule_name(schedule):
    try:
        if getattr(schedule, "Title", None):
            return schedule.Title
    except Exception:
        pass

    try:
        if getattr(schedule, "Name", None):
            return schedule.Name
    except Exception:
        pass

    return "Unnamed Schedule"

def get_parameter_builtin_id_value(param):
    if param is None:
        return None

    try:
        pid = getattr(param, "Id", None)
        if pid is None:
            return None
    except Exception:
        return None

    value = get_element_id_value(pid)
    if value == "":
        return None
    return value


def get_element_id_value(element_or_id):
    if element_or_id is None:
        return ""

    try:
        eid = element_or_id.Id
    except Exception:
        eid = element_or_id

    if eid is None:
        return ""

    try:
        return int(eid.Value)
    except Exception:
        pass

    try:
        return int(eid.IntegerValue)
    except Exception:
        pass

    return ""



def is_yesno_parameter_safe(param):
    if param is None:
        return False

    try:
        definition = getattr(param, "Definition", None)
        if definition is None:
            return False

        get_data_type = getattr(definition, "GetDataType", None)
        if get_data_type is None:
            return False

        data_type = get_data_type()
        if data_type is None:
            return False

        type_id = safe_text(getattr(data_type, "TypeId", "")).lower()
        if "yesno" in type_id or "spec.bool" in type_id or "spec.boolean" in type_id:
            return True
    except Exception:
        pass

    return False



def try_parse_yesno_value(value):
    text = normalize_text(value).lower()
    if text == "":
        return None
    if text in ("1", "true", "yes", "y", "si", "sí", "x"):
        return 1
    if text in ("0", "false", "no", "n"):
        return 0
    return None


def try_parse_float_value(value):
    text = normalize_text(value)
    if text == "":
        return None

    try:
        return float(text)
    except Exception:
        pass

    try:
        if "," in text and "." not in text:
            return float(text.replace(",", "."))
    except Exception:
        pass

    return None



def get_parameter_unit_type_id(param):
    if param is None:
        return None

    try:
        definition = getattr(param, "Definition", None)
        if definition is None:
            return None

        get_data_type = getattr(definition, "GetDataType", None)
        if get_data_type is None:
            return None

        data_type = get_data_type()
        if data_type is None:
            return None

        is_measurable_spec = getattr(DB.UnitUtils, "IsMeasurableSpec", None)
        if is_measurable_spec is None or not is_measurable_spec(data_type):
            return None

        format_options = doc.GetUnits().GetFormatOptions(data_type)
        if format_options is None:
            return None

        return format_options.GetUnitTypeId()
    except Exception:
        return None



def try_resolve_workset_id(value):
    text = normalize_text(value)
    if not text:
        return None

    try:
        collector = DB.FilteredWorksetCollector(doc)
        for workset in collector:
            try:
                if normalize_text(workset.Name) == text:
                    return workset.Id.IntegerValue
            except Exception:
                continue
    except Exception:
        pass

    return None

def ask_output_xlsx_path(schedule):
    return ask_output_xlsx_path_for_name(get_schedule_name(schedule))



def try_resolve_elementid_by_name(param, value):
    if param is None:
        return None

    text = normalize_text(value)
    if not text:
        return DB.ElementId.InvalidElementId

    try:
        return DB.ElementId(int(float(text)))
    except Exception:
        pass

    try:
        current_ref = doc.GetElement(param.AsElementId())
        if current_ref is None:
            return None

        category = getattr(current_ref, "Category", None)
        if category is None:
            return None

        collector = DB.FilteredElementCollector(doc).OfCategoryId(category.Id)
        for element in collector:
            try:
                if normalize_text(getattr(element, "Name", "")) == text:
                    return element.Id
            except Exception:
                continue
    except Exception:
        pass

    return None



def set_parameter_value(param, value):
    if param is None:
        return False

    try:
        if param.IsReadOnly:
            return False
    except Exception:
        return False

    value = normalize_text(value)

    try:
        storage_type = param.StorageType
    except Exception:
        return False

    try:
        if storage_type == DB.StorageType.String:
            param.Set(value)
            return True

        if storage_type == DB.StorageType.Integer:
            if value == "":
                return False

            param_id_value = get_parameter_builtin_id_value(param)
            if param_id_value == int(DB.BuiltInParameter.ELEM_PARTITION_PARAM):
                workset_id_value = try_resolve_workset_id(value)
                if workset_id_value is not None:
                    param.Set(workset_id_value)
                    return True

            yesno_value = try_parse_yesno_value(value) if is_yesno_parameter_safe(param) else None
            if yesno_value is not None:
                param.Set(yesno_value)
                return True

            try:
                if param.SetValueString(value):
                    return True
            except Exception:
                pass

            numeric_value = try_parse_float_value(value)
            if numeric_value is None:
                return False

            param.Set(int(numeric_value))
            return True

        if storage_type == DB.StorageType.Double:
            if value == "":
                return False

            try:
                if param.SetValueString(value):
                    return True
            except Exception:
                pass

            unit_type_id = get_parameter_unit_type_id(param)
            if unit_type_id is not None:
                # For measurable specs, falling back to raw internal doubles is unsafe.
                return False

            numeric_value = try_parse_float_value(value)
            if numeric_value is None:
                return False

            param.Set(numeric_value)
            return True

        if storage_type == DB.StorageType.ElementId:
            target_id = try_resolve_elementid_by_name(param, value)
            if target_id is None:
                return False

            param.Set(target_id)
            return True

    except Exception:
        return False

    return False

def run_import_apply(xlsx_path):
    excel_app = None
    workbooks = None
    workbook = None
    t = None

    updated = 0
    skipped = 0
    skipped_unresolved = 0
    skipped_unchanged = 0
    failed = 0
    duplicate_count = 0
    unresolved_count = 0
    missing_param_count = 0

    duplicate_lines = []
    unresolved_lines = []
    failed_lines = []

    try:
        excel_app = create_excel_application()
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

        workbooks = excel_app.Workbooks
        workbook = workbooks.Open(xlsx_path, False, True)

        data_info = read_data_sheet_for_import_preview(workbook)
        editable_columns = get_editable_columns_for_import_preview(data_info)
        rows = data_info.get("Rows", [])

        t = DB.Transaction(doc, "Import Excel to Revit")
        t.Start()

        for row in rows:
            excel_row = row.get("ExcelRow", 0)
            element, resolved_by = resolve_element_from_import_row(row)

            if element is None:
                skipped += 1
                skipped_unresolved += 1
                unresolved_count += 1

                if len(unresolved_lines) < 10:
                    unresolved_lines.append(
                        "Row {} | UniqueId='{}' | ElementId='{}'".format(
                            excel_row,
                            safe_text(row.get("Values", {}).get(1, "")),
                            safe_text(row.get("Values", {}).get(2, ""))
                        )
                    )
                continue

            row_values = row.get("Values", {})
            row_changed = False

            for col_info in editable_columns:
                excel_col = col_info.get("ExcelCol", 0)
                metadata = col_info.get("Metadata")
                field_name = col_info.get("Name", "")
                excel_value = normalize_text(row_values.get(excel_col, ""))

                param, origin = get_parameter_from_metadata(element, metadata)
                if param is None:
                    missing_param_count += 1
                    continue

                current_value = get_parameter_preview_value(param)

                if normalize_text(current_value) == normalize_text(excel_value):
                    continue

                if is_unique_controlled_field(field_name):
                    is_dup, dup_message = value_exists_in_other_elements(field_name, excel_value, element)
                    if is_dup:
                        failed += 1
                        duplicate_count += 1

                        if len(duplicate_lines) < 15:
                            duplicate_lines.append(
                                "Row {} | Id {} | {} | '{}' | {}".format(
                                    excel_row,
                                    safe_text(get_element_id_value(element)),
                                    field_name,
                                    excel_value,
                                    dup_message
                                )
                            )
                        continue

                ok = set_parameter_value(param, excel_value)

                if ok:
                    updated += 1
                    row_changed = True
                else:
                    failed += 1
                    if len(failed_lines) < 15:
                        failed_lines.append(
                            "Row {} | Id {} | {} | '{}'".format(
                                excel_row,
                                safe_text(get_element_id_value(element)),
                                field_name,
                                excel_value
                            )
                        )

            if not row_changed:
                skipped += 1
                skipped_unchanged += 1

        t.Commit()

        message = []
        message.append("Import completed.")
        message.append("")
        message.append("Updated values: {}".format(updated))
        message.append("Skipped: {}".format(skipped))
        message.append("Skipped unresolved: {}".format(skipped_unresolved))
        message.append("Skipped unchanged: {}".format(skipped_unchanged))
        message.append("Failed: {}".format(failed))
        message.append("Duplicate conflicts: {}".format(duplicate_count))
        message.append("Unresolved rows: {}".format(unresolved_count))
        message.append("Missing parameters: {}".format(missing_param_count))

        if duplicate_lines:
            message.append("")
            message.append("Sample duplicate conflicts:")
            for line in duplicate_lines:
                message.append(line)

        if unresolved_lines:
            message.append("")
            message.append("Sample unresolved rows:")
            for line in unresolved_lines:
                message.append(line)

        if failed_lines:
            message.append("")
            message.append("Sample failed writes:")
            for line in failed_lines:
                message.append(line)

        TaskDialog.Show(__title__, "\n".join(message))

    except Exception as ex:
        try:
            if t is not None:
                t.RollBack()
        except Exception:
            pass

        TaskDialog.Show(__title__, "Import failed.\n\n{}".format(safe_text(ex)))

    finally:
        try:
            if workbook is not None:
                workbook.Close(False)
        except Exception:
            pass

        try:
            if excel_app is not None:
                excel_app.Quit()
        except Exception:
            pass

        release_com_object(workbook)
        release_com_object(workbooks)
        release_com_object(excel_app)

def collect_schedules():
    items = []

    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)
    for schedule in collector:
        try:
            if getattr(schedule, "IsTemplate", False):
                continue
            if getattr(schedule, "IsTitleblockRevisionSchedule", False):
                continue
            items.append(ScheduleItem(schedule))
        except Exception as ex:
            logger.warning("Skipping schedule. %s", safe_text(ex))

    return sorted(items, key=lambda x: x.Name.lower())


def get_category_element_count(category):
    if category is None:
        return 0

    try:
        collector = DB.FilteredElementCollector(doc).OfCategoryId(category.Id).WhereElementIsNotElementType()
        return collector.GetElementCount()
    except Exception:
        return 0


def get_elements_for_scope(scope):
    scope_text = normalize_text(scope)
    elements = []

    if scope_text == "Current Selection":
        try:
            selection_ids = uidoc.Selection.GetElementIds()
        except Exception:
            return elements

        for element_id in selection_ids:
            try:
                element = doc.GetElement(element_id)
                if element is not None:
                    elements.append(element)
            except Exception:
                pass
        return elements

    if scope_text == "Active View":
        try:
            active_view = doc.ActiveView
            if active_view is None:
                return elements
            collector = DB.FilteredElementCollector(doc, active_view.Id).WhereElementIsNotElementType()
            for element in collector:
                if element is not None:
                    elements.append(element)
        except Exception:
            pass
        return elements

    try:
        collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
        for element in collector:
            if element is not None:
                elements.append(element)
    except Exception:
        pass

    return elements


def get_category_id_value(category):
    if category is None:
        return ""

    try:
        return get_element_id_value(category.Id)
    except Exception:
        return ""


def get_element_category_id_value(element):
    if element is None:
        return ""

    try:
        category = element.Category
        if category is None:
            return ""
        return get_category_id_value(category)
    except Exception:
        return ""


def get_category_element_count_for_scope(category, scope):
    scope_text = normalize_text(scope)
    if not scope_text or scope_text == "Whole Model":
        return get_category_element_count(category)

    category_id_value = get_category_id_value(category)
    if category_id_value == "":
        return 0

    count = 0
    for element in get_elements_for_scope(scope_text):
        if get_element_category_id_value(element) == category_id_value:
            count += 1

    return count


def collect_model_categories(scope="Whole Model"):
    results = []
    excluded_category_names = set([
        "project information",
        "informacion del proyecto",
        "información del proyecto",
    ])

    try:
        categories = doc.Settings.Categories
    except Exception:
        return results

    for category in categories:
        try:
            if category is None:
                continue
            if getattr(category, "IsTagCategory", False):
                continue

            category_name = normalize_text(getattr(category, "Name", ""))
            if not category_name:
                continue
            category_key = category_name.strip().lower()

            if category_key in excluded_category_names:
                continue

            if category_key not in MODEL_CATEGORY_VISIBLE_NAMES:
                continue

            if category.CategoryType != DB.CategoryType.Model and category_key not in ("views", "vistas", "sheets", "hojas"):
                continue

            element_count = get_category_element_count_for_scope(category, scope)
            if element_count <= 0:
                continue

            results.append(CategoryItem(category, element_count))
        except Exception:
            continue

    return sorted(results, key=lambda x: x.Name.lower())


def collect_annotation_categories(scope="Whole Model"):
    results = []

    try:
        categories = doc.Settings.Categories
    except Exception:
        return results

    for category in categories:
        try:
            if category is None:
                continue

            category_name = normalize_text(getattr(category, "Name", ""))
            if not category_name:
                continue

            if category.CategoryType != DB.CategoryType.Annotation:
                continue

            element_count = get_category_element_count_for_scope(category, scope)
            if element_count <= 0:
                continue

            results.append(CategoryItem(category, element_count))
        except Exception:
            continue

    return sorted(results, key=lambda x: x.Name.lower())


def merge_category_items(category_items):
    by_key = {}

    for category_item in category_items or []:
        key = normalize_text(category_item.Name).lower()
        if key and key not in by_key:
            by_key[key] = category_item

    return sorted(by_key.values(), key=lambda x: x.Name.lower())


def collect_element_categories(scope="Whole Model"):
    return merge_category_items(collect_model_categories(scope) + collect_annotation_categories(scope))


def collect_spatial_categories():
    spatial_categories = []
    spatial_defs = [
        ("Rooms", "OST_Rooms"),
        ("Spaces", "OST_MEPSpaces"),
    ]

    for display_name, built_in_category_name in spatial_defs:
        built_in_category = getattr(DB.BuiltInCategory, built_in_category_name, None)
        if built_in_category is None:
            continue

        try:
            category = DB.Category.GetCategory(doc, built_in_category)
        except Exception:
            category = None

        if category is None:
            try:
                category = doc.Settings.Categories.get_Item(built_in_category)
            except Exception:
                category = None

        if category is None:
            continue

        element_count = get_category_element_count(category)
        if element_count <= 0:
            continue

        item = CategoryItem(category, element_count)
        item.Name = display_name
        spatial_categories.append(item)

    return spatial_categories


def get_element_display_name(element):
    if element is None:
        return ""

    try:
        value = normalize_text(getattr(element, "Name", ""))
        if value:
            return value
    except Exception:
        pass

    try:
        value = normalize_text(get_type_name_from_element(element))
        if value:
            return value
    except Exception:
        pass

    try:
        value = normalize_text(get_family_name_from_element(element))
        if value:
            return value
    except Exception:
        pass

    try:
        category_name = safe_text(element.Category.Name)
        if category_name:
            return "{} {}".format(category_name, safe_text(get_element_id_value(element)))
    except Exception:
        pass

    try:
        return "Element {}".format(safe_text(get_element_id_value(element)))
    except Exception:
        return "Element"


def get_element_items_from_elements(elements):
    items = []

    for element in list(elements or []):
        if element is not None:
            items.append(ElementItem(element))

    return sorted(items, key=lambda x: (safe_text(x.Name).lower(), safe_text(x.ElementId)))


def get_element_items_for_category(category):
    return get_element_items_from_elements(get_category_elements(category))


def get_element_items_for_category_scope(category, scope):
    return get_element_items_from_elements(get_category_elements_for_scope(category, scope))


def get_category_elements(category):
    elements = []
    if category is None:
        return elements

    category_name_key = get_category_name_key(category)
    if category_name_key in ("views", "vistas"):
        return get_view_parameter_scan_elements()
    if category_name_key in ("sheets", "hojas"):
        return get_sheet_parameter_scan_elements()

    try:
        collector = DB.FilteredElementCollector(doc).OfCategoryId(category.Id).WhereElementIsNotElementType()
        for element in collector:
            try:
                if element is not None:
                    elements.append(element)
            except Exception:
                pass
    except Exception:
        pass

    return elements


def get_current_selection_element_ids():
    selected_ids = set()

    try:
        selection_ids = uidoc.Selection.GetElementIds()
    except Exception:
        return selected_ids

    try:
        for element_id in selection_ids:
            element_id_value = get_element_id_value(element_id)
            if element_id_value != "":
                selected_ids.add(element_id_value)
    except Exception:
        pass

    return selected_ids


def get_active_view_element_ids():
    active_ids = set()

    try:
        active_view = doc.ActiveView
        if active_view is None:
            return active_ids
        collector = DB.FilteredElementCollector(doc, active_view.Id).WhereElementIsNotElementType()
        for element in collector:
            element_id_value = get_element_id_value(element)
            if element_id_value != "":
                active_ids.add(element_id_value)
    except Exception:
        pass

    return active_ids


def element_is_in_active_view(element):
    if element is None:
        return False

    element_id_value = get_element_id_value(element)
    if element_id_value == "":
        return False

    return element_id_value in get_active_view_element_ids()


def filter_elements_by_scope(elements, scope):
    scoped_elements = list(elements or [])
    scope_text = normalize_text(scope)

    if scope_text == "Current Selection":
        selected_ids = get_current_selection_element_ids()
        if not selected_ids:
            return []
        return [x for x in scoped_elements if get_element_id_value(x) in selected_ids]

    if scope_text == "Active View":
        active_ids = get_active_view_element_ids()
        if not active_ids:
            return []
        return [x for x in scoped_elements if get_element_id_value(x) in active_ids]

    return scoped_elements


def get_category_elements_for_scope(category, scope):
    scope_text = normalize_text(scope)

    if category is None:
        return []

    if scope_text == "Active View":
        try:
            active_view = doc.ActiveView
            if active_view is not None:
                elements = []
                collector = DB.FilteredElementCollector(doc, active_view.Id).OfCategoryId(category.Id).WhereElementIsNotElementType()
                for element in collector:
                    if element is not None:
                        elements.append(element)
                return elements
        except Exception:
            pass

    return filter_elements_by_scope(get_category_elements(category), scope_text)


def get_category_sample_elements(category, limit_count):
    elements = get_category_elements(category)
    if limit_count <= 0:
        return elements
    return elements[:limit_count]


def build_category_parameter_metadata(name, parameter_id_value, origin, editable):
    return {
        "Name": name,
        "Status": build_status(origin, editable),
        "Origin": origin,
        "Editable": editable,
        "ScheduleId": "",
        "UsedParams": [parameter_id_value] if parameter_id_value is not None else [],
        "FieldIndex": "",
        "ColumnRole": "CategoryField",
        "Hidden": False
    }


MODEL_CATEGORY_ALLOWED_DUPLICATE_NAMES = set([
    "detail number",
    "reference label",
    "sheet number",
    "type",
])

MODEL_CATEGORY_FULL_SCAN_NAMES = set([
    "views",
])

MODEL_CATEGORY_EXCLUDED_PARAMETER_NAMES = {
    "views": set([
        "depth cueing",
    ]),
}


def get_category_name_key(category):
    if category is None:
        return ""
    return normalize_text(getattr(category, "Name", "")).lower()


def get_category_parameter_scan_elements(category):
    category_name_key = get_category_name_key(category)
    if category_name_key in ("views", "vistas"):
        return get_view_parameter_scan_elements()
    if category_name_key in ("sheets", "hojas"):
        return get_sheet_parameter_scan_elements()
    if category_name_key in MODEL_CATEGORY_FULL_SCAN_NAMES:
        return get_category_elements(category)
    return get_category_sample_elements(category, MAX_SAMPLE_ELEMENTS)


def get_view_parameter_scan_elements():
    elements = []
    fallback_elements = []

    try:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType()
        for view in collector:
            try:
                if view is None:
                    continue
                if getattr(view, "IsTemplate", False):
                    continue
                fallback_elements.append(view)
                category_name_key = get_category_name_key(getattr(view, "Category", None))
                if category_name_key not in ("views", "vistas"):
                    continue
                elements.append(view)
            except Exception:
                continue
    except Exception:
        pass

    return elements if elements else fallback_elements


def get_sheet_parameter_scan_elements():
    elements = []

    try:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).WhereElementIsNotElementType()
        for sheet in collector:
            try:
                if sheet is not None:
                    elements.append(sheet)
            except Exception:
                continue
    except Exception:
        pass

    return elements


def should_skip_category_parameter(category, parameter_name):
    category_name_key = get_category_name_key(category)
    excluded_names = MODEL_CATEGORY_EXCLUDED_PARAMETER_NAMES.get(category_name_key, set())
    return normalize_text(parameter_name).lower() in excluded_names


def get_parameter_nonempty_count(elements, parameter_id_value):
    nonempty_count = 0
    metadata = build_category_parameter_metadata("", parameter_id_value, "", "")

    for element in elements:
        try:
            value = normalize_text(get_export_value_from_metadata(element, metadata))
            if value:
                nonempty_count += 1
        except Exception:
            continue

    return nonempty_count


def get_parameter_used_param_id(parameter_item):
    metadata = getattr(parameter_item, "Metadata", {}) or {}
    used_params = metadata.get("UsedParams", [])
    if not used_params:
        return 0

    try:
        return int(used_params[0])
    except Exception:
        return 0


def get_category_parameter_sort_key(parameter_item):
    metadata = getattr(parameter_item, "Metadata", {}) or {}
    nonempty_count = metadata.get("NonEmptyCount", 0)
    name_text = normalize_text(getattr(parameter_item, "Name", ""))
    origin_text = safe_text(getattr(parameter_item, "Origin", ""))
    used_param_id = abs(get_parameter_used_param_id(parameter_item))

    origin_rank = 2
    if origin_text == "Instance":
        origin_rank = 0
    elif origin_text == "Type":
        origin_rank = 1

    return (-nonempty_count, origin_rank, used_param_id, name_text.lower())


def dedupe_category_parameter_items(parameter_items):
    grouped = {}

    for item in parameter_items:
        name_key = normalize_text(getattr(item, "Name", "")).lower()
        if not name_key:
            continue
        grouped.setdefault(name_key, []).append(item)

    deduped = []

    for name_key, items in grouped.items():
        ranked_items = sorted(items, key=get_category_parameter_sort_key)
        max_items = 2 if name_key in MODEL_CATEGORY_ALLOWED_DUPLICATE_NAMES else 1
        deduped.extend(ranked_items[:max_items])

    return sorted(deduped, key=lambda x: x.Name.lower())


def iter_element_parameters(element):
    if element is None:
        return []

    try:
        ordered_params = element.GetOrderedParameters()
        if ordered_params:
            return list(ordered_params)
    except Exception:
        pass

    parameters = []
    try:
        for param in element.Parameters:
            parameters.append(param)
    except Exception:
        pass

    return parameters


def get_category_parameters(category):
    sample_elements = get_category_parameter_scan_elements(category)
    if not sample_elements:
        return []

    candidates = {}

    for element in sample_elements:
        if element is None:
            continue

        scan_targets = [
            ("Instance", element),
            ("Type", get_type_element(element))
        ]
        for target_origin, target in scan_targets:
            if target is None:
                continue

            try:
                for param in iter_element_parameters(target):
                    try:
                        if param is None:
                            continue

                        definition = getattr(param, "Definition", None)
                        if definition is None:
                            continue

                        param_name = normalize_text(getattr(definition, "Name", ""))
                        if not param_name:
                            continue

                        if should_skip_category_parameter(category, param_name):
                            continue

                        pid = get_parameter_builtin_id_value(param)
                        if pid is None:
                            continue

                        candidate_info = candidates.get(pid)
                        if candidate_info is None:
                            candidate_info = {
                                "Name": param_name,
                                "HasInstance": False,
                                "HasType": False,
                                "EditableYesCount": 0,
                                "EditableNoCount": 0,
                                "NonEmptyCount": 0
                            }
                            candidates[pid] = candidate_info

                        if target_origin == "Type":
                            candidate_info["HasType"] = True
                        else:
                            candidate_info["HasInstance"] = True

                        if getattr(param, "IsReadOnly", True):
                            candidate_info["EditableNoCount"] += 1
                        else:
                            candidate_info["EditableYesCount"] += 1

                        try:
                            if normalize_text(get_parameter_preview_value(param)):
                                candidate_info["NonEmptyCount"] += 1
                        except Exception:
                            pass
                    except Exception:
                        continue
            except Exception:
                pass

    parameter_items = []

    for parameter_id_value, candidate_info in candidates.items():
        has_instance = candidate_info.get("HasInstance", False)
        has_type = candidate_info.get("HasType", False)
        editable_yes_count = candidate_info.get("EditableYesCount", 0)
        editable_no_count = candidate_info.get("EditableNoCount", 0)

        if not has_instance and not has_type:
            continue

        if has_instance and has_type:
            origin = "Mixed"
        elif has_type:
            origin = "Type"
        else:
            origin = "Instance"

        if editable_yes_count and editable_no_count:
            editable = "Unknown"
        elif editable_yes_count:
            editable = "Yes"
        elif editable_no_count:
            editable = "No"
        else:
            editable = "Unknown"

        param_name = candidate_info.get("Name", "")
        metadata = build_category_parameter_metadata(param_name, parameter_id_value, origin, editable)
        metadata["NonEmptyCount"] = candidate_info.get("NonEmptyCount", 0)
        parameter_items.append(ParameterItem(param_name, origin, editable, metadata))

    return dedupe_category_parameter_items(parameter_items)


def get_field_name(field):
    try:
        if getattr(field, "ColumnHeading", None):
            text = normalize_text(field.ColumnHeading)
            if text and text != ".":
                return text
    except Exception:
        pass

    try:
        name = field.GetName()
        if name:
            text = normalize_text(name)
            if text and text != ".":
                return text
    except Exception:
        pass

    try:
        schedulable = field.GetSchedulableField()
        if schedulable:
            name = schedulable.GetName(doc)
            if name:
                text = normalize_text(name)
                if text and text != ".":
                    return text
    except Exception:
        pass

    return "Unnamed Field"


def get_field_type_text(field):
    try:
        return normalize_text(field.FieldType).lower()
    except Exception:
        return ""


def get_schedulable_field_safe(field):
    try:
        return field.GetSchedulableField()
    except Exception:
        return None


def get_has_schedulable_field(field):
    try:
        value = getattr(field, "HasSchedulableField", None)
        if value is not None:
            return bool(value)
    except Exception:
        pass

    schedulable = get_schedulable_field_safe(field)
    return schedulable is not None


def get_parameter_id_value(schedulable_field):
    if schedulable_field is None:
        return None

    try:
        parameter_id = schedulable_field.ParameterId
        if parameter_id is None:
            return None
    except Exception:
        return None

    value = get_element_id_value(parameter_id)
    if value == "":
        return None
    return value


def get_schedule_sample_elements(schedule, max_count):
    elements = []

    try:
        collector = DB.FilteredElementCollector(doc, schedule.Id).WhereElementIsNotElementType()
        for element in collector:
            if element is None:
                continue
            elements.append(element)
            if len(elements) >= max_count:
                break
    except Exception as ex:
        logger.warning(
            "Could not collect sample elements for '%s'. %s",
            get_schedule_name(schedule),
            safe_text(ex)
        )

    return elements


def get_schedule_elements(schedule):
    elements = []

    try:
        collector = DB.FilteredElementCollector(doc, schedule.Id).WhereElementIsNotElementType()
        for element in collector:
            if element is None:
                continue
            elements.append(element)
    except Exception as ex:
        logger.warning(
            "Could not collect schedule elements for '%s'. %s",
            get_schedule_name(schedule),
            safe_text(ex)
        )

    return elements


def get_type_element(element):
    if element is None:
        return None

    try:
        type_id = element.GetTypeId()
        if type_id is None:
            return None
    except Exception:
        return None

    try:
        if type_id == DB.ElementId.InvalidElementId:
            return None
    except Exception:
        pass

    try:
        if getattr(type_id, "IntegerValue", -1) < 0:
            return None
    except Exception:
        pass

    try:
        return doc.GetElement(type_id)
    except Exception:
        return None


def get_element_type_id_value(element):
    if element is None:
        return ""

    try:
        type_id = element.GetTypeId()
    except Exception:
        return ""

    if type_id is None:
        return ""

    try:
        if type_id == DB.ElementId.InvalidElementId:
            return ""
    except Exception:
        pass

    type_id_value = get_element_id_value(type_id)
    if type_id_value == "":
        return ""

    try:
        if int(type_id_value) < 0:
            return ""
    except Exception:
        pass

    return safe_text(type_id_value)


def find_parameter_by_element_id(element, parameter_id_value):
    if element is None or parameter_id_value is None:
        return None

    try:
        for param in element.Parameters:
            try:
                if param is None:
                    continue
                pid = getattr(param, "Id", None)
                if pid is None:
                    continue
                if get_element_id_value(pid) == parameter_id_value:
                    return param
            except Exception:
                continue
    except Exception:
        pass

    return None


def find_parameter_by_builtin(element, parameter_id_value):
    if element is None or parameter_id_value is None:
        return None

    if parameter_id_value >= 0:
        return None

    try:
        bip = DB.BuiltInParameter(parameter_id_value)
        param = element.get_Parameter(bip)
        if param is not None:
            return param
    except Exception:
        pass

    return None


def find_parameter_on_element(element, parameter_id_value):
    param = find_parameter_by_element_id(element, parameter_id_value)
    if param is not None:
        return param

    param = find_parameter_by_builtin(element, parameter_id_value)
    if param is not None:
        return param

    return None


def evaluate_parameter_on_element(element, parameter_id_value):
    if element is None:
        return None

    param = find_parameter_on_element(element, parameter_id_value)
    if param is not None:
        is_read_only = getattr(param, "IsReadOnly", True)
        return ("Instance", "No" if is_read_only else "Yes")

    type_element = get_type_element(element)
    param = find_parameter_on_element(type_element, parameter_id_value)
    if param is not None:
        is_read_only = getattr(param, "IsReadOnly", True)
        return ("Type", "No" if is_read_only else "Yes")

    return None


def resolve_field_from_schedule_context(schedule, schedulable):
    parameter_id_value = get_parameter_id_value(schedulable)
    if parameter_id_value is None:
        return None

    sample_elements = get_schedule_sample_elements(schedule, MAX_SAMPLE_ELEMENTS)
    if not sample_elements:
        return None

    origin_counts = {}
    editable_counts = {}

    for element in sample_elements:
        try:
            result = evaluate_parameter_on_element(element, parameter_id_value)
            if result is None:
                continue

            origin, editable = result
            origin_counts[origin] = origin_counts.get(origin, 0) + 1
            editable_counts[editable] = editable_counts.get(editable, 0) + 1
        except Exception:
            pass

    if not origin_counts:
        return None

    origin = sorted(origin_counts.items(), key=lambda x: (-x[1], x[0]))[0][0]
    editable = sorted(editable_counts.items(), key=lambda x: (-x[1], x[0]))[0][0]

    if len(editable_counts) > 1:
        editable = "Unknown"

    if len(origin_counts) > 1:
        origin = "Mixed"

    return (origin, editable)


def analyze_field(schedule, field):
    field_type_text = get_field_type_text(field)

    try:
        if getattr(field, "IsCalculatedField", False):
            return ("Calculated", "No")
    except Exception:
        pass

    try:
        if getattr(field, "IsCombinedParameterField", False):
            return ("Combined", "No")
    except Exception:
        pass

    schedulable = get_schedulable_field_safe(field)
    has_schedulable = get_has_schedulable_field(field)

    if not has_schedulable and schedulable is None:
        return ("Special", "No")

    readonly_tokens = (
        "formula",
        "count",
        "percentage",
        "percent",
        "room",
        "fromroom",
        "toroom",
        "projectinfo",
        "materialquantity",
        "revitlinkinstance",
        "revitlinktype"
    )
    for token in readonly_tokens:
        if token in field_type_text:
            return ("Special", "No")

    validated = resolve_field_from_schedule_context(schedule, schedulable)
    if validated is not None:
        return validated

    if "elementtype" in field_type_text:
        return ("Type", "Unknown")

    if "instance" in field_type_text:
        return ("Instance", "Unknown")

    if field_type_text == "type":
        return ("Type", "Unknown")

    parameter_id_value = get_parameter_id_value(schedulable)
    if parameter_id_value is not None and parameter_id_value < 0:
        return ("Special", "Unknown")

    if schedulable is not None:
        return ("Special", "Unknown")

    return ("Special", "Unknown")


def get_visible_schedule_fields(schedule):
    visible_fields = []

    try:
        definition = schedule.Definition
        field_count = definition.GetFieldCount()
    except Exception as ex:
        logger.warning(
            "Could not read definition for '%s'. %s",
            get_schedule_name(schedule),
            safe_text(ex)
        )
        return visible_fields

    visible_index = 0

    for i in range(field_count):
        try:
            field = definition.GetField(i)

            try:
                if getattr(field, "IsHidden", False):
                    continue
            except Exception:
                pass

            field_name = get_field_name(field)
            origin, editable = analyze_field(schedule, field)
            schedulable = get_schedulable_field_safe(field)
            parameter_id_value = get_parameter_id_value(schedulable)

            metadata = {
                "Name": field_name,
                "Status": build_status(origin, editable),
                "Origin": origin,
                "Editable": editable,
                "ScheduleId": getattr(schedule.Id, "IntegerValue", 0),
                "UsedParams": [parameter_id_value] if parameter_id_value is not None else [],
                "FieldIndex": i,
                "VisibleColumnIndex": visible_index
            }

            visible_fields.append({
                "field": field,
                "name": field_name,
                "origin": origin,
                "editable": editable,
                "status": build_status(origin, editable),
                "metadata": metadata
            })

            visible_index += 1

        except Exception as ex:
            logger.warning(
                "Skipping field index %s in '%s'. %s",
                i,
                get_schedule_name(schedule),
                safe_text(ex).split("\n")[0]
            )

    return visible_fields


def get_schedule_parameters(schedule):
    results = []

    visible_fields = get_visible_schedule_fields(schedule)
    for item in visible_fields:
        results.append(ParameterItem(item["name"], item["origin"], item["editable"]))

    def schedule_parameter_sort_key(item):
        editable = safe_text(getattr(item, "Editable", ""))
        if editable == "Yes":
            editable_rank = 0
        elif editable == "Unknown":
            editable_rank = 1
        else:
            editable_rank = 2
        return (editable_rank, getattr(item, "Name", "").lower())

    return sorted(results, key=schedule_parameter_sort_key)


def build_schedule_export_options():
    options = DB.ViewScheduleExportOptions()
    try:
        options.FieldDelimiter = EXPORT_DELIMITER
    except Exception:
        pass
    return options


def read_csv_rows(csv_path):
    rows = []
    with io.open(csv_path, "r", encoding="utf-8-sig") as fp:
        reader = csv.reader(fp, delimiter=EXPORT_DELIMITER)
        for row in reader:
            rows.append([normalize_text(x) for x in row])
    return rows


def release_com_object(obj):
    if THIS_DIR not in sys.path:
        sys.path.append(THIS_DIR)
    import excel_utils
    obj = unwrap_com_value(obj)
    return excel_utils.release_com_object(obj, Marshal)


def get_status_fill_color(status_text):
    text = safe_text(status_text).lower()
    if "locked" in text:
        return 0xACC7F7
    if "unknown" in text:
        return 0xF2F2F2
    return None


def fit_export_columns(worksheet, export_columns, min_width, max_width, last_data_row):
    total_cols = len(export_columns)

    for i in range(1, total_cols + 1):
        try:
            col = worksheet.Columns[i]
            role = export_columns[i - 1].get("ColumnRole", "")
            hidden = export_columns[i - 1].get("Hidden", False)

            if hidden:
                col.Hidden = True
                continue

            if role == "UniqueId":
                col.Hidden = True
                continue

            if role == "ElementId":
                col.ColumnWidth = 12
                continue

            if last_data_row >= 2:
                fit_range = worksheet.Range[
                    worksheet.Cells[2, i],
                    worksheet.Cells[last_data_row, i]
                ]
                fit_range.Columns.AutoFit()

            width = col.ColumnWidth

            if width < min_width:
                col.ColumnWidth = min_width
            elif width > max_width:
                col.ColumnWidth = max_width

        except Exception:
            pass


def get_visible_field_headers(schedule):
    headers = []
    try:
        definition = schedule.Definition
        for i in range(definition.GetFieldCount()):
            field = definition.GetField(i)
            try:
                if getattr(field, "IsHidden", False):
                    continue
            except Exception:
                pass
            headers.append(get_field_name(field))
    except Exception:
        pass
    return headers


def has_visible_elementid_field(schedule):
    headers = get_visible_field_headers(schedule)
    for h in headers:
        if normalize_header(h) == "elementid":
            return True
    return False


def find_elementid_schedulable_field(schedule):
    try:
        definition = schedule.Definition
        schedulable_fields = definition.GetSchedulableFields()
    except Exception:
        return None

    for sf in schedulable_fields:
        try:
            name = normalize_text(sf.GetName(doc))
            if normalize_header(name) == "elementid":
                return sf
        except Exception:
            continue

    return None


def add_temp_elementid_field(schedule):
    if has_visible_elementid_field(schedule):
        return None

    schedulable_field = find_elementid_schedulable_field(schedule)
    if schedulable_field is None:
        return None

    tx = DB.Transaction(doc, "Temp add Element ID for export")
    tx.Start()
    try:
        definition = schedule.Definition
        new_field = definition.AddField(schedulable_field)
        field_id = new_field.FieldId
        tx.Commit()
        return field_id
    except Exception:
        try:
            tx.RollBack()
        except Exception:
            pass
        return None


def remove_temp_field(schedule, field_id):
    if field_id is None:
        return

    tx = DB.Transaction(doc, "Remove temp Element ID after export")
    tx.Start()
    try:
        schedule.Definition.RemoveField(field_id)
        tx.Commit()
    except Exception as ex:
        try:
            tx.RollBack()
        except Exception:
            pass
        logger.warning("Could not remove temp field. %s", safe_text(ex))

def capture_schedule_grouping_state(schedule):
    state = {
        "IsItemized": True,
        "Fields": []
    }

    try:
        definition = schedule.Definition
        state["IsItemized"] = definition.IsItemized

        sort_count = definition.GetSortGroupFieldCount()
        for i in range(sort_count):
            sgf = definition.GetSortGroupField(i)
            state["Fields"].append({
                "ShowHeader": sgf.ShowHeader,
                "ShowFooter": sgf.ShowFooter,
                "ShowBlankLine": sgf.ShowBlankLine,
                "ShowFooterCount": sgf.ShowFooterCount,
                "ShowFooterTitle": sgf.ShowFooterTitle
            })
    except Exception:
        pass

    return state


def flatten_schedule_for_export(schedule):
    tx = DB.Transaction(doc, "Flatten schedule for export")
    tx.Start()
    try:
        definition = schedule.Definition
        definition.IsItemized = True

        sort_count = definition.GetSortGroupFieldCount()
        for i in range(sort_count):
            sgf = definition.GetSortGroupField(i)

            try:
                sgf.ShowHeader = False
            except Exception:
                pass

            try:
                sgf.ShowFooter = False
            except Exception:
                pass

            try:
                sgf.ShowBlankLine = False
            except Exception:
                pass

            try:
                sgf.ShowFooterCount = False
            except Exception:
                pass

            try:
                sgf.ShowFooterTitle = False
            except Exception:
                pass

            definition.SetSortGroupField(i, sgf)

        tx.Commit()
        return True
    except Exception:
        try:
            tx.RollBack()
        except Exception:
            pass
        return False


def restore_schedule_grouping_state(schedule, state):
    if not state:
        return

    tx = DB.Transaction(doc, "Restore schedule grouping")
    tx.Start()
    try:
        definition = schedule.Definition
        definition.IsItemized = state.get("IsItemized", True)

        saved_fields = state.get("Fields", [])
        sort_count = definition.GetSortGroupFieldCount()

        for i in range(min(sort_count, len(saved_fields))):
            sgf = definition.GetSortGroupField(i)
            saved = saved_fields[i]

            try:
                sgf.ShowHeader = saved.get("ShowHeader", False)
            except Exception:
                pass

            try:
                sgf.ShowFooter = saved.get("ShowFooter", False)
            except Exception:
                pass

            try:
                sgf.ShowBlankLine = saved.get("ShowBlankLine", False)
            except Exception:
                pass

            try:
                sgf.ShowFooterCount = saved.get("ShowFooterCount", False)
            except Exception:
                pass

            try:
                sgf.ShowFooterTitle = saved.get("ShowFooterTitle", False)
            except Exception:
                pass

            definition.SetSortGroupField(i, sgf)

        tx.Commit()
    except Exception:
        try:
            tx.RollBack()
        except Exception:
            pass

def export_schedule_to_temp_csv(schedule):
    temp_field_id = add_temp_elementid_field(schedule)
    grouping_state = capture_schedule_grouping_state(schedule)

    schedule_name = get_schedule_name(schedule)
    file_name = "{}_{}.csv".format(
        sanitize_filename(schedule_name),
        str(int(time.time() * 1000))
    )
    full_path = os.path.join(TEMP_FOLDER, file_name)

    try:
        flatten_schedule_for_export(schedule)

        options = build_schedule_export_options()
        schedule.Export(TEMP_FOLDER, file_name, options)
    finally:
        restore_schedule_grouping_state(schedule, grouping_state)
        remove_temp_field(schedule, temp_field_id)

    return full_path


def find_header_row(csv_rows):
    best_index = -1
    best_score = -1

    for i, row in enumerate(csv_rows):
        cleaned = [normalize_text(x) for x in row]
        non_empty = [x for x in cleaned if x]
        if not non_empty:
            continue

        score = len(non_empty)

        has_element_id = False
        for cell in cleaned:
            if normalize_header(cell) == "elementid":
                has_element_id = True
                score += 100
                break

        if has_element_id:
            return i

        if score > best_score:
            best_score = score
            best_index = i

    return best_index


def find_element_id_column_index(headers):
    for i, head in enumerate(headers):
        if normalize_header(head) == "elementid":
            return i
    return -1

def find_column_index_by_header(headers, target_name):
    target = normalize_header(target_name)
    for i, head in enumerate(headers):
        if normalize_header(head) == target:
            return i
    return -1


def row_has_real_schedule_content(row, csv_headers):
    if not row:
        return False

    view_name_idx = find_column_index_by_header(csv_headers, "View Name")
    title_on_sheet_idx = find_column_index_by_header(csv_headers, "Title on Sheet")
    sheet_number_idx = find_column_index_by_header(csv_headers, "Sheet Number")

    # Si tiene alguno de estos campos con valor, la tratamos como fila real
    candidate_indexes = [
        view_name_idx,
        title_on_sheet_idx,
        sheet_number_idx
    ]

    for idx in candidate_indexes:
        if 0 <= idx < len(row):
            if normalize_text(row[idx]):
                return True

    return False

def build_unique_id_map_from_rows(data_rows, element_id_col_index):
    unique_ids = []
    has_valid_ids = False

    for row in data_rows:
        value = ""
        if 0 <= element_id_col_index < len(row):
            value = normalize_text(row[element_id_col_index])

        unique_id = ""
        if value:
            try:
                element_id_int = int(float(value))
                elem = doc.GetElement(DB.ElementId(element_id_int))
                if elem is not None:
                    unique_id = safe_text(elem.UniqueId)
                    has_valid_ids = True
            except Exception:
                pass

        unique_ids.append(unique_id)

    return unique_ids, has_valid_ids
def build_id_data_from_schedule_elements(schedule, row_count):
    elements = get_schedule_elements(schedule)

    element_id_values = ["" for _ in range(row_count)]
    unique_id_values = ["" for _ in range(row_count)]

    # The collector order is not guaranteed to match exported schedule rows,
    # so we do not assign technical IDs from row position.
    return element_id_values, unique_id_values, False, len(elements)
def build_schema_rows(export_columns):
    rows = []
    index = 1

    for col in export_columns:
        rows.append([
            index,
            safe_text(col.get("ExcelIndex", "")),
            safe_text(col.get("Name", "")),
            safe_text(col.get("Status", "")),
            safe_text(col.get("Origin", "")),
            safe_text(col.get("Editable", "")),
            safe_text(col.get("ScheduleId", "")),
            json.dumps(col.get("UsedParams", [])),
            safe_text(col.get("FieldIndex", "")),
            safe_text(col.get("ColumnRole", "")),
            safe_text(col.get("Hidden", ""))
        ])
        index += 1

    return rows


def make_excel_sheet_name(name, fallback_name):
    if THIS_DIR not in sys.path:
        sys.path.append(THIS_DIR)
    import excel_utils
    return excel_utils.make_excel_sheet_name(name, fallback_name, normalize_text)


def write_matrix_to_range(worksheet, start_row, start_col, matrix):
    if THIS_DIR not in sys.path:
        sys.path.append(THIS_DIR)
    import excel_utils
    return excel_utils.write_matrix_to_range(worksheet, start_row, start_col, matrix, Array, Object)


def ensure_workbook_sheet_count(workbook, required_count):
    if THIS_DIR not in sys.path:
        sys.path.append(THIS_DIR)
    import excel_utils
    return excel_utils.ensure_workbook_sheet_count(workbook, required_count)


def open_path_with_default_app(full_path):
    if THIS_DIR not in sys.path:
        sys.path.append(THIS_DIR)
    import excel_utils
    return excel_utils.open_path_with_default_app(full_path)

def get_project_standards_section_label(section_key):
    for key, label in PROJECT_STANDARDS_SECTIONS:
        if key == section_key:
            return label
    return safe_text(section_key)


def get_enum_label(enum_value):
    if enum_value is None:
        return ""

    try:
        return safe_text(DB.LabelUtils.GetLabelFor(enum_value))
    except Exception:
        return safe_text(enum_value)


def get_definition_group_label(definition):
    if definition is None:
        return ""

    try:
        return get_enum_label(definition.ParameterGroup)
    except Exception:
        return ""


def get_definition_type_label(definition):
    if definition is None:
        return ""

    try:
        param_type = definition.ParameterType
        if param_type is not None:
            return safe_text(param_type)
    except Exception:
        pass

    try:
        data_type = definition.GetDataType()
        if data_type is not None:
            return safe_text(data_type.TypeId)
    except Exception:
        pass

    return ""


def get_binding_type_label(binding):
    if binding is None:
        return ""

    try:
        if isinstance(binding, DB.InstanceBinding):
            return "Instance"
    except Exception:
        pass

    try:
        if isinstance(binding, DB.TypeBinding):
            return "Type"
    except Exception:
        pass

    return safe_text(type(binding).__name__)


def get_binding_categories_label(binding):
    if binding is None:
        return ""

    names = []
    try:
        categories = binding.Categories
    except Exception:
        categories = None

    if categories is None:
        return ""

    try:
        for category in categories:
            category_name = safe_text(getattr(category, "Name", ""))
            if category_name:
                names.append(category_name)
    except Exception:
        pass

    return ", ".join(sorted(set(names)))


def get_bool_label(value):
    return "Yes" if value else "No"


def get_color_label(color_value):
    if color_value is None:
        return ""

    try:
        return "{},{},{}".format(color_value.Red, color_value.Green, color_value.Blue)
    except Exception:
        return ""


def get_line_pattern_name(pattern_id):
    try:
        if pattern_id is None or pattern_id == DB.ElementId.InvalidElementId:
            return ""
    except Exception:
        pass

    try:
        pattern_element = doc.GetElement(pattern_id)
        if pattern_element is not None:
            return safe_text(pattern_element.Name)
    except Exception:
        pass

    return ""


def get_category_material_name(category):
    if category is None:
        return ""

    try:
        material = category.Material
        if material is not None:
            return safe_text(material.Name)
    except Exception:
        pass

    try:
        material_id = category.Material.Id
        if material_id is not None and material_id != DB.ElementId.InvalidElementId:
            material = doc.GetElement(material_id)
            if material is not None:
                return safe_text(material.Name)
    except Exception:
        pass

    return ""


def get_category_line_weight(category, graphics_style_type):
    if category is None:
        return ""

    try:
        return safe_text(category.GetLineWeight(graphics_style_type))
    except Exception:
        return ""


def build_project_standards_plain_sheet(headers, data_rows):
    rows = [list(headers or [])]
    for data_row in data_rows or []:
        rows.append(list(data_row))
    return rows


def get_integer_id_text(element_or_id):
    if element_or_id is None:
        return ""
    try:
        return safe_text(element_or_id.IntegerValue)
    except Exception:
        pass
    try:
        return safe_text(element_or_id.Id.IntegerValue)
    except Exception:
        return ""


def get_color_label_spaced(color_value):
    if color_value is None:
        return ""
    try:
        return "{0}, {1}, {2}".format(color_value.Red, color_value.Green, color_value.Blue)
    except Exception:
        return safe_text(color_value)


def get_category_bucket_key(category):
    if category is None:
        return ""
    try:
        category_type = category.CategoryType
        if category_type == DB.CategoryType.Model:
            return "model"
        if category_type == DB.CategoryType.Annotation:
            return "annotation"
        if category_type == DB.CategoryType.AnalyticalModel:
            return "analytical"
    except Exception:
        pass

    category_type_text = normalize_text(getattr(category, "CategoryType", ""))
    if "annotation" in category_type_text:
        return "annotation"
    if "analytical" in category_type_text:
        return "analytical"
    if "model" in category_type_text:
        return "model"
    return ""


def make_hierarchical_child_name(child_name):
    return "       |---- {0}".format(safe_text(child_name))


def collect_object_style_rows(bucket_key):
    rows = []
    try:
        categories = list(doc.Settings.Categories)
    except Exception:
        categories = []

    categories.sort(key=lambda item: normalize_text(getattr(item, "Name", "")))

    for category in categories:
        if get_category_bucket_key(category) != bucket_key:
            continue

        include_cut = bucket_key == "model"
        parent_row = [
            get_integer_id_text(category),
            safe_text(getattr(category, "Name", "")),
            get_category_line_weight(category, DB.GraphicsStyleType.Projection),
        ]
        if include_cut:
            parent_row.append(get_category_line_weight(category, DB.GraphicsStyleType.Cut))
        parent_row.append(get_color_label_spaced(getattr(category, "LineColor", None)))
        rows.append(parent_row)

        try:
            subcategories = list(category.SubCategories)
        except Exception:
            subcategories = []

        subcategories.sort(key=lambda item: normalize_text(getattr(item, "Name", "")))
        for subcategory in subcategories:
            child_row = [
                get_integer_id_text(subcategory),
                make_hierarchical_child_name(getattr(subcategory, "Name", "")),
                get_category_line_weight(subcategory, DB.GraphicsStyleType.Projection),
            ]
            if include_cut:
                child_row.append(get_category_line_weight(subcategory, DB.GraphicsStyleType.Cut))
            child_row.append(get_color_label_spaced(getattr(subcategory, "LineColor", None)))
            rows.append(child_row)

    return rows


def build_object_style_sheet_rows(bucket_key):
    include_cut = bucket_key == "model"
    headers = ["Id", "Name", "Projection"]
    if include_cut:
        headers.append("Cut")
    headers.append("Color")
    return build_project_standards_plain_sheet(headers, collect_object_style_rows(bucket_key))


def build_project_information_rows():
    rows = [[
        "Parameter Name",
        "Value",
        "Group",
        "Storage Type",
        "Read Only",
    ]]

    project_info = getattr(doc, "ProjectInformation", None)
    if project_info is None:
        return rows

    try:
        parameters = project_info.Parameters
    except Exception:
        parameters = []

    for param in parameters:
        try:
            definition = param.Definition
        except Exception:
            definition = None

        rows.append([
            safe_text(getattr(definition, "Name", "")),
            get_parameter_preview_value(param),
            get_definition_group_label(definition),
            safe_text(getattr(param, "StorageType", "")),
            get_bool_label(getattr(param, "IsReadOnly", True)),
        ])

    return rows


def build_project_parameters_rows():
    rows = [[
        "Parameter Name",
        "Type of Parameter",
        "Group Parameter Under",
        "Binding",
        "Categories",
        "Shared",
        "GUID",
        "Visible",
        "User Modifiable",
    ]]

    try:
        iterator = doc.ParameterBindings.ForwardIterator()
        iterator.Reset()
    except Exception:
        return rows

    while iterator.MoveNext():
        try:
            definition = iterator.Key
            binding = iterator.Current
        except Exception:
            continue

        shared_guid = ""
        shared_value = False
        visible_value = True
        user_modifiable_value = True

        try:
            guid_value = definition.GUID
            shared_guid = safe_text(guid_value)
            shared_value = bool(shared_guid)
        except Exception:
            shared_guid = ""

        try:
            visible_value = bool(definition.Visible)
        except Exception:
            visible_value = True

        try:
            user_modifiable_value = bool(definition.UserModifiable)
        except Exception:
            user_modifiable_value = True

        rows.append([
            safe_text(getattr(definition, "Name", "")),
            get_definition_type_label(definition),
            get_definition_group_label(definition),
            get_binding_type_label(binding),
            get_binding_categories_label(binding),
            get_bool_label(shared_value),
            shared_guid,
            get_bool_label(visible_value),
            get_bool_label(user_modifiable_value),
        ])

    return rows


def build_object_styles_rows():
    rows = [[
        "Category",
        "Subcategory",
        "Category Type",
        "Projection Weight",
        "Cut Weight",
        "Color",
        "Material",
    ]]

    try:
        categories = doc.Settings.Categories
    except Exception:
        categories = []

    for category in categories:
        category_name = safe_text(getattr(category, "Name", ""))
        category_type = safe_text(getattr(category, "CategoryType", ""))
        rows.append([
            category_name,
            "",
            category_type,
            get_category_line_weight(category, DB.GraphicsStyleType.Projection),
            get_category_line_weight(category, DB.GraphicsStyleType.Cut),
            get_color_label(getattr(category, "LineColor", None)),
            get_category_material_name(category),
        ])

        try:
            subcategories = category.SubCategories
        except Exception:
            subcategories = []

        try:
            for subcategory in subcategories:
                rows.append([
                    category_name,
                    safe_text(getattr(subcategory, "Name", "")),
                    category_type,
                    get_category_line_weight(subcategory, DB.GraphicsStyleType.Projection),
                    get_category_line_weight(subcategory, DB.GraphicsStyleType.Cut),
                    get_color_label(getattr(subcategory, "LineColor", None)),
                    get_category_material_name(subcategory),
                ])
        except Exception:
            pass

    return rows


def build_line_styles_rows():
    rows = [[
        "Line Style",
        "Projection Weight",
        "Color",
        "Line Pattern",
    ]]

    try:
        lines_category = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
        subcategories = lines_category.SubCategories
    except Exception:
        subcategories = []

    try:
        for subcategory in subcategories:
            rows.append([
                safe_text(getattr(subcategory, "Name", "")),
                get_category_line_weight(subcategory, DB.GraphicsStyleType.Projection),
                get_color_label(getattr(subcategory, "LineColor", None)),
                get_line_pattern_name(subcategory.GetLinePatternId(DB.GraphicsStyleType.Projection)),
            ])
    except Exception:
        pass

    return rows


def build_families_rows():
    rows = [[
        "Family Name",
        "Category",
        "Type Count",
        "In Place",
        "Editable",
    ]]

    try:
        families = DB.FilteredElementCollector(doc).OfClass(DB.Family)
    except Exception:
        families = []

    for family in families:
        type_count = 0
        try:
            type_count = len(list(family.GetFamilySymbolIds()))
        except Exception:
            type_count = 0

        rows.append([
            safe_text(getattr(family, "Name", "")),
            safe_text(getattr(getattr(family, "FamilyCategory", None), "Name", "")),
            safe_text(type_count),
            get_bool_label(getattr(family, "IsInPlace", False)),
            get_bool_label(getattr(family, "IsEditable", False)),
        ])

    return rows


def build_project_standards_instructions_rows(selected_section_keys):
    selected_labels = []
    for section_key in selected_section_keys:
        selected_labels.append(get_project_standards_section_label(section_key))

    return [
        ["SheetLink - pyMENVIC"],
        ["Project Standards Export"],
        [""],
        ["Selected sections:"],
        [", ".join(selected_labels) if selected_labels else "None"],
        [""],
        ["Notes:"],
        ["This workbook documents project standards exported from the current Revit model."],
        ["Project Parameters lists parameter definitions and their bindings, not per-element values."],
        ["Editability and import back into Revit are not part of this standards export workflow yet."],
    ]


def get_project_standards_rows(section_key):
    if section_key == "project_information":
        return build_project_information_rows()
    if section_key == "project_parameters":
        return build_project_parameters_rows()
    if section_key == "object_styles":
        return [["Info"], ["Object styles now export through split worksheets."]]
    if section_key == "line_styles":
        return build_line_styles_rows()
    if section_key == "families":
        return build_families_rows()
    return [["Info"], ["Section not available"]]


def format_project_standards_sheet(worksheet, row_count, col_count):
    if worksheet is None or row_count <= 0 or col_count <= 0:
        return

    try:
        used_range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[row_count, col_count]]
        used_range.Borders.LineStyle = 1
    except Exception:
        pass

    try:
        header_range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, col_count]]
        header_range.Font.Bold = True
        header_range.Font.Color = 0xFFFFFF
        header_range.Interior.Color = 0x5B7D95
        header_range.RowHeight = 24
    except Exception:
        pass

    try:
        worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[max(1, row_count), col_count]].AutoFilter()
    except Exception:
        pass

    try:
        worksheet.Application.ActiveWindow.SplitRow = 1
        worksheet.Application.ActiveWindow.FreezePanes = True
    except Exception:
        pass

    for col_index in range(1, col_count + 1):
        try:
            worksheet.Columns[col_index].AutoFit()
            if worksheet.Columns[col_index].ColumnWidth > 36:
                worksheet.Columns[col_index].ColumnWidth = 36
        except Exception:
            pass


def format_project_standards_instructions_sheet(worksheet, row_count, col_count):
    if worksheet is None or row_count <= 0 or col_count <= 0:
        return

    try:
        title_range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, 1]]
        title_range.Font.Bold = True
    except Exception:
        pass

    try:
        worksheet.Columns[1].ColumnWidth = 90
        worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[row_count, 1]].WrapText = True
    except Exception:
        pass


def export_project_standards_to_xlsx(full_path, selected_section_keys, open_after_export):
    if THIS_DIR not in sys.path:
        sys.path.append(THIS_DIR)
    import imp
    module_path = os.path.join(THIS_DIR, "project_standards_export.py")
    project_standards_export = imp.load_source("project_standards_export_active", module_path)
    return project_standards_export.export_project_standards_to_xlsx(
        full_path,
        selected_section_keys,
        open_after_export,
        globals(),
    )


class ProjectStandardsWindow(forms.WPFWindow):
    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)
        self.result_confirmed = False
        self.open_after_export = True
        self.selected_sections = []

        try:
            self.chkOpenAfterExport.IsChecked = True
            self.chkProjectInformation.IsChecked = True
            self.chkProjectParameters.IsChecked = True
            self.chkObjectStyles.IsChecked = True
            self.chkAnnotationObjects.IsChecked = True
            self.chkAnalyticalModelObjects.IsChecked = True
            self.chkLineStyles.IsChecked = True
            self.chkFamilies.IsChecked = True
            self.btnStandardsGoogle.IsEnabled = False
        except Exception:
            pass

        self.btnStandardsExcel.Click += self.on_excel_clicked
        self.btnStandardsCancel.Click += self.on_cancel_clicked

    def _collect_selected_sections(self):
        selections = []

        mapping = [
            ("project_information", getattr(self, "chkProjectInformation", None)),
            ("project_parameters", getattr(self, "chkProjectParameters", None)),
            ("line_styles", getattr(self, "chkLineStyles", None)),
            ("families", getattr(self, "chkFamilies", None)),
        ]

        for section_key, checkbox in mapping:
            try:
                if checkbox is not None and checkbox.IsChecked:
                    selections.append(section_key)
            except Exception:
                pass

        object_checkboxes = [
            getattr(self, "chkObjectStyles", None),
            getattr(self, "chkAnnotationObjects", None),
            getattr(self, "chkAnalyticalModelObjects", None),
        ]
        for checkbox in object_checkboxes:
            try:
                if checkbox is not None and checkbox.IsChecked:
                    selections.insert(2, "object_styles")
                    break
            except Exception:
                pass

        return selections

    def on_excel_clicked(self, sender, args):
        selected_sections = self._collect_selected_sections()
        if not selected_sections:
            TaskDialog.Show(__title__, "Select at least one Project Standards category to export.")
            return

        self.open_after_export = bool(getattr(self.chkOpenAfterExport, "IsChecked", True))
        self.selected_sections = selected_sections
        self.result_confirmed = True
        self.Close()

    def on_cancel_clicked(self, sender, args):
        self.result_confirmed = False
        self.Close()

def get_cell_value(worksheet, row, col):
    try:
        return worksheet.Cells[row, col].Value2
    except Exception:
        return None


def get_cell_text(worksheet, row, col):
    try:
        return worksheet.Cells[row, col].Text
    except Exception:
        value = get_cell_value(worksheet, row, col)
        return safe_text(value)


def get_worksheet_by_name(workbook, sheet_name):
    try:
        return workbook.Worksheets[sheet_name]
    except Exception:
        return None


def get_last_used_row_col(worksheet):
    used_range = None
    try:
        used_range = worksheet.UsedRange
        start_row = used_range.Row
        start_col = used_range.Column
        row_count = used_range.Rows.Count
        col_count = used_range.Columns.Count

        last_row = start_row + row_count - 1
        last_col = start_col + col_count - 1
        return last_row, last_col
    finally:
        release_com_object(used_range)


def parse_metadata_cell(value):
    text = normalize_text(value)
    if not text:
        return None

    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return data
    except Exception:
        pass

    return None


def get_metadata_column_by_role(metadata_by_col, role_name):
    role_text = normalize_text(role_name)
    if not role_text:
        return None

    for col_index, metadata in metadata_by_col.items():
        try:
            if metadata is None:
                continue
            if normalize_text(metadata.get("ColumnRole", "")) == role_text:
                return col_index
        except Exception:
            continue

    return None


def get_import_row_reference_value(row_info, role_name, fallback_col):
    values = row_info.get("Values", {})
    metadata_by_col = row_info.get("MetadataByCol", {})
    col_index = get_metadata_column_by_role(metadata_by_col, role_name)

    if col_index is None:
        col_index = fallback_col

    return normalize_text(values.get(col_index, ""))


def get_parameter_preview_value(param):
    if param is None:
        return ""

    try:
        storage_type = param.StorageType
    except Exception:
        storage_type = None

    try:
        if storage_type == DB.StorageType.String:
            return normalize_text(param.AsString())
    except Exception:
        pass

    try:
        value_string = param.AsValueString()
        if value_string not in (None, ""):
            return normalize_text(value_string)
    except Exception:
        pass

    try:
        if storage_type == DB.StorageType.Integer:
            return normalize_text(param.AsInteger())
    except Exception:
        pass

    try:
        if storage_type == DB.StorageType.Double:
            return normalize_text(param.AsDouble())
    except Exception:
        pass

    try:
        if storage_type == DB.StorageType.ElementId:
            eid = param.AsElementId()
            if eid is not None:
                eid_value = get_element_id_value(eid)
                if eid_value != "" and eid_value < 0:
                    return "None"
                return safe_text(eid_value)
    except Exception:
        pass

    return ""


def is_annotation_element(element):
    try:
        category = getattr(element, "Category", None)
        if category is None:
            return False
        return category.CategoryType == DB.CategoryType.Annotation
    except Exception:
        return False


def is_zero_number_text(value):
    text = normalize_text(value)
    if not text:
        return False

    try:
        return abs(float(text)) < 0.0000001
    except Exception:
        return False


def get_type_name_from_element(element):
    type_element = get_type_element(element)
    if type_element is None:
        return ""

    try:
        type_param = find_parameter_by_name(type_element, "Type")
        type_value = get_parameter_preview_value(type_param)
        if normalize_text(type_value):
            return normalize_text(type_value)
    except Exception:
        pass

    try:
        return normalize_text(type_element.Name)
    except Exception:
        pass

    try:
        return normalize_text(getattr(type_element, "Name", ""))
    except Exception:
        return ""


def get_family_name_from_element(element):
    type_element = get_type_element(element)
    if type_element is None:
        return ""

    try:
        family_name = normalize_text(getattr(type_element, "FamilyName", ""))
        if family_name:
            return family_name
    except Exception:
        pass

    try:
        family = getattr(type_element, "Family", None)
        family_name = normalize_text(getattr(family, "Name", ""))
        if family_name:
            return family_name
    except Exception:
        pass

    return ""


def find_parameter_by_name(element, target_name):
    if element is None:
        return None

    name_text = normalize_text(target_name)
    if not name_text:
        return None

    try:
        param = element.LookupParameter(name_text)
        if param is not None:
            return param
    except Exception:
        pass

    normalized_target = normalize_header(name_text)

    try:
        for param in element.Parameters:
            try:
                if param is None:
                    continue

                definition = getattr(param, "Definition", None)
                if definition is None:
                    continue

                def_name = normalize_text(getattr(definition, "Name", ""))
                if not def_name:
                    continue

                if def_name == name_text:
                    return param

                if normalize_header(def_name) == normalized_target:
                    return param
            except Exception:
                continue
    except Exception:
        pass

    return None


def find_parameter_from_metadata_name(element, metadata):
    if element is None or metadata is None:
        return None, None

    candidate_names = []

    try:
        candidate_names.append(metadata.get("Name", ""))
    except Exception:
        pass

    for candidate_name in candidate_names:
        param = find_parameter_by_name(element, candidate_name)
        if param is not None:
            return param, "Instance"

    type_element = get_type_element(element)
    for candidate_name in candidate_names:
        param = find_parameter_by_name(type_element, candidate_name)
        if param is not None:
            return param, "Type"

    return None, None


def get_parameter_from_metadata(element, metadata):
    if element is None or metadata is None:
        return None, None

    used_params = metadata.get("UsedParams", [])
    for used_param in used_params:
        parameter_id_value = None
        try:
            parameter_id_value = int(used_param)
        except Exception:
            continue

        param = find_parameter_on_element(element, parameter_id_value)
        if param is not None:
            return param, "Instance"

        type_element = get_type_element(element)
        param = find_parameter_on_element(type_element, parameter_id_value)
        if param is not None:
            return param, "Type"

    param, origin = find_parameter_from_metadata_name(element, metadata)
    if param is not None:
        return param, origin

    return None, None


def read_data_sheet_for_import_preview(workbook):
    worksheet = get_primary_data_worksheet(workbook)
    if worksheet is None:
        raise Exception("Missing worksheet with import data.")

    last_row, last_col = get_last_used_row_col(worksheet)
    if last_row < 3 or last_col < 2:
        raise Exception("The worksheet does not contain valid rows.")

    metadata_by_col = {}
    for c in range(1, last_col + 1):
        metadata_by_col[c] = parse_metadata_cell(get_cell_value(worksheet, 1, c))

    data_rows = []
    for r in range(3, last_row + 1):
        row_values = {}
        has_any_data = False

        for c in range(1, last_col + 1):
            text_value = normalize_text(get_cell_text(worksheet, r, c))
            row_values[c] = text_value
            if text_value:
                has_any_data = True

        if has_any_data:
            data_rows.append({
                "ExcelRow": r,
                "Values": row_values,
                "MetadataByCol": metadata_by_col
            })

    return {
        "Worksheet": worksheet,
        "LastRow": last_row,
        "LastCol": last_col,
        "MetadataByCol": metadata_by_col,
        "Rows": data_rows
    }


def get_primary_data_worksheet(workbook):
    worksheet = get_worksheet_by_name(workbook, "Data")
    if worksheet is not None:
        return worksheet

    try:
        worksheets = workbook.Worksheets
        total = worksheets.Count
    except Exception:
        return None

    for index in range(1, total + 1):
        try:
            candidate = worksheets[index]
            last_row, last_col = get_last_used_row_col(candidate)
            if last_row < 2 or last_col < 2:
                continue

            metadata1 = parse_metadata_cell(get_cell_value(candidate, 1, 1)) or {}
            metadata2 = parse_metadata_cell(get_cell_value(candidate, 1, 2)) or {}
            role1 = normalize_text(metadata1.get("ColumnRole", ""))
            role2 = normalize_text(metadata2.get("ColumnRole", ""))
            if role1 == "UniqueId" and role2 == "ElementId":
                return candidate
        except Exception:
            continue

    return None


def get_editable_columns_for_import_preview(data_info):
    results = []
    metadata_by_col = data_info.get("MetadataByCol", {})

    for c in sorted(metadata_by_col.keys()):
        md = metadata_by_col.get(c)
        if md is None:
            continue

        editable = normalize_text(md.get("Editable", ""))
        role = normalize_text(md.get("ColumnRole", ""))

        if role in ("UniqueId", "ElementId"):
            continue

        if editable == "Yes":
            results.append({
                "ExcelCol": c,
                "Metadata": md,
                "Name": normalize_text(md.get("Name", ""))
            })

    return results


def resolve_element_from_import_row(row_info):
    unique_id_text = get_import_row_reference_value(row_info, "UniqueId", 1)
    element_id_text = get_import_row_reference_value(row_info, "ElementId", 2)

    if unique_id_text:
        try:
            element = doc.GetElement(unique_id_text)
            if element is not None:
                return element, "UniqueId"
        except Exception:
            pass

    if element_id_text:
        try:
            eid = DB.ElementId(int(float(element_id_text)))
            element = doc.GetElement(eid)
            if element is not None:
                return element, "ElementId"
        except Exception:
            pass

    return None, "None"

UNIQUE_FIELD_RULES = [
    "View Name",
    "Sheet Number",
    "Level Name",
    "Grid Name"
]


def is_unique_controlled_field(field_name):
    return normalize_text(field_name) in [normalize_text(x) for x in UNIQUE_FIELD_RULES]


def value_exists_in_other_elements(field_name, new_value, current_element):
    field_name = normalize_text(field_name)
    new_value = normalize_text(new_value)

    if not new_value:
        return False, ""

    current_id = get_element_id_value(current_element)

    if field_name == normalize_text("View Name"):
        collector = DB.FilteredElementCollector(doc).OfClass(DB.View)
        for elem in collector:
            try:
                if elem is None:
                    continue
                if current_id != "" and get_element_id_value(elem) == current_id:
                    continue
                if normalize_text(elem.Name) == new_value:
                    return True, "Already used by View Id {}".format(get_element_id_value(elem))
            except Exception:
                pass

    elif field_name == normalize_text("Sheet Number"):
        collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)
        for elem in collector:
            try:
                if elem is None:
                    continue
                if current_id != "" and get_element_id_value(elem) == current_id:
                    continue

                param = elem.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER)
                if param is None:
                    continue

                existing_value = normalize_text(param.AsString())
                if existing_value == new_value:
                    return True, "Already used by Sheet Id {}".format(get_element_id_value(elem))
            except Exception:
                pass

    elif field_name == normalize_text("Level Name"):
        collector = DB.FilteredElementCollector(doc).OfClass(DB.Level)
        for elem in collector:
            try:
                if elem is None:
                    continue
                if current_id != "" and get_element_id_value(elem) == current_id:
                    continue
                if normalize_text(elem.Name) == new_value:
                    return True, "Already used by Level Id {}".format(get_element_id_value(elem))
            except Exception:
                pass

    elif field_name == normalize_text("Grid Name"):
        collector = DB.FilteredElementCollector(doc).OfClass(DB.Grid)
        for elem in collector:
            try:
                if elem is None:
                    continue
                if current_id != "" and get_element_id_value(elem) == current_id:
                    continue
                if normalize_text(elem.Name) == new_value:
                    return True, "Already used by Grid Id {}".format(get_element_id_value(elem))
            except Exception:
                pass

    return False, ""

def run_import_preview(xlsx_path):
    excel_app = None
    workbooks = None
    workbook = None

    try:
        excel_app = create_excel_application()
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

        workbooks = excel_app.Workbooks
        workbook = workbooks.Open(xlsx_path, False, True)

        data_info = read_data_sheet_for_import_preview(workbook)
        editable_columns = get_editable_columns_for_import_preview(data_info)
        rows = data_info.get("Rows", [])

        total_rows = len(rows)
        resolved_rows = 0
        unresolved_rows = 0
        changed_cells = 0
        rows_with_changes = 0
        duplicate_count = 0
        same_count = 0
        missing_param_count = 0

        preview_lines = []
        unresolved_lines = []
        duplicate_lines = []

        for row in rows:
            excel_row = row.get("ExcelRow", 0)
            element, resolved_by = resolve_element_from_import_row(row)

            if element is None:
                unresolved_rows += 1
                if len(unresolved_lines) < 10:
                    unresolved_lines.append(
                        "Row {} | UniqueId='{}' | ElementId='{}'".format(
                            excel_row,
                            safe_text(row.get("Values", {}).get(1, "")),
                            safe_text(row.get("Values", {}).get(2, ""))
                        )
                    )
                continue

            resolved_rows += 1
            row_change_count = 0

            for col_info in editable_columns:
                excel_col = col_info.get("ExcelCol", 0)
                metadata = col_info.get("Metadata")
                field_name = col_info.get("Name", "")
                excel_value = normalize_text(row.get("Values", {}).get(excel_col, ""))

                param, param_origin = get_parameter_from_metadata(element, metadata)
                if param is None:
                    missing_param_count += 1
                    continue

                current_value = get_parameter_preview_value(param)

                if normalize_text(current_value) == normalize_text(excel_value):
                    same_count += 1
                    continue

                if is_unique_controlled_field(field_name):
                    is_dup, dup_message = value_exists_in_other_elements(field_name, excel_value, element)
                    if is_dup:
                        duplicate_count += 1
                        if len(duplicate_lines) < 15:
                            duplicate_lines.append(
                                "Row {} | Id {} | {} | '{}' | {}".format(
                                    excel_row,
                                    safe_text(get_element_id_value(element)),
                                    field_name,
                                    excel_value,
                                    dup_message
                                )
                            )
                        continue

                row_change_count += 1
                changed_cells += 1

                if len(preview_lines) < 15:
                    preview_lines.append(
                        "Row {} | Id {} | {} | '{}' -> '{}'".format(
                            excel_row,
                            safe_text(get_element_id_value(element)),
                            field_name,
                            current_value,
                            excel_value
                        )
                    )

            if row_change_count > 0:
                rows_with_changes += 1

        message = []
        message.append("Import preview completed.")
        message.append("")
        message.append("File: {}".format(xlsx_path))
        message.append("Rows: {}".format(total_rows))
        message.append("Resolved rows: {}".format(resolved_rows))
        message.append("Unresolved rows: {}".format(unresolved_rows))
        message.append("Rows with changes: {}".format(rows_with_changes))
        message.append("Changed cells: {}".format(changed_cells))
        message.append("Duplicate conflicts: {}".format(duplicate_count))
        message.append("Same values: {}".format(same_count))
        message.append("Missing parameters: {}".format(missing_param_count))

        if preview_lines:
            message.append("")
            message.append("Sample valid changes:")
            for line in preview_lines:
                message.append(line)

        if duplicate_lines:
            message.append("")
            message.append("Sample duplicate conflicts:")
            for line in duplicate_lines:
                message.append(line)

        if unresolved_lines:
            message.append("")
            message.append("Sample unresolved rows:")
            for line in unresolved_lines:
                message.append(line)

        TaskDialog.Show(__title__, "\n".join(message))

    finally:
        try:
            if workbook is not None:
                workbook.Close(False)
        except Exception:
            pass

        try:
            if excel_app is not None:
                excel_app.Quit()
        except Exception:
            pass

        release_com_object(workbook)
        release_com_object(workbooks)
        release_com_object(excel_app)

def get_schedule_field_export_value(element, field_info):
    if element is None or field_info is None:
        return ""

    metadata = field_info.get("metadata") or {}
    param, origin = get_parameter_from_metadata(element, metadata)
    if param is None:
        return ""

    return get_parameter_preview_value(param)



def build_export_rows_from_schedule_elements(schedule, visible_fields):
    rows = []
    elements = get_schedule_elements(schedule)

    for element in elements:
        if element is None:
            continue

        row = []

        try:
            row.append(safe_text(element.UniqueId))
        except Exception:
            row.append("")

        row.append(safe_text(get_element_id_value(element)))

        for field_info in visible_fields:
            row.append(get_schedule_field_export_value(element, field_info))

        rows.append(row)

    return rows, len(elements)

def row_has_schedule_payload(row_values):
    if not row_values:
        return False

    for value in row_values[2:]:
        if normalize_text(value):
            return True

    return False


def get_export_value_from_metadata(element, metadata):
    if element is None or metadata is None:
        return ""

    param, _ = get_parameter_from_metadata(element, metadata)
    parameter_name = normalize_text(metadata.get("Name", ""))
    parameter_key = parameter_name.lower()

    if param is None:
        if parameter_key == "family name":
            return get_family_name_from_element(element)
        if parameter_key == "type name":
            return get_type_name_from_element(element)
        return ""

    value = get_parameter_preview_value(param)

    if parameter_key in ("area", "volume") and is_annotation_element(element) and is_zero_number_text(value):
        return ""

    if not normalize_text(value):
        if parameter_key == "family name":
            return get_family_name_from_element(element)
        if parameter_key == "type name":
            return get_type_name_from_element(element)

    if parameter_key in ("design option", "host id", "level"):
        value_text = normalize_text(value)
        if value_text == "-1":
            return "None"
        if value_text.lower() == "none":
            return "None"

    return value


def build_export_rows_from_elements(elements, parameter_items, include_type_id=False):
    rows = []
    elements = list(elements or [])

    for element in elements:
        if element is None:
            continue

        row = []

        try:
            row.append(safe_text(element.UniqueId))
        except Exception:
            row.append("")

        row.append(safe_text(get_element_id_value(element)))

        if include_type_id:
            row.append(safe_text(get_element_type_id_value(element)))

        for param_item in parameter_items:
            metadata = getattr(param_item, "Metadata", {}) or {}
            row.append(get_export_value_from_metadata(element, metadata))

        rows.append(row)

    return rows, len(elements)


def build_export_rows_from_category_elements(category, parameter_items, include_type_id=False):
    return build_export_rows_from_elements(get_category_elements(category), parameter_items, include_type_id=include_type_id)


def export_category_to_xlsx(category_item, parameter_items, full_path, source_elements=None, sheet_suffix="ModelCategory", keep_empty_rows=False, include_type_id=False):
    if category_item is None:
        raise Exception("No model category selected.")

    if not parameter_items:
        raise Exception("No parameters were found in the selected category.")

    final_headers = ["GUID", "Element ID"]
    export_columns = [
        {
            "ExcelIndex": 1,
            "Name": "UniqueId",
            "Status": "Technical",
            "Origin": "Technical",
            "Editable": "Hidden",
            "ScheduleId": "",
            "UsedParams": [],
            "FieldIndex": "",
            "ColumnRole": "UniqueId",
            "Hidden": True
        },
        {
            "ExcelIndex": 2,
            "Name": "ElementId",
            "Status": "Technical",
            "Origin": "Technical",
            "Editable": "Visible",
            "ScheduleId": "",
            "UsedParams": [],
            "FieldIndex": "",
            "ColumnRole": "ElementId",
            "Hidden": False
        }
    ]

    if include_type_id:
        final_headers.append("Type ID")
        export_columns.append({
            "ExcelIndex": len(export_columns) + 1,
            "Name": "TypeId",
            "Status": "Technical",
            "Origin": "Technical",
            "Editable": "Visible",
            "ScheduleId": "",
            "UsedParams": [],
            "FieldIndex": "",
            "ColumnRole": "TypeId",
            "Hidden": False
        })

    for param_item in parameter_items:
        metadata = dict(getattr(param_item, "Metadata", {}) or {})

        final_headers.append(param_item.Name)
        export_columns.append({
            "ExcelIndex": len(export_columns) + 1,
            "Name": param_item.Name,
            "Status": metadata.get("Status", param_item.Status),
            "Origin": metadata.get("Origin", param_item.Origin),
            "Editable": metadata.get("Editable", param_item.Editable),
            "ScheduleId": "",
            "UsedParams": metadata.get("UsedParams", []),
            "FieldIndex": "",
            "ColumnRole": "CategoryField",
            "Hidden": False
        })

    if source_elements is None:
        all_rows, category_element_count = build_export_rows_from_category_elements(category_item.Category, parameter_items, include_type_id=include_type_id)
    else:
        all_rows, category_element_count = build_export_rows_from_elements(source_elements, parameter_items, include_type_id=include_type_id)

    if keep_empty_rows:
        final_rows = list(all_rows)
    else:
        final_rows = [row for row in all_rows if row_has_schedule_payload(row)]
    metadata_row = [json.dumps(column_info, separators=(",", ":")) for column_info in export_columns]

    excel_app = None
    workbooks = None
    workbook = None
    data_sheet = None
    instructions_sheet = None
    paramvalues_sheet = None

    try:
        excel_app = create_excel_application()
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        workbooks = excel_app.Workbooks
        workbook = workbooks.Add()
        ensure_workbook_sheet_count(workbook, 3)

        data_sheet = workbook.Worksheets[1]
        instructions_sheet = workbook.Worksheets[2]
        paramvalues_sheet = workbook.Worksheets[3]

        data_sheet.Name = make_excel_sheet_name(category_item.Name, sheet_suffix)
        instructions_sheet.Name = "Instructions"
        paramvalues_sheet.Name = "ParamValues"

        write_matrix_to_range(data_sheet, 1, 1, [metadata_row, final_headers])

        metadata_range = data_sheet.Range[data_sheet.Cells[1, 1], data_sheet.Cells[1, len(final_headers)]]
        metadata_range.Font.Color = 0xFFFFFF
        metadata_range.Interior.Color = 0x404040
        metadata_range.RowHeight = 18
        data_sheet.Rows[1].Hidden = True

        header_range = data_sheet.Range[data_sheet.Cells[2, 1], data_sheet.Cells[2, len(final_headers)]]
        header_range.Font.Bold = True
        header_range.Font.Color = 0xFFFFFF
        header_range.Interior.Color = 0xC9951A
        header_range.VerticalAlignment = -4108
        header_range.WrapText = True
        header_range.RowHeight = 36

        if final_rows:
            write_matrix_to_range(data_sheet, 3, 1, final_rows)

        last_data_row = max(2, len(final_rows) + 2)
        used_range = data_sheet.Range[data_sheet.Cells[2, 1], data_sheet.Cells[last_data_row, len(final_headers)]]
        used_range.Borders.LineStyle = 1
        used_range.VerticalAlignment = -4108

        try:
            used_range.AutoFilter()
        except Exception:
            pass

        total_cols = len(final_headers)
        for c in range(1, total_cols + 1):
            role = export_columns[c - 1].get("ColumnRole", "")
            hidden = export_columns[c - 1].get("Hidden", False)
            editable = export_columns[c - 1].get("Editable", "Unknown")
            origin = export_columns[c - 1].get("Origin", "Unknown")

            if hidden:
                data_sheet.Columns[c].Hidden = True

            header_cell = data_sheet.Cells[2, c]
            header_cell.Font.Bold = True
            header_cell.Locked = False

            if role in ("ElementId", "TypeId"):
                header_cell.Font.Color = 0xFFFFFF
                header_cell.Interior.Color = 0x1450BE
            elif editable == "No":
                header_cell.Font.Color = 0x000000
                header_cell.Interior.Color = 0xD3D9F7
            elif origin == "Type":
                header_cell.Font.Color = 0x000000
                header_cell.Interior.Color = 0xC9F1F9
            else:
                header_cell.Font.Color = 0x000000
                header_cell.Interior.Color = 0xD5E2FB

            if last_data_row < 3:
                continue

            data_range = data_sheet.Range[
                data_sheet.Cells[3, c],
                data_sheet.Cells[last_data_row, c]
            ]
            data_range.Font.Color = 0x000000

            if role == "UniqueId":
                data_range.Interior.Color = 0xF2F2F2
            elif role in ("ElementId", "TypeId"):
                data_range.Interior.Pattern = -4142
            elif editable == "No":
                data_range.Interior.Color = 0xF7D9D3
            elif origin == "Type":
                data_range.Interior.Color = 0xF9F1C9
            else:
                data_range.Interior.Color = 0xEEFAD7

        fit_export_columns(data_sheet, export_columns, 8, 40, last_data_row)

        instruction_rows = [
            [None, None, None],
            [None, "Cell Fill Colour", "Description"],
            [None, None, "Type value"],
            [None, None, "Read-only value"],
            [None, None, "Parameter does not exist for this element"],
            [None, None, None],
            [None, "Note:", None],
            [None, "If you are altering the value of 'Type Parameters', ensure that you have the same value for all elements with the same 'Type ID'", None],
            [None, None, None]
        ]

        instruction_fill_colors = {
            3: 0xC9F1F9,
            4: 0xD3D9F7,
            5: 0xF2F2F2
        }

        write_matrix_to_range(instructions_sheet, 1, 1, instruction_rows)

        instructions_sheet.Columns[1].ColumnWidth = 3
        instructions_sheet.Columns[2].ColumnWidth = 20
        instructions_sheet.Columns[3].ColumnWidth = 65
        instructions_sheet.Rows[8].RowHeight = 52

        for row_index, fill_color in instruction_fill_colors.items():
            try:
                instructions_sheet.Cells[row_index, 2].Interior.Color = fill_color
                instructions_sheet.Cells[row_index, 2].Borders.LineStyle = 1
            except Exception:
                pass

        try:
            instructions_sheet.Cells[2, 2].Font.Bold = True
            instructions_sheet.Cells[2, 3].Font.Bold = True
            instructions_sheet.Cells[7, 2].Font.Bold = True
            instructions_sheet.Range[instructions_sheet.Cells[8, 2], instructions_sheet.Cells[8, 3]].Merge()
            instructions_sheet.Range[instructions_sheet.Cells[2, 2], instructions_sheet.Cells[8, 3]].WrapText = True
            instructions_sheet.Range[instructions_sheet.Cells[2, 2], instructions_sheet.Cells[8, 3]].VerticalAlignment = -4108
        except Exception:
            pass

        param_value_rows = [
            ["Yes", "Undefined", "Architectural", "Grids and Levels", "Wireframe"],
            ["No", "Coarse", "Structural", "Views Overall", "Hidden Line"],
            ["", "Medium", "Mechanical", "None", None],
            [None, "Fine", "Electrical", None, None],
            [None, None, "Plumbing", None, None],
            [None, None, "Coordination", None, None]
        ]

        write_matrix_to_range(paramvalues_sheet, 1, 1, param_value_rows)

        for col_index in range(1, 6):
            try:
                paramvalues_sheet.Columns[col_index].AutoFit()
                if paramvalues_sheet.Columns[col_index].ColumnWidth < 12:
                    paramvalues_sheet.Columns[col_index].ColumnWidth = 12
            except Exception:
                pass

        try:
            paramvalues_sheet.Visible = 0
        except Exception:
            pass

        try:
            while workbook.Worksheets.Count > 3:
                workbook.Worksheets[workbook.Worksheets.Count].Delete()
        except Exception:
            pass

        try:
            data_sheet.Activate()
        except Exception:
            pass

        workbook.SaveAs(full_path)
        workbook.Close(True)
        excel_app.Quit()

        return full_path, len(final_rows), category_element_count

    finally:
        release_com_object(paramvalues_sheet)
        release_com_object(instructions_sheet)
        release_com_object(data_sheet)
        release_com_object(workbook)
        release_com_object(workbooks)
        release_com_object(excel_app)


def export_schedule_to_xlsx(schedule, full_path):
    visible_fields = get_visible_schedule_fields(schedule)
    if not visible_fields:
        raise Exception("No visible fields were found in the selected schedule.")

    final_headers = ["UniqueId", "ElementId"]
    export_columns = [
        {
            "ExcelIndex": 1,
            "Name": "UniqueId",
            "Status": "Technical",
            "Origin": "Technical",
            "Editable": "Hidden",
            "ScheduleId": "",
            "UsedParams": [],
            "FieldIndex": "",
            "ColumnRole": "UniqueId",
            "Hidden": True
        },
        {
            "ExcelIndex": 2,
            "Name": "ElementId",
            "Status": "Technical",
            "Origin": "Technical",
            "Editable": "Visible",
            "ScheduleId": "",
            "UsedParams": [],
            "FieldIndex": "",
            "ColumnRole": "ElementId",
            "Hidden": False
        }
    ]

    for field_info in visible_fields:
        md = dict(field_info.get("metadata", {}))
        final_headers.append(field_info.get("name", ""))
        export_columns.append({
            "ExcelIndex": len(export_columns) + 1,
            "Name": field_info.get("name", ""),
            "Status": md.get("Status", "Unknown"),
            "Origin": md.get("Origin", "Special"),
            "Editable": md.get("Editable", "Unknown"),
            "ScheduleId": md.get("ScheduleId", ""),
            "UsedParams": md.get("UsedParams", []),
            "FieldIndex": md.get("FieldIndex", ""),
            "ColumnRole": "ScheduleField",
            "Hidden": False
        })

    all_rows, schedule_element_count = build_export_rows_from_schedule_elements(
        schedule,
        visible_fields
    )

    final_rows = [row for row in all_rows if row_has_schedule_payload(row)]

    has_valid_element_ids = False
    for row in final_rows:
        if len(row) > 1 and normalize_text(row[1]):
            has_valid_element_ids = True
            break

    metadata_row = []
    for col in export_columns:
        metadata_row.append(json.dumps({
            "Name": col.get("Name", ""),
            "Status": col.get("Status", ""),
            "Origin": col.get("Origin", ""),
            "Editable": col.get("Editable", ""),
            "ScheduleId": col.get("ScheduleId", ""),
            "UsedParams": col.get("UsedParams", []),
            "FieldIndex": col.get("FieldIndex", ""),
            "ColumnRole": col.get("ColumnRole", ""),
            "Hidden": col.get("Hidden", False)
        }, separators=(",", ":")))

    excel_app = None
    workbooks = None
    workbook = None
    data_sheet = None
    instructions_sheet = None
    schema_sheet = None

    try:
        excel_app = create_excel_application()
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

        workbooks = excel_app.Workbooks
        workbook = workbooks.Add()
        ensure_workbook_sheet_count(workbook, 3)

        data_sheet = workbook.Worksheets[1]
        instructions_sheet = workbook.Worksheets[2]
        schema_sheet = workbook.Worksheets[3]

        data_sheet.Name = "Data"
        instructions_sheet.Name = "Instructions"
        schema_sheet.Name = "Schema"

        total_cols = len(final_headers)

        write_matrix_to_range(data_sheet, 1, 1, [metadata_row, final_headers])
        if final_rows:
            write_matrix_to_range(data_sheet, 3, 1, final_rows)

        meta_row_range = data_sheet.Range[data_sheet.Cells[1, 1], data_sheet.Cells[1, total_cols]]
        meta_row_range.Font.Color = 0x808080
        meta_row_range.Interior.Color = 0xF2F2F2
        data_sheet.Rows[1].Hidden = True

        last_data_row = max(2, len(final_rows) + 2)

        for c in range(1, total_cols + 1):
            role = export_columns[c - 1].get("ColumnRole", "")
            hidden = export_columns[c - 1].get("Hidden", False)
            editable = export_columns[c - 1].get("Editable", "Unknown")
            status_text = export_columns[c - 1].get("Status", "Unknown")

            if hidden:
                data_sheet.Columns[c].Hidden = True

            header_cell = data_sheet.Cells[2, c]
            header_cell.Font.Bold = True
            header_cell.Locked = False

            if role == "ElementId":
                header_cell.Font.Color = 0x000000
                header_cell.Interior.Color = 0x1450BE
            elif editable == "Yes":
                header_cell.Font.Color = 0x000000
                header_cell.Interior.Color = 0xD5E2FB
            else:
                header_cell.Font.Color = 0x000000
                header_cell.Interior.Color = 0xACC7F7

            if last_data_row < 3:
                continue

            data_range = data_sheet.Range[
                data_sheet.Cells[3, c],
                data_sheet.Cells[last_data_row, c]
            ]

            if role in ("UniqueId", "ElementId"):
                data_range.Locked = True
            else:
                data_range.Locked = editable == "No"

            data_range.Font.Color = 0x000000

            if role == "ElementId":
                data_range.Interior.Pattern = -4142
            elif role == "UniqueId":
                data_range.Interior.Color = 0xF2F2F2
            elif editable == "Yes":
                data_range.Interior.Pattern = -4142
            else:
                fill_color = get_status_fill_color(status_text)
                if fill_color is not None:
                    data_range.Interior.Color = fill_color
                else:
                    data_range.Interior.Pattern = -4142

        used_range = data_sheet.Range[data_sheet.Cells[2, 1], data_sheet.Cells[last_data_row, total_cols]]
        used_range.Borders.LineStyle = 1
        used_range.AutoFilter()
        fit_export_columns(data_sheet, export_columns, 8, 40, last_data_row)

        try:
            data_sheet.Columns[4].ColumnWidth = 24
        except Exception:
            pass

        # Leave the worksheet unprotected so Excel sorting, filtering, and free-form edits work normally.

        instruction_rows = [
            [None, None, None],
            [None, "Cell Fill Colour", "Description"],
            [None, None, "Type value"],
            [None, None, "Read-only value"],
            [None, None, "Parameter does not exist for this element"],
            [None, None, None],
            [None, "Note:", None],
            [None, "Edit values on the first worksheet only. The hidden metadata row in Data is used for import, so keep row 1 untouched.", None],
            [None, None, None]
        ]

        instruction_fill_colors = {
            3: 0xC9F1F9,
            4: 0xD3D9F7,
            5: 0xF2F2F2
        }

        write_matrix_to_range(instructions_sheet, 1, 1, instruction_rows)

        instructions_sheet.Columns[1].ColumnWidth = 3
        instructions_sheet.Columns[2].ColumnWidth = 20
        instructions_sheet.Columns[3].ColumnWidth = 65
        instructions_sheet.Rows[8].RowHeight = 52

        for row_index, fill_color in instruction_fill_colors.items():
            try:
                instructions_sheet.Cells[row_index, 2].Interior.Color = fill_color
                instructions_sheet.Cells[row_index, 2].Borders.LineStyle = 1
            except Exception:
                pass

        try:
            instructions_sheet.Cells[2, 2].Font.Bold = True
            instructions_sheet.Cells[2, 3].Font.Bold = True
            instructions_sheet.Cells[7, 2].Font.Bold = True
            instructions_sheet.Range[instructions_sheet.Cells[8, 2], instructions_sheet.Cells[8, 3]].Merge()
            instructions_sheet.Range[instructions_sheet.Cells[2, 2], instructions_sheet.Cells[8, 3]].WrapText = True
            instructions_sheet.Range[instructions_sheet.Cells[2, 2], instructions_sheet.Cells[8, 3]].VerticalAlignment = -4108
        except Exception:
            pass

        schema_headers = [
            "Schema Index", "Excel Column", "Name", "Status", "Origin",
            "Editable", "ScheduleId", "UsedParams", "FieldIndex", "ColumnRole", "Hidden"
        ]
        for c in range(1, len(schema_headers) + 1):
            schema_sheet.Cells[1, c].Value2 = schema_headers[c - 1]

        schema_header_range = schema_sheet.Range[schema_sheet.Cells[1, 1], schema_sheet.Cells[1, len(schema_headers)]]
        schema_header_range.Font.Bold = True
        schema_header_range.Font.Color = 0xFFFFFF
        schema_header_range.Interior.Color = 0x5B7D95

        schema_rows = build_schema_rows(export_columns)
        row_index = 2
        for row_vals in schema_rows:
            for c in range(1, len(row_vals) + 1):
                schema_sheet.Cells[row_index, c].Value2 = row_vals[c - 1]
            row_index += 1

        schema_used = schema_sheet.Range[schema_sheet.Cells[1, 1], schema_sheet.Cells[max(1, row_index - 1), len(schema_headers)]]
        schema_used.Borders.LineStyle = 1
        schema_last_row = row_index - 1

        for r in range(2, schema_last_row + 1):
            try:
                name_value = safe_text(schema_sheet.Cells[r, 3].Value2).strip()
                editable_value = safe_text(schema_sheet.Cells[r, 6].Value2).strip().lower()

                row_range = schema_sheet.Range[
                    schema_sheet.Cells[r, 1],
                    schema_sheet.Cells[r, len(schema_headers)]
                ]

                row_range.Font.Color = 0x000000

                if name_value in ("UniqueId", "ElementId"):
                    row_range.Interior.Color = 0x1450BE
                elif editable_value == "yes":
                    row_range.Interior.Pattern = -4142
                elif editable_value == "no":
                    row_range.Interior.Color = 0xACC7F7
                else:
                    row_range.Interior.Color = 0xF2F2F2
            except Exception:
                pass

        for c in range(1, len(schema_headers) + 1):
            try:
                schema_sheet.Columns[c].AutoFit()
                if schema_sheet.Columns[c].ColumnWidth > 28:
                    schema_sheet.Columns[c].ColumnWidth = 28
            except Exception:
                pass

        try:
            schema_sheet.Visible = 0
        except Exception:
            pass

        try:
            data_sheet.Activate()
        except Exception:
            pass

        workbook.SaveAs(full_path)
        workbook.Close(True)
        excel_app.Quit()

        return full_path, has_valid_element_ids, "ScheduleElements", len(final_rows), schedule_element_count, 0

    finally:
        release_com_object(schema_sheet)
        release_com_object(instructions_sheet)
        release_com_object(data_sheet)
        release_com_object(workbook)
        release_com_object(workbooks)
        release_com_object(excel_app)
class ScheduleBrowserWindow(forms.WPFWindow):
    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)

        self.active_view = "Model Categories"
        self.preview_source_view = "Model Categories"
        self.model_scope = "Whole Model"
        self.annotation_scope = "Whole Model"
        self.elements_scope = "Whole Model"
        self.model_export_type_id = False
        self.annotation_export_type_id = False
        self.elements_export_type_id = False
        self._suspend_selection_changed = False
        self._set_progress(5, "Starting")
        self.all_schedules = collect_schedules()
        self._set_progress(30, "Schedules loaded")
        self.filtered_schedules = list(self.all_schedules)
        self.current_parameters = []
        self.filtered_parameters = []
        self.all_model_categories = collect_model_categories(self.model_scope)
        self._set_progress(55, "Categories loaded")
        self.filtered_model_categories = list(self.all_model_categories)
        self.current_model_parameters = []
        self.filtered_model_parameters = []
        self.selected_model_parameters = []
        self.filtered_selected_model_parameters = []
        self.annotation_categories_loaded = False
        self.element_categories_loaded = False
        self.spatial_categories_loaded = False
        self.all_annotation_categories = []
        self.filtered_annotation_categories = []
        self.current_annotation_parameters = []
        self.filtered_annotation_parameters = []
        self.selected_annotation_parameters = []
        self.filtered_selected_annotation_parameters = []
        self.all_element_categories = []
        self.filtered_element_categories = []
        self.current_element_items = []
        self.filtered_element_items = []
        self.current_element_parameters = []
        self.filtered_element_parameters = []
        self.selected_element_parameters = []
        self.filtered_selected_element_parameters = []
        self.all_spatial_categories = []
        self.filtered_spatial_categories = []
        self.current_spatial_items = []
        self.filtered_spatial_items = []
        self.current_spatial_parameters = []
        self.filtered_spatial_parameters = []
        self.selected_spatial_parameters = []
        self.filtered_selected_spatial_parameters = []
        self.preview_table = None
        self.preview_row_elements = []
        self.preview_column_map = {}
        self.preview_original_values = {}
        self.preview_pending_change_count = 0
        self.preview_import_mode = False
        self.preview_import_table = None
        self.preview_import_source_path = ""
        self.preview_import_row_count = 0
        self.preview_import_editable_count = 0
        self.preview_changed_marker_map = {}
        self._suspend_preview_autosave = False
        self._is_committing_preview_edits = False

        self._configure_branding()
        self._setup_navigation()
        self._bind_events()
        self._configure_export_state()
        self._set_progress(75, "Preparing interface")
        self._refresh_schedule_list()
        self._refresh_parameter_grid([])
        self._refresh_model_category_list()
        self._refresh_model_parameter_grid([])
        self._refresh_model_selected_parameter_list([])
        self._clear_preview_grid()
        self._set_active_view("Model Categories")
        self._update_status()
        self._set_progress(100, "Ready")

    def _refresh_progress_ui(self):
        try:
            self.UpdateLayout()
        except Exception:
            pass

        try:
            Application.DoEvents()
        except Exception:
            pass

    def _set_progress(self, value, message=None):
        try:
            progress_value = int(value)
        except Exception:
            progress_value = 0

        if progress_value < 0:
            progress_value = 0
        if progress_value > 100:
            progress_value = 100

        try:
            self.prgHeaderProgress.Value = progress_value
        except Exception:
            pass

        label = safe_text(message).strip()
        if not label:
            label = "Completed"

        try:
            self.lblHeaderProgress.Text = "{}   {}%".format(label, progress_value)
        except Exception:
            pass

        self._refresh_progress_ui()

    def _set_completed_export_summary(self, label, rows_written, fields_exported):
        try:
            summary_text = "{} | Rows: {} | Fields: {}".format(
                safe_text(label),
                safe_text(rows_written),
                safe_text(fields_exported)
            )
            self.lblHeaderProgress.Text = summary_text
            self.prgHeaderProgress.Value = 100
        except Exception:
            pass

        self._refresh_progress_ui()

    def _configure_branding(self):
        logo_path = os.path.join(THIS_DIR, "logo.png")
        if not os.path.exists(logo_path):
            return

        try:
            self.set_image_source(self.imgLogo, logo_path)
            if hasattr(self, "imgHeaderMark"):
                self.set_image_source(self.imgHeaderMark, logo_path)
        except Exception:
            logger.debug("Could not load logo image from %s", logo_path)

    def _setup_navigation(self):
        self.nav_buttons = {
            "Model Categories": self.btnNavModelCategories,
            "Annotation Categories": self.btnNavAnnotationCategories,
            "Elements": self.btnNavElements,
            "Schedules": self.btnNavSchedules,
            "Spatial": self.btnNavSpatial,
            "Preview/Edit": self.btnNavPreviewEdit,
        }

        self.view_panels = {
            "Model Categories": self.panelModelCategories,
            "Annotation Categories": self.panelAnnotationCategories,
            "Elements": self.panelElements,
            "Schedules": self.panelSchedules,
            "Spatial": self.panelElements,
            "Preview/Edit": self.panelPreviewEdit,
        }

    def _bind_events(self):
        for nav_button in self.nav_buttons.values():
            nav_button.Click += self.on_nav_clicked

        self.txtSearchSchedules.TextChanged += self.on_schedule_search_changed
        self.txtSearchParameters.TextChanged += self.on_parameter_search_changed
        self.txtSearchModelCategories.TextChanged += self.on_model_category_search_changed
        self.txtSearchModelParameters.TextChanged += self.on_model_parameter_search_changed
        self.txtSearchModelSelectedParameters.TextChanged += self.on_model_selected_parameter_search_changed
        self.txtSearchAnnotationCategories.TextChanged += self.on_annotation_category_search_changed
        self.txtSearchAnnotationParameters.TextChanged += self.on_annotation_parameter_search_changed
        self.txtSearchAnnotationSelectedParameters.TextChanged += self.on_annotation_selected_parameter_search_changed
        self.txtSearchElementCategories.TextChanged += self.on_element_category_search_changed
        self.txtSearchElementItems.TextChanged += self.on_element_item_search_changed
        self.txtSearchElementParameters.TextChanged += self.on_element_parameter_search_changed
        self.txtSearchElementSelectedParameters.TextChanged += self.on_element_selected_parameter_search_changed
        self.cmbModelScope.SelectionChanged += self.on_model_scope_changed
        self.cmbAnnotationScope.SelectionChanged += self.on_annotation_scope_changed
        self.cmbElementsScope.SelectionChanged += self.on_elements_scope_changed
        self.chkModelExportTypeId.Checked += self.on_model_export_type_id_changed
        self.chkModelExportTypeId.Unchecked += self.on_model_export_type_id_changed
        self.chkAnnotationExportTypeId.Checked += self.on_annotation_export_type_id_changed
        self.chkAnnotationExportTypeId.Unchecked += self.on_annotation_export_type_id_changed
        self.chkElementsExportTypeId.Checked += self.on_elements_export_type_id_changed
        self.chkElementsExportTypeId.Unchecked += self.on_elements_export_type_id_changed
        self.lstSchedules.SelectionChanged += self.on_schedule_selection_changed
        self.lstModelCategories.SelectionChanged += self.on_model_category_selection_changed
        self.lstAnnotationCategories.SelectionChanged += self.on_annotation_category_selection_changed
        self.lstElementCategories.SelectionChanged += self.on_element_category_selection_changed
        self.lstElementItems.SelectionChanged += self.on_element_item_selection_changed
        self.btnModelAddParameter.Click += self.on_model_add_parameter_clicked
        self.btnModelRemoveParameter.Click += self.on_model_remove_parameter_clicked
        self.btnModelSelectedMoveUp.Click += self.on_model_selected_move_up_clicked
        self.btnModelSelectedMoveDown.Click += self.on_model_selected_move_down_clicked
        self.btnAnnotationAddParameter.Click += self.on_annotation_add_parameter_clicked
        self.btnAnnotationRemoveParameter.Click += self.on_annotation_remove_parameter_clicked
        self.btnAnnotationSelectedMoveUp.Click += self.on_annotation_selected_move_up_clicked
        self.btnAnnotationSelectedMoveDown.Click += self.on_annotation_selected_move_down_clicked
        self.btnElementAddParameter.Click += self.on_element_add_parameter_clicked
        self.btnElementRemoveParameter.Click += self.on_element_remove_parameter_clicked
        self.btnElementSelectedMoveUp.Click += self.on_element_selected_move_up_clicked
        self.btnElementSelectedMoveDown.Click += self.on_element_selected_move_down_clicked
        self.btnImportExcel.Click += self.on_import_excel_clicked
        self.miImportFromExcel.Click += self.on_import_from_excel_clicked
        self.miPreviewImport.Click += self.on_preview_import_clicked
        self.btnExport.Click += self.on_export_clicked
        self.btnResetValues.Click += self.on_reset_values_clicked
        self.btnOpenPreviewEdit.Click += self.on_open_preview_edit_clicked
        self.btnExportProjectStandards.Click += self.on_export_project_standards_clicked
        self.btnPreviewImport.Click += self.on_preview_update_clicked
        self.btnPreviewExport.Click += self.on_export_clicked
        self.btnGoSchedules.Click += self.on_go_schedules_clicked
        self.dgPreviewEdit.AutoGeneratingColumn += self.on_preview_auto_generating_column
        self.dgPreviewEdit.CurrentCellChanged += self.on_preview_current_cell_changed

    def _configure_export_state(self):
        self.btnExport.IsEnabled = True
        self.btnExport.ToolTip = "Export the current selection to Excel"
        self.btnImportExcel.ToolTip = "Preview or import Excel changes back into Revit"
        self.btnResetValues.ToolTip = "Reset the current workspace state"
        self.btnOpenPreviewEdit.ToolTip = "Open Preview/Edit for the current selection"
        self.btnExportProjectStandards.ToolTip = "Export project standards workflow"
        self.btnPreviewImport.ToolTip = "Apply editable Preview/Edit cell changes to Revit"
        self.btnPreviewExport.ToolTip = "Export the selected schedule to Excel"
        self.btnGoSchedules.Content = "Back"
        self.btnGoSchedules.ToolTip = "Go back to the source workspace"

    def _ensure_annotation_categories_loaded(self):
        if self.annotation_categories_loaded:
            return

        self._set_progress(20, "Loading annotation categories")
        self.all_annotation_categories = collect_annotation_categories(self.annotation_scope)
        self.filtered_annotation_categories = list(self.all_annotation_categories)
        self.annotation_categories_loaded = True
        self._set_progress(100, "Ready")

    def _ensure_element_categories_loaded(self):
        if self.element_categories_loaded:
            return

        self._ensure_annotation_categories_loaded()
        self._set_progress(25, "Loading element categories")
        self.all_element_categories = collect_element_categories(self.elements_scope)
        self.filtered_element_categories = list(self.all_element_categories)
        self.element_categories_loaded = True
        self._set_progress(100, "Ready")

    def _ensure_spatial_categories_loaded(self):
        if self.spatial_categories_loaded:
            return

        self._set_progress(25, "Loading rooms/spaces")
        self.all_spatial_categories = collect_spatial_categories()
        self.filtered_spatial_categories = list(self.all_spatial_categories)
        self.spatial_categories_loaded = True
        self._set_progress(100, "Ready")

    def _ensure_view_data_loaded(self, view_name):
        if view_name == "Annotation Categories":
            self._ensure_annotation_categories_loaded()
        elif view_name == "Elements":
            self._ensure_element_categories_loaded()
        elif view_name == "Spatial":
            self._ensure_spatial_categories_loaded()

    def _set_active_view(self, view_name):
        if view_name not in self.view_panels:
            return

        if view_name != "Preview/Edit":
            self.preview_source_view = view_name

        self.active_view = view_name
        self._ensure_view_data_loaded(view_name)

        active_panel = self.view_panels[view_name]
        handled_panels = []
        for panel_name, panel in self.view_panels.items():
            if panel in handled_panels:
                continue
            handled_panels.append(panel)
            panel.Visibility = Visibility.Visible if panel is active_panel else Visibility.Collapsed

        for button_name, button in self.nav_buttons.items():
            is_active = button_name == view_name
            button.Background = self.Resources["TabActiveBg"] if is_active else self.Resources["TabBg"]
            button.FontWeight = FontWeights.SemiBold if is_active else FontWeights.Normal

        intro_text = {
            "Model Categories": "Select one model category, review its parameters, then export or import Excel changes.",
            "Annotation Categories": "Select one annotation category, review its parameters, then export or import Excel changes.",
            "Elements": "Work at element level when you want to target instances or types directly instead of going through schedules.",
            "Schedules": "Select one schedule, review its parameters, then export or import Excel changes.",
            "Spatial": "Use the spatial branch for rooms and spaces with their own Excel-first workflow.",
            "Preview/Edit": "Review the current selection and launch export or import actions from a lightweight control room.",
        }

        self.lblIntroText.Text = intro_text.get(view_name, "")
        if view_name == "Schedules":
            self._refresh_schedule_list()
        elif view_name == "Model Categories":
            self._refresh_model_category_list()
            self._sync_model_parameter_views()
        elif view_name == "Annotation Categories":
            self._refresh_annotation_category_list()
            self._sync_annotation_parameter_views()
        elif view_name == "Elements":
            self._refresh_element_category_list()
            self._refresh_element_list()
            self._sync_element_parameter_views()
        elif view_name == "Spatial":
            self._refresh_spatial_category_list()
            self._refresh_spatial_list()
            self._sync_spatial_parameter_views()

        self._update_preview_panel()
        self._update_context_panels()
        self._update_action_buttons()
        self._update_status()

    def _update_action_buttons(self):
        selected_item = self._get_selected_item()
        selected_model_category = self._get_selected_model_category_item()
        selected_annotation_category = self._get_selected_annotation_category_item()
        selected_element_category = self._get_selected_element_category_item()
        selected_elements = self._get_selected_element_items()
        selected_spatial_category = self._get_selected_spatial_category_item()
        selected_spatial_items = self._get_selected_spatial_items()
        can_use_excel_actions = self.active_view in ("Schedules", "Preview/Edit", "Model Categories", "Annotation Categories", "Elements", "Spatial")

        if self.active_view == "Model Categories":
            has_selection = selected_model_category is not None and len(self.selected_model_parameters) > 0
        elif self.active_view == "Annotation Categories":
            has_selection = selected_annotation_category is not None and len(self.selected_annotation_parameters) > 0
        elif self.active_view == "Elements":
            has_selection = selected_element_category is not None and len(selected_elements) > 0 and len(self.selected_element_parameters) > 0
        elif self.active_view == "Spatial":
            has_selection = selected_spatial_category is not None and len(selected_spatial_items) > 0 and len(self.selected_spatial_parameters) > 0
        elif self.active_view == "Preview/Edit":
            if getattr(self, "preview_import_mode", False):
                has_selection = getattr(self, "preview_table", None) is not None
            else:
                source_view = self._get_preview_source_view()
                if source_view == "Schedules":
                    has_selection = selected_item is not None
                elif source_view == "Model Categories":
                    has_selection = selected_model_category is not None and len(self.selected_model_parameters) > 0
                elif source_view == "Annotation Categories":
                    has_selection = selected_annotation_category is not None and len(self.selected_annotation_parameters) > 0
                elif source_view == "Elements":
                    has_selection = selected_element_category is not None and len(selected_elements) > 0 and len(self.selected_element_parameters) > 0
                elif source_view == "Spatial":
                    selected_spatial_category = self._get_selected_spatial_category_for_preview()
                    selected_spatial_items = self._get_selected_spatial_items_for_preview()
                    has_selection = selected_spatial_category is not None and len(selected_spatial_items) > 0 and len(self.selected_spatial_parameters) > 0
                else:
                    has_selection = False
        else:
            has_selection = selected_item is not None

        is_preview = self.active_view == "Preview/Edit"
        has_preview_edits = is_preview and getattr(self, "preview_table", None) is not None
        editable_preview_columns = self._get_preview_editable_column_count() if is_preview else 0

        self.btnResetValues.Visibility = Visibility.Visible
        self.btnOpenPreviewEdit.Visibility = Visibility.Collapsed if is_preview else Visibility.Visible
        self.btnExportProjectStandards.Visibility = Visibility.Visible if not is_preview else Visibility.Collapsed

        if is_preview:
            self.btnResetValues.ToolTip = "Reset Preview/Edit values to the current Revit model state"
            self.btnResetValues.IsEnabled = has_preview_edits
        else:
            self.btnResetValues.ToolTip = "Clear the current selections and parameter picks in this workspace"
            self.btnResetValues.IsEnabled = True

        if is_preview:
            self.btnImportExcel.Content = "Update Model"
            self.btnImportExcel.ToolTip = "Apply Preview/Edit changes to Revit"
            self.btnImportExcel.IsEnabled = (
                has_selection
                and editable_preview_columns > 0
                and getattr(self, "preview_pending_change_count", 0) > 0
            )
            self.btnExport.Visibility = Visibility.Collapsed
            self.btnExport.IsEnabled = False
            self.btnOpenPreviewEdit.IsEnabled = False
        else:
            self.btnImportExcel.Content = "Import ▾"
            self.btnImportExcel.ToolTip = "Preview or import Excel changes back into Revit"
            self.btnImportExcel.IsEnabled = can_use_excel_actions
            self.btnExport.Visibility = Visibility.Visible
            self.btnExport.IsEnabled = can_use_excel_actions and has_selection
            self.btnOpenPreviewEdit.IsEnabled = has_selection

        self.btnPreviewImport.IsEnabled = (
            self.active_view == "Preview/Edit"
            and has_selection
            and getattr(self, "preview_pending_change_count", 0) > 0
        )
        self.btnPreviewExport.IsEnabled = has_selection

    def _update_context_panels(self):
        schedule_count = len(self.all_schedules)
        selected_item = self._get_selected_item()
        selected_name = selected_item.Name if selected_item is not None else "no schedule selected"
        selected_category_item = self._get_selected_model_category_item()
        selected_category_name = selected_category_item.Name if selected_category_item is not None else "no model category selected"
        selected_annotation_category_item = self._get_selected_annotation_category_item()
        selected_annotation_category_name = selected_annotation_category_item.Name if selected_annotation_category_item is not None else "no annotation category selected"
        selected_element_category_item = self._get_selected_element_category_item()
        selected_element_category_name = selected_element_category_item.Name if selected_element_category_item is not None else "no element category selected"
        selected_element_count = len(self._get_selected_element_items())
        selected_spatial_category_item = self._get_selected_spatial_category_item()
        selected_spatial_category_name = selected_spatial_category_item.Name if selected_spatial_category_item is not None else "no spatial type selected"
        selected_spatial_count = len(self._get_selected_spatial_items())
        model_category_count = len(self.all_model_categories)
        annotation_category_count = len(self.all_annotation_categories)
        element_category_count = len(self.all_element_categories)
        spatial_category_count = len(self.all_spatial_categories)

        self.lblModelCategoriesInfo.Text = (
            "Model categories available: {}. Current selection: {}. Choose a category to detect instance and type parameters ready for Excel export/import.".format(
                model_category_count,
                selected_category_name
            )
        )
        self.lblModelCategoryHint.Text = (
            "Selection drives the parameter list and category-based Excel workflow. Current schedules available in the model: {}.".format(
                schedule_count
            )
        )
        self.lblModelSelectedHint.Text = "Only selected parameters will be used for Preview/Edit and Excel export. Current selection: {}.".format(
            len(self.selected_model_parameters)
        )
        self.lblAnnotationCategoriesInfo.Text = (
            "Annotation categories available: {}. Current selection: {}. Choose a category to detect annotation parameters ready for Excel export/import.".format(
                annotation_category_count,
                selected_annotation_category_name
            )
        )
        self.lblAnnotationCategoryHint.Text = (
            "Selection drives the parameter list and annotation-based Excel workflow. Current schedules available in the model: {}.".format(
                schedule_count
            )
        )
        self.lblAnnotationSelectedHint.Text = "Only selected parameters will be used for Preview/Edit and Excel export. Current selection: {}.".format(
            len(self.selected_annotation_parameters)
        )
        if self.active_view == "Spatial":
            self.lblElementsInfo.Text = (
                "Spatial types available: {}. Current selection: {}. Selected rooms/spaces: {}. Choose rooms or spaces and parameters for an Excel export.".format(
                    spatial_category_count,
                    selected_spatial_category_name,
                    selected_spatial_count
                )
            )
            self.lblElementCategoryHint.Text = (
                "Pick Rooms or Spaces first, then choose the exact items to export. Current schedules available in the model: {}.".format(
                    schedule_count
                )
            )
            self.lblElementSelectedHint.Text = "Only selected parameters will be used for Preview/Edit and Excel export. Current selection: {}.".format(
                len(self.selected_spatial_parameters)
            )
        else:
            self.lblElementsInfo.Text = (
                "Element categories available: {}. Current selection: {}. Selected elements: {}. Choose elements and parameters for an Excel export.".format(
                    element_category_count,
                    selected_element_category_name,
                    selected_element_count
                )
            )
            self.lblElementCategoryHint.Text = (
                "Pick a category first, then choose the exact elements to export. Current schedules available in the model: {}.".format(
                    schedule_count
                )
            )
            self.lblElementSelectedHint.Text = "Only selected parameters will be used for Preview/Edit and Excel export. Current selection: {}.".format(
                len(self.selected_element_parameters)
            )
        self.lblSpatialInfo.Text = (
            "Navigation is active. Spatial workflows for rooms and spaces can grow here without colliding with schedule export rules."
        )

    def _make_unique_preview_column_name(self, table, base_name):
        base_name = normalize_text(base_name) or "Column"
        candidate = base_name
        index = 2

        while table.Columns.Contains(candidate):
            candidate = "{} {}".format(base_name, index)
            index += 1

        return candidate

    def _clear_preview_grid(self):
        try:
            self.preview_table = None
            self.preview_row_elements = []
            self.preview_column_map = {}
            self.preview_original_values = {}
            self.preview_changed_marker_map = {}
            self.preview_pending_change_count = 0
            self.dgPreviewEdit.ItemsSource = None
        except Exception:
            pass

    def _reset_preview_import_state(self):
        self.preview_import_mode = False
        self.preview_import_table = None
        self.preview_import_source_path = ""
        self.preview_import_row_count = 0
        self.preview_import_editable_count = 0
        self.preview_changed_marker_map = {}

    def _bind_preview_table(self, table):
        self.preview_table = table
        was_suspended = getattr(self, "_suspend_preview_autosave", False)
        self._suspend_preview_autosave = True
        try:
            self.dgPreviewEdit.ItemsSource = None
            self.dgPreviewEdit.ItemsSource = table.DefaultView
            self._update_preview_change_state()
        finally:
            self._suspend_preview_autosave = was_suspended

    def _prune_empty_preview_columns(self, table, column_map):
        if table is None:
            return table

        removable_columns = []
        for column_name in list(column_map.keys()):
            if not table.Columns.Contains(column_name):
                continue

            column_info = column_map.get(column_name) or {}
            if is_editable_metadata_value(column_info.get("Editable", "")):
                continue

            has_value = False
            for row_index in range(table.Rows.Count):
                row = table.Rows[row_index]
                if normalize_text(row[column_name]):
                    has_value = True
                    break

            if not has_value:
                removable_columns.append(column_name)

        for column_name in removable_columns:
            try:
                table.Columns.Remove(column_name)
            except Exception:
                pass
            try:
                del column_map[column_name]
            except Exception:
                pass

        return table

    def _build_preview_table_from_import_data(self, data_info):
        table = DataTable()
        table.Columns.Add("ElementId")

        column_map = {
            "ElementId": {
                "Name": "ElementId",
                "Editable": "Visible",
                "Status": "Technical",
                "Metadata": {"ColumnRole": "ElementId", "Status": "Technical"}
            }
        }
        original_values = {}
        row_elements = []
        changed_marker_map = {}

        editable_columns = get_editable_columns_for_import_preview(data_info)
        preview_columns = []

        for col_info in editable_columns:
            metadata = col_info.get("Metadata") or {}
            base_name = normalize_text(col_info.get("Name", "")) or normalize_text(metadata.get("Name", "")) or "Column"
            column_name = self._make_unique_preview_column_name(table, base_name)
            table.Columns.Add(column_name)
            column_map[column_name] = {
                "Name": normalize_text(metadata.get("Name", "")) or base_name,
                "Editable": normalize_text(metadata.get("Editable", "")),
                "Status": normalize_text(metadata.get("Status", "")) or build_status(
                    metadata.get("Origin", ""),
                    metadata.get("Editable", "")
                ),
                "Metadata": metadata
            }
            marker_name = "__chg__{}".format(column_name)
            table.Columns.Add(marker_name, Object)
            changed_marker_map[column_name] = marker_name
            preview_columns.append((column_name, col_info))

        rows = list(data_info.get("Rows", []) or [])
        for row_index, row_info in enumerate(rows):
            values = row_info.get("Values", {}) or {}
            element, resolved_by = resolve_element_from_import_row(row_info)
            row_elements.append(element)

            row = table.NewRow()
            if element is not None:
                element_id_text = safe_text(get_element_id_value(element))
                if not element_id_text:
                    element_id_text = get_import_row_reference_value(row_info, "ElementId", 2)
            else:
                element_id_text = get_import_row_reference_value(row_info, "ElementId", 2)
            row["ElementId"] = element_id_text

            for column_name, col_info in preview_columns:
                excel_col = col_info.get("ExcelCol", 0)
                metadata = col_info.get("Metadata") or {}
                excel_value = normalize_text(values.get(excel_col, ""))
                row[column_name] = excel_value

                current_value = excel_value
                if element is not None:
                    try:
                        param, origin = get_parameter_from_metadata(element, metadata)
                        if param is not None:
                            current_value = normalize_text(get_parameter_preview_value(param))
                    except Exception:
                        current_value = excel_value

                original_values[(row_index, column_name)] = current_value
                row[changed_marker_map[column_name]] = (excel_value != current_value)

            table.Rows.Add(row)

        self.preview_row_elements = row_elements
        self.preview_column_map = column_map
        self.preview_original_values = original_values
        self.preview_changed_marker_map = changed_marker_map
        self.preview_import_row_count = table.Rows.Count
        self.preview_import_editable_count = len(preview_columns)
        return self._prune_empty_preview_columns(table, column_map)

    def _get_preview_source_view(self):
        if self.active_view == "Preview/Edit":
            return getattr(self, "preview_source_view", "Model Categories")
        return self.active_view

    def _get_selected_spatial_category_for_preview(self):
        try:
            return self.lstElementCategories.SelectedItem
        except Exception:
            return None

    def _get_selected_spatial_items_for_preview(self):
        try:
            return list(self.lstElementItems.SelectedItems)
        except Exception:
            return []

    def _dedupe_preview_parameter_items(self, parameter_items):
        deduped_by_name = {}
        order = []

        for item in list(parameter_items or []):
            name_key = normalize_text(getattr(item, "Name", "")).lower()
            if not name_key:
                continue

            if name_key not in deduped_by_name:
                deduped_by_name[name_key] = item
                order.append(name_key)
                continue

            current = deduped_by_name.get(name_key)
            current_editable = is_editable_metadata_value(getattr(current, "Editable", ""))
            item_editable = is_editable_metadata_value(getattr(item, "Editable", ""))

            # Revit can expose the same display name twice. Keep the writable
            # parameter when possible so Preview/Edit and Excel avoid twins.
            if item_editable and not current_editable:
                deduped_by_name[name_key] = item

        return [deduped_by_name[key] for key in order]

    def _is_sheets_category_item(self, category_item):
        name = normalize_text(getattr(category_item, "Name", "")).lower()
        if name == "sheets":
            return True

        try:
            category_name = normalize_text(category_item.Category.Name).lower()
            return category_name == "sheets"
        except Exception:
            return False

    def _find_best_parameter_by_name(self, parameter_items, parameter_names):
        desired_names = set([normalize_text(x).lower() for x in parameter_names])
        fallback_item = None

        for item in list(parameter_items or []):
            item_name = normalize_text(getattr(item, "Name", "")).lower()
            if item_name not in desired_names:
                continue

            if fallback_item is None:
                fallback_item = item

            if is_editable_metadata_value(getattr(item, "Editable", "")):
                return item

        return fallback_item

    def _with_required_model_context_parameters(self, category_item, parameter_items):
        parameter_items = self._dedupe_preview_parameter_items(parameter_items)

        if not self._is_sheets_category_item(category_item):
            return parameter_items

        if self._find_best_parameter_by_name(parameter_items, ["Sheet Name", "Name"]) is not None:
            return parameter_items

        sheet_name_parameter = self._find_best_parameter_by_name(
            list(getattr(self, "current_model_parameters", []) or []),
            ["Sheet Name", "Name"]
        )
        if sheet_name_parameter is None:
            return parameter_items

        # Sheets need their title next to the number/sort fields; otherwise
        # Preview/Edit and exported workbooks lose the main human-readable key.
        return self._dedupe_preview_parameter_items([sheet_name_parameter] + list(parameter_items))

    def _should_add_preview_name_column(self, parameter_items):
        display_name_fields = set(["name", "sheet name", "view name"])

        for item in list(parameter_items or []):
            if normalize_text(getattr(item, "Name", "")).lower() in display_name_fields:
                return False

        return True

    def _build_preview_table_from_elements(self, elements, parameter_items, include_type_id=False):
        parameter_items = self._dedupe_preview_parameter_items(parameter_items)
        table = DataTable()
        table.Columns.Add("ElementId")
        if include_type_id:
            table.Columns.Add("Type ID")
        self.preview_changed_marker_map = {}
        include_name_column = self._should_add_preview_name_column(parameter_items)
        if include_name_column:
            table.Columns.Add("Name")

        parameter_columns = []
        column_map = {}
        row_elements = []
        original_values = {}

        for param_item in list(parameter_items or []):
            column_name = self._make_unique_preview_column_name(table, getattr(param_item, "Name", "Parameter"))
            table.Columns.Add(column_name)
            metadata = getattr(param_item, "Metadata", {}) or {}
            column_map[column_name] = {
                "Metadata": metadata,
                "Name": getattr(param_item, "Name", column_name),
                "Editable": getattr(param_item, "Editable", ""),
                "Status": getattr(param_item, "Status", "")
            }
            parameter_columns.append((column_name, param_item))

        for element in list(elements or []):
            if element is None:
                continue

            row_index = len(row_elements)
            row_elements.append(element)
            row = table.NewRow()

            row["ElementId"] = safe_text(get_element_id_value(element))

            if include_type_id:
                row["Type ID"] = safe_text(get_element_type_id_value(element))

            if include_name_column:
                row["Name"] = get_element_display_name(element)

            for column_name, param_item in parameter_columns:
                metadata = getattr(param_item, "Metadata", {}) or {}
                value = get_export_value_from_metadata(element, metadata)
                row[column_name] = value
                original_values[(row_index, column_name)] = normalize_text(value)

            table.Rows.Add(row)

        self.preview_row_elements = row_elements
        self.preview_column_map = column_map
        self.preview_original_values = original_values
        return self._prune_empty_preview_columns(table, column_map)

    def _build_preview_table_from_schedule(self, schedule):
        table = DataTable()
        table.Columns.Add("ElementId")
        self.preview_changed_marker_map = {}

        visible_fields = get_visible_schedule_fields(schedule)
        field_columns = []
        column_map = {}
        row_elements = []
        original_values = {}

        for field_info in visible_fields:
            metadata = field_info.get("metadata") or {}
            field_name = (
                field_info.get("name")
                or metadata.get("Name")
                or "Field"
            )
            column_name = self._make_unique_preview_column_name(table, field_name)
            table.Columns.Add(column_name)
            column_map[column_name] = {
                "Metadata": metadata,
                "Name": field_name,
                "Editable": field_info.get("editable", metadata.get("Editable", "")),
                "Status": metadata.get("Status", build_status(metadata.get("Origin", ""), field_info.get("editable", metadata.get("Editable", ""))))
            }
            field_columns.append((column_name, field_info))

        for element in get_schedule_elements(schedule):
            if element is None:
                continue

            row_index = len(row_elements)
            row_elements.append(element)
            row = table.NewRow()

            row["ElementId"] = safe_text(get_element_id_value(element))

            for column_name, field_info in field_columns:
                value = get_schedule_field_export_value(element, field_info)
                row[column_name] = value
                original_values[(row_index, column_name)] = normalize_text(value)

            table.Rows.Add(row)

        self.preview_row_elements = row_elements
        self.preview_column_map = column_map
        self.preview_original_values = original_values
        return self._prune_empty_preview_columns(table, column_map)

    def _refresh_preview_grid(self):
        if self.active_view != "Preview/Edit":
            return 0

        try:
            self._set_progress(20, "Collecting preview data")
            if getattr(self, "preview_import_mode", False):
                table = getattr(self, "preview_import_table", None)
                if table is not None:
                    self._set_progress(60, "Building import preview")
                    self._bind_preview_table(table)
                    self._set_progress(100, "Preview ready")
                    return table.Rows.Count

            source_view = self._get_preview_source_view()

            selected_item = self._get_selected_item()
            if source_view == "Schedules" and selected_item is not None:
                self._set_progress(60, "Building schedule preview")
                table = self._build_preview_table_from_schedule(selected_item.Schedule)
                self._bind_preview_table(table)
                self._set_progress(100, "Preview ready")
                return table.Rows.Count

            selected_model_category = self._get_selected_model_category_item()
            if source_view == "Model Categories" and selected_model_category is not None and len(self.selected_model_parameters) > 0:
                self._set_progress(60, "Building model category preview")
                table = self._build_preview_table_from_elements(
                    get_category_elements_for_scope(selected_model_category.Category, self.model_scope),
                    self._with_required_model_context_parameters(
                        selected_model_category,
                        self.selected_model_parameters
                    ),
                    include_type_id=self.model_export_type_id
                )
                self._bind_preview_table(table)
                self._set_progress(100, "Preview ready")
                return table.Rows.Count

            selected_annotation_category = self._get_selected_annotation_category_item()
            if source_view == "Annotation Categories" and selected_annotation_category is not None and len(self.selected_annotation_parameters) > 0:
                self._set_progress(60, "Building annotation preview")
                table = self._build_preview_table_from_elements(
                    get_category_elements_for_scope(selected_annotation_category.Category, self.annotation_scope),
                    self.selected_annotation_parameters,
                    include_type_id=self.annotation_export_type_id
                )
                self._bind_preview_table(table)
                self._set_progress(100, "Preview ready")
                return table.Rows.Count

            selected_element_category = self._get_selected_element_category_item()
            selected_elements = self._get_selected_element_items()
            if source_view == "Elements" and selected_element_category is not None and len(selected_elements) > 0 and len(self.selected_element_parameters) > 0:
                self._set_progress(60, "Building element preview")
                elements = []
                for item in selected_elements:
                    element = getattr(item, "Element", None)
                    if element is not None:
                        elements.append(element)
                elements = filter_elements_by_scope(elements, self.elements_scope)

                table = self._build_preview_table_from_elements(elements, self.selected_element_parameters, include_type_id=self.elements_export_type_id)
                self._bind_preview_table(table)
                self._set_progress(100, "Preview ready")
                return table.Rows.Count

            selected_spatial_category = self._get_selected_spatial_category_for_preview()
            selected_spatial_items = self._get_selected_spatial_items_for_preview()
            if source_view == "Spatial" and selected_spatial_category is not None and len(selected_spatial_items) > 0 and len(self.selected_spatial_parameters) > 0:
                self._set_progress(60, "Building spatial preview")
                elements = []
                for item in selected_spatial_items:
                    element = getattr(item, "Element", None)
                    if element is not None:
                        elements.append(element)

                table = self._build_preview_table_from_elements(elements, self.selected_spatial_parameters)
                self._bind_preview_table(table)
                self._set_progress(100, "Preview ready")
                return table.Rows.Count

            self._clear_preview_grid()
            self._set_progress(100, "Preview ready")
            return 0

        except Exception as ex:
            self._clear_preview_grid()
            try:
                self.lblPreviewWorkflow.Text = "Preview could not be built: {}".format(safe_text(ex))
            except Exception:
                pass
            logger.warning("Could not build Preview/Edit grid. %s", safe_text(ex))
            self._set_progress(0, "Preview failed")
            return 0

    def _has_pending_preview_changes(self, table, column_map, original_values):
        if table is None or not column_map:
            return False

        for row_index in range(table.Rows.Count):
            row = table.Rows[row_index]
            for column_name in column_map.keys():
                if not table.Columns.Contains(column_name):
                    continue
                new_value = normalize_text(row[column_name])
                old_value = original_values.get((row_index, column_name), "")
                if new_value != old_value:
                    return True

        return False

    def _get_pending_preview_change_count(self, table, column_map, original_values):
        if table is None or not column_map:
            return 0

        count = 0
        for row_index in range(table.Rows.Count):
            row = table.Rows[row_index]
            for column_name in column_map.keys():
                if not table.Columns.Contains(column_name):
                    continue
                new_value = normalize_text(row[column_name])
                old_value = original_values.get((row_index, column_name), "")
                if new_value != old_value:
                    count += 1
        return count

    def _update_preview_change_state(self):
        table = getattr(self, "preview_table", None)
        column_map = getattr(self, "preview_column_map", {}) or {}
        original_values = getattr(self, "preview_original_values", {}) or {}
        pending = self._get_pending_preview_change_count(table, column_map, original_values)
        self.preview_pending_change_count = pending
        try:
            self.btnPreviewImport.IsEnabled = pending > 0
        except Exception:
            pass
        try:
            if self.active_view == "Preview/Edit":
                self.btnImportExcel.IsEnabled = pending > 0
        except Exception:
            pass
        return pending

    def _build_preview_workflow_message(self, editable_preview_columns, pending_changes, context_label):
        if editable_preview_columns == 0:
            return (
                "All selected {} parameters are locked by Revit. You can review them here, "
                "but they will not update the model. Red means locked; beige marks the key ID column.".format(context_label)
            )
        if pending_changes > 0:
            return (
                "{} pending change(s). Review the editable cells and press 'Update Model' "
                "when you are ready. Green means editable instance values, yellow means editable type values.".format(pending_changes)
            )
        return (
            "Editable cells can be changed here. Green means editable instance values, "
            "yellow means editable type values, red means locked, and beige marks the key ID column."
        )

    def _get_preview_column_style_key(self, header, column_info):
        if header in ("ElementId", "Type ID"):
            return "PreviewElementIdCellStyle"
        if header == "Name":
            return "PreviewNameCellStyle"
        if column_info is None:
            return "PreviewNameCellStyle"
        if not is_editable_metadata_value(column_info.get("Editable", "")):
            return "PreviewLockedCellStyle"
        origin = normalize_text((column_info.get("Metadata", {}) or {}).get("Origin", ""))
        if origin.lower() == "type":
            return "PreviewTypeCellStyle"
        return "PreviewEditableCellStyle"

    def _format_preview_column_header(self, header, column_info):
        if header in ("ElementId", "Type ID", "Name") or column_info is None:
            return header
        status = normalize_text(column_info.get("Status", "")) or "Locked"
        return "{}\n{}".format(header, status)

    def _apply_preview_column_style(self, column, style_key):
        if not style_key:
            return
        try:
            base_style = self.FindResource(style_key)
            marker_map = getattr(self, "preview_changed_marker_map", {}) or {}
            marker_name = marker_map.get(normalize_text(getattr(column, "SortMemberPath", "")))
            if getattr(self, "preview_import_mode", False) and marker_name:
                highlight_style = Style(DataGridCell)
                highlight_style.BasedOn = base_style
                trigger = DataTrigger()
                trigger.Binding = Binding("[{}]".format(marker_name))
                trigger.Value = True
                brush_converter = BrushConverter()
                trigger.Setters.Add(Setter(DataGridCell.BackgroundProperty, brush_converter.ConvertFrom("#118C8C")))
                trigger.Setters.Add(Setter(DataGridCell.ForegroundProperty, brush_converter.ConvertFrom("#FFFFFF")))
                highlight_style.Triggers.Add(trigger)
                column.CellStyle = highlight_style
            else:
                column.CellStyle = base_style
        except Exception:
            pass

    def _commit_preview_edits(self, show_dialog=True):
        if getattr(self, "_is_committing_preview_edits", False):
            return

        if self.active_view != "Preview/Edit":
            if show_dialog:
                TaskDialog.Show(__title__, "Open Preview/Edit first.")
            return

        self._is_committing_preview_edits = True
        try:
            try:
                self.dgPreviewEdit.CommitEdit()
                self.dgPreviewEdit.CommitEdit()
            except Exception:
                pass

            table = getattr(self, "preview_table", None)
            row_elements = list(getattr(self, "preview_row_elements", []) or [])
            column_map = getattr(self, "preview_column_map", {}) or {}
            original_values = getattr(self, "preview_original_values", {}) or {}

            if table is None or not row_elements or not column_map:
                if show_dialog:
                    TaskDialog.Show(__title__, "Nothing editable was found in Preview/Edit.")
                return

            if not self._has_pending_preview_changes(table, column_map, original_values):
                if show_dialog:
                    TaskDialog.Show(
                        __title__,
                        "No changes to update.\n\nEverything in Preview/Edit already matches the model."
                    )
                return

            updated = 0
            unchanged = 0
            locked = 0
            missing = 0
            failed = 0
            duplicate_count = 0
            failed_lines = []
            duplicate_lines = []

            transaction = DB.Transaction(doc, "SheetLink Preview/Edit Update")
            started = False

            try:
                self._set_progress(35, "Saving Preview/Edit changes")
                transaction.Start()
                started = True

                for row_index in range(table.Rows.Count):
                    if row_index >= len(row_elements):
                        continue

                    element = row_elements[row_index]
                    if element is None:
                        missing += 1
                        continue

                    row = table.Rows[row_index]

                    for column_name, column_info in column_map.items():
                        try:
                            if not table.Columns.Contains(column_name):
                                continue

                            new_value = normalize_text(row[column_name])
                            old_value = original_values.get((row_index, column_name), "")

                            if new_value == old_value:
                                unchanged += 1
                                continue

                            editable = normalize_text(column_info.get("Editable", ""))
                            if not is_editable_metadata_value(editable):
                                locked += 1
                                continue

                            metadata = column_info.get("Metadata") or {}
                            param, param_origin = get_parameter_from_metadata(element, metadata)
                            if param is None:
                                missing += 1
                                continue

                            try:
                                if param.IsReadOnly:
                                    locked += 1
                                    continue
                            except Exception:
                                locked += 1
                                continue

                            field_name = column_info.get("Name", column_name)
                            if is_unique_controlled_field(field_name):
                                is_duplicate, duplicate_message = value_exists_in_other_elements(field_name, new_value, element)
                                if is_duplicate:
                                    duplicate_count += 1
                                    if len(duplicate_lines) < 8:
                                        duplicate_lines.append(
                                            "{} | {} | {}".format(
                                                get_element_display_name(element),
                                                field_name,
                                                duplicate_message
                                            )
                                        )
                                    continue

                            if set_parameter_value(param, new_value):
                                refreshed_value = get_parameter_preview_value(param)
                                row[column_name] = refreshed_value
                                original_values[(row_index, column_name)] = normalize_text(refreshed_value)
                                marker_name = (getattr(self, "preview_changed_marker_map", {}) or {}).get(column_name)
                                if marker_name and table.Columns.Contains(marker_name):
                                    try:
                                        row[marker_name] = False
                                    except Exception:
                                        pass
                                updated += 1
                            else:
                                failed += 1
                                if len(failed_lines) < 8:
                                    failed_lines.append(
                                        "{} | {}".format(
                                            get_element_display_name(element),
                                            field_name
                                        )
                                    )

                        except Exception as cell_ex:
                            failed += 1
                            if len(failed_lines) < 8:
                                failed_lines.append(
                                    "{} | {} | {}".format(
                                        get_element_display_name(element),
                                        column_name,
                                        safe_text(cell_ex)
                                    )
                                )

                transaction.Commit()
                started = False

            except Exception as ex:
                if started:
                    try:
                        transaction.RollBack()
                    except Exception:
                        pass
                self._set_progress(100, "Update failed")
                if show_dialog:
                    TaskDialog.Show(__title__, "Preview/Edit update failed.\n\n{}".format(safe_text(ex)))
                else:
                    try:
                        self.lblPreviewWorkflow.Text = "Auto-save failed: {}".format(safe_text(ex))
                    except Exception:
                        pass
                return

            self.preview_original_values = original_values
            self._update_preview_change_state()
            self._set_progress(100, "Saved")

            has_issues = missing or duplicate_count or failed
            has_locked_info = locked > 0
            only_noop = updated == 0 and (unchanged > 0 or locked > 0) and missing == 0 and duplicate_count == 0 and failed == 0
            message = []

            if updated == 0 and not has_issues and not has_locked_info:
                message.append("No changes to update.")
                message.append("")
                message.append("Everything in Preview/Edit already matches the model.")
            else:
                message.append("Preview/Edit update complete.")
                message.append("")
                if updated:
                    message.append("Updated values: {}".format(updated))
                if unchanged and not updated:
                    message.append("Unchanged cells: {}".format(unchanged))
                if unchanged and updated:
                    message.append("Unchanged cells: {}".format(unchanged))
                if has_locked_info:
                    message.append("Read-only skipped: {}".format(locked))
                if has_issues:
                    message.append("")
                    message.append("Needs attention:")
                    if missing:
                        message.append("Missing parameters: {}".format(missing))
                    if duplicate_count:
                        message.append("Duplicate conflicts: {}".format(duplicate_count))
                    if failed:
                        message.append("Failed writes: {}".format(failed))
                elif has_locked_info:
                    message.append("")
                    message.append("Locked parameters were left unchanged, which is expected.")

            if duplicate_lines:
                message.append("")
                message.append("Duplicate conflict examples:")
                for line in duplicate_lines:
                    message.append(line)

            if failed_lines:
                message.append("")
                message.append("Failed write examples:")
                for line in failed_lines:
                    message.append(line)

            if show_dialog:
                if only_noop:
                    try:
                        detail_bits = []
                        if unchanged:
                            detail_bits.append("{} unchanged".format(unchanged))
                        if locked:
                            detail_bits.append("{} locked".format(locked))
                        summary = "No editable changes to apply."
                        if detail_bits:
                            summary = "{} {}".format(summary, " | ".join(detail_bits))
                        self.lblPreviewWorkflow.Text = summary
                    except Exception:
                        pass
                elif updated > 0 and not has_issues:
                    TaskDialog.Show(__title__, "Model updated successfully.")
                else:
                    TaskDialog.Show(__title__, "\n".join(message))
                self._refresh_preview_grid()
                self._update_preview_panel()
            else:
                try:
                    if updated:
                        self.lblPreviewWorkflow.Text = "Auto-saved to Revit. Updated values: {}".format(updated)
                    elif only_noop:
                        detail_bits = []
                        if unchanged:
                            detail_bits.append("{} unchanged".format(unchanged))
                        if locked:
                            detail_bits.append("{} locked".format(locked))
                        self.lblPreviewWorkflow.Text = "No editable changes to apply. {}".format(" | ".join(detail_bits))
                    elif has_issues:
                        self.lblPreviewWorkflow.Text = "Auto-save finished with warnings. Review locked, missing or duplicate values."
                    elif has_locked_info:
                        self.lblPreviewWorkflow.Text = "Update complete. {} locked values were skipped as expected.".format(locked)
                except Exception:
                    pass

            self._update_status()
        finally:
            self._is_committing_preview_edits = False

    def _get_preview_editable_column_count(self):
        count = 0
        for column_info in (getattr(self, "preview_column_map", {}) or {}).values():
            if is_editable_metadata_value(column_info.get("Editable", "")):
                count += 1
        return count

    def on_preview_auto_generating_column(self, sender, args):
        header = normalize_text(getattr(args.Column, "Header", ""))
        if header.startswith("__chg__"):
            args.Cancel = True
            return
        column_info = (getattr(self, "preview_column_map", {}) or {}).get(header)
        args.Column.Header = self._format_preview_column_header(header, column_info)
        self._apply_preview_column_style(args.Column, self._get_preview_column_style_key(header, column_info))

        if header in ("ElementId", "Type ID"):
            args.Column.IsReadOnly = True
            try:
                args.Column.Width = 100
            except Exception:
                pass
            return
        if header == "Name":
            args.Column.IsReadOnly = True
            try:
                args.Column.Width = 240
            except Exception:
                pass
            return

        if column_info is not None and not is_editable_metadata_value(column_info.get("Editable", "")):
            args.Column.IsReadOnly = True
            return

        args.Column.IsReadOnly = False
        try:
            binding = Binding(header)
            binding.Mode = BindingMode.TwoWay
            binding.UpdateSourceTrigger = UpdateSourceTrigger.LostFocus
            args.Column.Binding = binding
        except Exception:
            pass

    def on_preview_current_cell_changed(self, sender, args):
        if getattr(self, "_suspend_preview_autosave", False):
            return
        if self.active_view != "Preview/Edit":
            return
        self._update_preview_change_state()

    def on_preview_update_clicked(self, sender, args):
        self._commit_preview_edits()

    def on_reset_values_clicked(self, sender, args):
        if self.active_view == "Preview/Edit":
            try:
                self.dgPreviewEdit.CancelEdit()
            except Exception:
                pass

            self._set_progress(35, "Resetting preview values")
            if getattr(self, "preview_import_mode", False):
                self._reset_preview_import_state()
            self._refresh_preview_grid()
            try:
                self.lblPreviewWorkflow.Text = "Preview values reset to the current Revit model state."
            except Exception:
                pass
            self._update_preview_panel()
            self._update_action_buttons()
            self._update_status()
            self._update_preview_change_state()
            self._set_progress(100, "Values reset")
            return

        self._set_progress(25, "Resetting workspace")

        if self.active_view == "Schedules":
            try:
                self.lstSchedules.SelectedItem = None
            except Exception:
                pass
            self.current_parameters = []
            self.filtered_parameters = []
            self._refresh_parameter_grid([])
        elif self.active_view == "Model Categories":
            try:
                self.lstModelCategories.SelectedItem = None
            except Exception:
                pass
            self.current_model_parameters = []
            self.filtered_model_parameters = []
            self.selected_model_parameters = []
            self.filtered_selected_model_parameters = []
            self._refresh_model_parameter_grid([])
            self._refresh_model_selected_parameter_list([])
        elif self.active_view == "Annotation Categories":
            try:
                self.lstAnnotationCategories.SelectedItem = None
            except Exception:
                pass
            self.current_annotation_parameters = []
            self.filtered_annotation_parameters = []
            self.selected_annotation_parameters = []
            self.filtered_selected_annotation_parameters = []
            self._refresh_annotation_parameter_grid([])
            self._refresh_annotation_selected_parameter_list([])
        elif self.active_view == "Elements":
            try:
                self.lstElementCategories.SelectedItem = None
            except Exception:
                pass
            self.current_element_items = []
            self.filtered_element_items = []
            self.current_element_parameters = []
            self.filtered_element_parameters = []
            self.selected_element_parameters = []
            self.filtered_selected_element_parameters = []
            self._refresh_element_list()
            self._refresh_element_parameter_grid([])
            self._refresh_element_selected_parameter_list([])
        elif self.active_view == "Spatial":
            try:
                self.lstElementCategories.SelectedItem = None
            except Exception:
                pass
            self.current_spatial_items = []
            self.filtered_spatial_items = []
            self.current_spatial_parameters = []
            self.filtered_spatial_parameters = []
            self.selected_spatial_parameters = []
            self.filtered_selected_spatial_parameters = []
            self._refresh_spatial_list()
            self._refresh_spatial_parameter_grid([])
            self._refresh_spatial_selected_parameter_list([])

        self._clear_preview_grid()
        self._update_context_panels()
        self._update_preview_panel()
        self._update_action_buttons()
        self._update_status()
        self._set_progress(100, "Workspace reset")

    def on_open_preview_edit_clicked(self, sender, args):
        self._set_active_view("Preview/Edit")

    def on_export_project_standards_clicked(self, sender, args):
        if not os.path.exists(PROJECT_STANDARDS_XAML_PATH):
            TaskDialog.Show(__title__, "Project Standards dialog file is missing.")
            return

        dialog = ProjectStandardsWindow(PROJECT_STANDARDS_XAML_PATH)
        dialog.ShowDialog()

        if not dialog.result_confirmed:
            return

        if not dialog.selected_sections:
            TaskDialog.Show(__title__, "Select at least one Project Standards category.")
            return

        full_path = ask_output_xlsx_path_for_name("Project_Standards")
        if not full_path:
            return

        self._set_progress(15, "Preparing standards export")

        try:
            export_ok = export_project_standards_to_xlsx(full_path, dialog.selected_sections, dialog.open_after_export)
        except Exception as export_error:
            self._set_progress(100, "Ready")
            TaskDialog.Show(
                __title__,
                "Project Standards export failed.\n\n{}".format(safe_text(export_error))
            )
            return

        if not export_ok:
            self._set_progress(100, "Ready")
            return

        self._set_progress(100, "Project standards exported")
        TaskDialog.Show(__title__, "Project standards exported successfully.")

    def _update_preview_panel(self):
        selected_item = self._get_selected_item()
        selected_model_category = self._get_selected_model_category_item()
        selected_annotation_category = self._get_selected_annotation_category_item()
        selected_element_category = self._get_selected_element_category_item()
        selected_elements = self._get_selected_element_items()
        selected_spatial_category = self._get_selected_spatial_category_item()
        selected_spatial_items = self._get_selected_spatial_items()
        source_view = self._get_preview_source_view()

        if self.active_view == "Preview/Edit":
            if source_view != "Schedules":
                selected_item = None
            if source_view != "Model Categories":
                selected_model_category = None
            if source_view != "Annotation Categories":
                selected_annotation_category = None
            if source_view != "Elements":
                selected_element_category = None
                selected_elements = []
            if source_view == "Spatial":
                selected_spatial_category = self._get_selected_spatial_category_for_preview()
                selected_spatial_items = self._get_selected_spatial_items_for_preview()
            else:
                selected_spatial_category = None
                selected_spatial_items = []

        preview_rows = self._refresh_preview_grid()
        editable_preview_columns = self._get_preview_editable_column_count()
        pending_changes = self._update_preview_change_state() if self.active_view == "Preview/Edit" else 0

        if self.active_view == "Preview/Edit" and getattr(self, "preview_import_mode", False):
            source_name = os.path.basename(getattr(self, "preview_import_source_path", "") or "")
            self.lblPreviewSchedule.Text = "Preview Import"
            self.lblPreviewParameterCount.Text = "Rows loaded: {} | Editable fields: {}".format(
                preview_rows,
                editable_preview_columns
            )
            if pending_changes > 0:
                self.lblPreviewWorkflow.Text = (
                    "{} imported change(s) are ready for review. Check the grid and press 'Update Model' only when you want to write them into Revit.".format(
                        pending_changes
                    )
                )
            else:
                self.lblPreviewWorkflow.Text = (
                    "Imported workbook loaded from '{}'. Nothing is written to Revit until you press 'Update Model'.".format(
                        source_name or "Excel"
                    )
                )
            return

        if self.active_view == "Spatial" and selected_spatial_category is not None:
            self.lblPreviewSchedule.Text = "Spatial: {}".format(selected_spatial_category.Name)
            self.lblPreviewParameterCount.Text = "Rooms/spaces selected: {} | Parameters selected: {}".format(
                len(selected_spatial_items),
                len(self.selected_spatial_parameters)
            )
            self.lblPreviewWorkflow.Text = self._build_preview_workflow_message(
                editable_preview_columns,
                pending_changes,
                "spatial"
            )
            return

        if self.active_view == "Elements" and selected_element_category is not None:
            self.lblPreviewSchedule.Text = "Elements: {}".format(selected_element_category.Name)
            self.lblPreviewParameterCount.Text = "Elements selected: {} | Parameters selected: {}".format(
                len(selected_elements),
                len(self.selected_element_parameters)
            )
            self.lblPreviewWorkflow.Text = self._build_preview_workflow_message(
                editable_preview_columns,
                pending_changes,
                "element"
            )
            return

        if selected_item is None and selected_model_category is None and selected_annotation_category is None and selected_element_category is None and selected_spatial_category is None:
            self.lblPreviewSchedule.Text = "No schedule selected"
            self.lblPreviewParameterCount.Text = "Parameters ready: 0"
            self.lblPreviewWorkflow.Text = (
                "Pick a schedule, category, element set, or spatial set first. Once selected, this view becomes a quick launch point for export and Excel import."
            )
            return

        if selected_item is not None:
            parameter_count = len(self.current_parameters)
            self.lblPreviewSchedule.Text = selected_item.Name
            self.lblPreviewParameterCount.Text = "Rows previewed: {} | Parameters ready: {}".format(
                preview_rows,
                parameter_count
            )
            self.lblPreviewWorkflow.Text = self._build_preview_workflow_message(
                editable_preview_columns,
                pending_changes,
                "schedule"
            )
            return

        if selected_model_category is not None:
            self.lblPreviewSchedule.Text = "Model Category: {}".format(selected_model_category.Name)
            self.lblPreviewParameterCount.Text = "Rows previewed: {} | Parameters selected: {}".format(
                preview_rows,
                len(self.selected_model_parameters)
            )
            self.lblPreviewWorkflow.Text = self._build_preview_workflow_message(
                editable_preview_columns,
                pending_changes,
                "model-category"
            )
            return

        if selected_annotation_category is not None:
            self.lblPreviewSchedule.Text = "Annotation Category: {}".format(selected_annotation_category.Name)
            self.lblPreviewParameterCount.Text = "Rows previewed: {} | Parameters selected: {}".format(
                preview_rows,
                len(self.selected_annotation_parameters)
            )
            self.lblPreviewWorkflow.Text = self._build_preview_workflow_message(
                editable_preview_columns,
                pending_changes,
                "annotation"
            )
            return

        if selected_spatial_category is not None:
            self.lblPreviewSchedule.Text = "Spatial: {}".format(selected_spatial_category.Name)
            self.lblPreviewParameterCount.Text = "Rooms/spaces selected: {} | Parameters selected: {}".format(
                len(selected_spatial_items),
                len(self.selected_spatial_parameters)
            )
            self.lblPreviewWorkflow.Text = (
                "Current room/space selection is ready for a targeted Excel export."
            )
            return

        self.lblPreviewSchedule.Text = "Elements: {}".format(selected_element_category.Name)
        self.lblPreviewParameterCount.Text = "Rows previewed: {} | Elements selected: {} | Parameters selected: {}".format(
            preview_rows,
            len(selected_elements),
            len(self.selected_element_parameters)
        )
        if self.selected_element_parameters and editable_preview_columns == 0:
            self.lblPreviewWorkflow.Text = (
                "All selected element parameters are locked by Revit. Pick at least one green editable parameter if you want to edit in this grid."
            )
        else:
            self.lblPreviewWorkflow.Text = (
                "Editable cells can be changed directly here. Locked columns are marked and protected."
            )

    def _get_list_item_key(self, item):
        if item is None:
            return ""

        for attr_name in ("Schedule", "Element", "Category"):
            try:
                source = getattr(item, attr_name, None)
                if source is not None:
                    source_id = getattr(source, "Id", None)
                    if source_id is not None:
                        try:
                            return "{}:{}".format(attr_name, get_element_id_value(source_id))
                        except Exception:
                            return "{}:{}".format(attr_name, safe_text(source_id))
            except Exception:
                pass

        return normalize_text(getattr(item, "Name", ""))

    def _refresh_items_preserving_selection(self, control, items):
        selected_item = None
        try:
            selected_item = control.SelectedItem
        except Exception:
            pass

        selected_key = self._get_list_item_key(selected_item)
        was_suspended = getattr(self, "_suspend_selection_changed", False)
        self._suspend_selection_changed = True
        try:
            control.ItemsSource = None
            control.ItemsSource = items
            if selected_key:
                for item in list(items or []):
                    if item is selected_item or self._get_list_item_key(item) == selected_key:
                        try:
                            control.SelectedItem = item
                        except Exception:
                            pass
                        break
        finally:
            self._suspend_selection_changed = was_suspended

    def _refresh_schedule_list(self):
        self._refresh_items_preserving_selection(self.lstSchedules, self.filtered_schedules)

    def _refresh_parameter_grid(self, parameter_items):
        self.dgParameters.ItemsSource = None
        self.dgParameters.ItemsSource = parameter_items

    def _refresh_model_category_list(self):
        self._refresh_items_preserving_selection(self.lstModelCategories, self.filtered_model_categories)

    def _refresh_model_parameter_grid(self, parameter_items):
        self.lstModelAvailableParameters.ItemsSource = None
        self.lstModelAvailableParameters.ItemsSource = parameter_items

    def _refresh_model_selected_parameter_list(self, parameter_items):
        self.lstModelSelectedParameters.ItemsSource = None
        self.lstModelSelectedParameters.ItemsSource = parameter_items

    def _refresh_annotation_category_list(self):
        self._refresh_items_preserving_selection(self.lstAnnotationCategories, self.filtered_annotation_categories)

    def _refresh_annotation_parameter_grid(self, parameter_items):
        self.lstAnnotationAvailableParameters.ItemsSource = None
        self.lstAnnotationAvailableParameters.ItemsSource = parameter_items

    def _refresh_annotation_selected_parameter_list(self, parameter_items):
        self.lstAnnotationSelectedParameters.ItemsSource = None
        self.lstAnnotationSelectedParameters.ItemsSource = parameter_items

    def _refresh_element_category_list(self):
        self._refresh_items_preserving_selection(self.lstElementCategories, self.filtered_element_categories)

    def _refresh_element_list(self):
        self._refresh_items_preserving_selection(self.lstElementItems, self.filtered_element_items)

    def _refresh_element_parameter_grid(self, parameter_items):
        self.lstElementAvailableParameters.ItemsSource = None
        self.lstElementAvailableParameters.ItemsSource = parameter_items

    def _refresh_element_selected_parameter_list(self, parameter_items):
        self.lstElementSelectedParameters.ItemsSource = None
        self.lstElementSelectedParameters.ItemsSource = parameter_items

    def _refresh_spatial_category_list(self):
        if self.active_view != "Spatial":
            return
        self._refresh_items_preserving_selection(self.lstElementCategories, self.filtered_spatial_categories)

    def _refresh_spatial_list(self):
        if self.active_view != "Spatial":
            return
        self._refresh_items_preserving_selection(self.lstElementItems, self.filtered_spatial_items)

    def _refresh_spatial_parameter_grid(self, parameter_items):
        if self.active_view != "Spatial":
            return
        self.lstElementAvailableParameters.ItemsSource = None
        self.lstElementAvailableParameters.ItemsSource = parameter_items

    def _refresh_spatial_selected_parameter_list(self, parameter_items):
        if self.active_view != "Spatial":
            return
        self.lstElementSelectedParameters.ItemsSource = None
        self.lstElementSelectedParameters.ItemsSource = parameter_items

    def _update_status(self):
        total_count = len(self.all_schedules)
        if self.active_view == "Model Categories":
            field_count = len(self.filtered_selected_model_parameters)
        elif self.active_view == "Annotation Categories":
            field_count = len(self.filtered_selected_annotation_parameters)
        elif self.active_view == "Elements":
            field_count = len(self.filtered_selected_element_parameters)
        elif self.active_view == "Spatial":
            field_count = len(self.filtered_selected_spatial_parameters)
        else:
            field_count = len(self.filtered_parameters)
        selected_item = self._get_selected_item()
        selected_count = 1 if selected_item is not None else 0
        selected_model_category = self._get_selected_model_category_item()
        selected_model_count = 1 if selected_model_category is not None else 0
        selected_annotation_category = self._get_selected_annotation_category_item()
        selected_annotation_count = 1 if selected_annotation_category is not None else 0
        selected_element_category = self._get_selected_element_category_item()
        selected_element_category_count = 1 if selected_element_category is not None else 0
        selected_element_count = len(self._get_selected_element_items())
        selected_spatial_category = self._get_selected_spatial_category_item()
        selected_spatial_category_count = 1 if selected_spatial_category is not None else 0
        selected_spatial_count = len(self._get_selected_spatial_items())

        self.lblStatus.Text = "View: {} | Schedules: {} | Fields: {}".format(
            self.active_view,
            total_count,
            field_count
        )

        if hasattr(self, "lblFooterSummary"):
            if self.active_view == "Preview/Edit":
                self.lblFooterSummary.Text = (
                    "Preview/Edit | selected schedules {} | parameters ready {}".format(
                        selected_count,
                        len(self.current_parameters)
                    )
                )
            elif self.active_view == "Schedules":
                self.lblFooterSummary.Text = (
                    "Schedules selected {} | parameters found {}".format(
                        selected_count,
                        field_count
                    )
                )
            elif self.active_view == "Model Categories":
                self.lblFooterSummary.Text = (
                    "Model categories selected {} | parameters found {} | parameters selected {}".format(
                        selected_model_count,
                        len(self.filtered_model_parameters),
                        len(self.selected_model_parameters)
                    )
                )
            elif self.active_view == "Annotation Categories":
                self.lblFooterSummary.Text = (
                    "Annotation categories selected {} | parameters found {} | parameters selected {}".format(
                        selected_annotation_count,
                        len(self.filtered_annotation_parameters),
                        len(self.selected_annotation_parameters)
                    )
                )
            elif self.active_view == "Elements":
                self.lblFooterSummary.Text = (
                    "Element categories selected {} | elements selected {} | parameters found {} | parameters selected {}".format(
                        selected_element_category_count,
                        selected_element_count,
                        len(self.filtered_element_parameters),
                        len(self.selected_element_parameters)
                    )
                )
            elif self.active_view == "Spatial":
                self.lblFooterSummary.Text = (
                    "Spatial types selected {} | rooms/spaces selected {} | parameters found {} | parameters selected {}".format(
                        selected_spatial_category_count,
                        selected_spatial_count,
                        len(self.filtered_spatial_parameters),
                        len(self.selected_spatial_parameters)
                    )
                )
            else:
                self.lblFooterSummary.Text = (
                    "{} | schedules available {}".format(
                        self.active_view,
                        total_count
                    )
                )

    def _apply_schedule_filter(self):
        search = safe_text(self.txtSearchSchedules.Text).strip().lower()

        if not search:
            self.filtered_schedules = list(self.all_schedules)
        else:
            self.filtered_schedules = [
                x for x in self.all_schedules
                if search in x.Name.lower()
            ]

        self._refresh_schedule_list()
        self._update_status()

    def _apply_parameter_filter(self):
        search = safe_text(self.txtSearchParameters.Text).strip().lower()

        if not search:
            self.filtered_parameters = list(self.current_parameters)
        else:
            self.filtered_parameters = [
                x for x in self.current_parameters
                if search in x.Name.lower() or search in x.Status.lower()
            ]

        self._refresh_parameter_grid(self.filtered_parameters)
        self._update_status()

    def _apply_model_category_filter(self):
        search = safe_text(self.txtSearchModelCategories.Text).strip().lower()

        if not search:
            self.filtered_model_categories = list(self.all_model_categories)
        else:
            self.filtered_model_categories = [
                x for x in self.all_model_categories
                if search in x.Name.lower()
            ]

        self._refresh_model_category_list()
        self._update_status()

    def _apply_model_parameter_filter(self):
        search = safe_text(self.txtSearchModelParameters.Text).strip().lower()
        available_parameters = [x for x in self.current_model_parameters if not self._is_model_parameter_selected(x)]

        if not search:
            self.filtered_model_parameters = list(available_parameters)
        else:
            self.filtered_model_parameters = [
                x for x in available_parameters
                if search in x.Name.lower() or search in x.Status.lower()
            ]

        self._refresh_model_parameter_grid(self.filtered_model_parameters)
        self._update_status()

    def _apply_model_selected_parameter_filter(self):
        search = safe_text(self.txtSearchModelSelectedParameters.Text).strip().lower()

        if not search:
            self.filtered_selected_model_parameters = list(self.selected_model_parameters)
        else:
            self.filtered_selected_model_parameters = [
                x for x in self.selected_model_parameters
                if search in x.Name.lower() or search in x.Status.lower()
            ]

        self._refresh_model_selected_parameter_list(self.filtered_selected_model_parameters)
        self._update_status()
        self._update_preview_panel()
        self._update_action_buttons()

    def _apply_annotation_category_filter(self):
        search = safe_text(self.txtSearchAnnotationCategories.Text).strip().lower()

        if not search:
            self.filtered_annotation_categories = list(self.all_annotation_categories)
        else:
            self.filtered_annotation_categories = [
                x for x in self.all_annotation_categories
                if search in x.Name.lower()
            ]

        self._refresh_annotation_category_list()
        self._update_status()

    def _apply_annotation_parameter_filter(self):
        search = safe_text(self.txtSearchAnnotationParameters.Text).strip().lower()
        available_parameters = [x for x in self.current_annotation_parameters if not self._is_annotation_parameter_selected(x)]

        if not search:
            self.filtered_annotation_parameters = list(available_parameters)
        else:
            self.filtered_annotation_parameters = [
                x for x in available_parameters
                if search in x.Name.lower() or search in x.Status.lower()
            ]

        self._refresh_annotation_parameter_grid(self.filtered_annotation_parameters)
        self._update_status()

    def _apply_annotation_selected_parameter_filter(self):
        search = safe_text(self.txtSearchAnnotationSelectedParameters.Text).strip().lower()

        if not search:
            self.filtered_selected_annotation_parameters = list(self.selected_annotation_parameters)
        else:
            self.filtered_selected_annotation_parameters = [
                x for x in self.selected_annotation_parameters
                if search in x.Name.lower() or search in x.Status.lower()
            ]

        self._refresh_annotation_selected_parameter_list(self.filtered_selected_annotation_parameters)
        self._update_status()
        self._update_preview_panel()
        self._update_action_buttons()

    def _apply_element_category_filter(self):
        search = safe_text(self.txtSearchElementCategories.Text).strip().lower()

        if not search:
            self.filtered_element_categories = list(self.all_element_categories)
        else:
            self.filtered_element_categories = [
                x for x in self.all_element_categories
                if search in x.Name.lower()
            ]

        self._refresh_element_category_list()
        self._update_status()

    def _apply_element_item_filter(self):
        search = safe_text(self.txtSearchElementItems.Text).strip().lower()

        if not search:
            self.filtered_element_items = list(self.current_element_items)
        else:
            self.filtered_element_items = [
                x for x in self.current_element_items
                if search in x.Name.lower() or search in x.ElementId.lower()
            ]

        self._refresh_element_list()
        self._update_status()
        self._update_context_panels()
        self._update_action_buttons()

    def _apply_element_parameter_filter(self):
        search = safe_text(self.txtSearchElementParameters.Text).strip().lower()
        available_parameters = [x for x in self.current_element_parameters if not self._is_element_parameter_selected(x)]

        if not search:
            self.filtered_element_parameters = list(available_parameters)
        else:
            self.filtered_element_parameters = [
                x for x in available_parameters
                if search in x.Name.lower() or search in x.Status.lower()
            ]

        self._refresh_element_parameter_grid(self.filtered_element_parameters)
        self._update_status()

    def _apply_element_selected_parameter_filter(self):
        search = safe_text(self.txtSearchElementSelectedParameters.Text).strip().lower()

        if not search:
            self.filtered_selected_element_parameters = list(self.selected_element_parameters)
        else:
            self.filtered_selected_element_parameters = [
                x for x in self.selected_element_parameters
                if search in x.Name.lower() or search in x.Status.lower()
            ]

        self._refresh_element_selected_parameter_list(self.filtered_selected_element_parameters)
        self._update_status()
        self._update_context_panels()
        self._update_preview_panel()
        self._update_action_buttons()

    def _apply_spatial_category_filter(self):
        search = safe_text(self.txtSearchElementCategories.Text).strip().lower()

        if not search:
            self.filtered_spatial_categories = list(self.all_spatial_categories)
        else:
            self.filtered_spatial_categories = [
                x for x in self.all_spatial_categories
                if search in x.Name.lower()
            ]

        self._refresh_spatial_category_list()
        self._update_status()

    def _apply_spatial_item_filter(self):
        search = safe_text(self.txtSearchElementItems.Text).strip().lower()

        if not search:
            self.filtered_spatial_items = list(self.current_spatial_items)
        else:
            self.filtered_spatial_items = [
                x for x in self.current_spatial_items
                if search in x.Name.lower() or search in x.ElementId.lower()
            ]

        self._refresh_spatial_list()
        self._update_status()
        self._update_context_panels()
        self._update_action_buttons()

    def _apply_spatial_parameter_filter(self):
        search = safe_text(self.txtSearchElementParameters.Text).strip().lower()
        available_parameters = [x for x in self.current_spatial_parameters if not self._is_spatial_parameter_selected(x)]

        if not search:
            self.filtered_spatial_parameters = list(available_parameters)
        else:
            self.filtered_spatial_parameters = [
                x for x in available_parameters
                if search in x.Name.lower() or search in x.Status.lower()
            ]

        self._refresh_spatial_parameter_grid(self.filtered_spatial_parameters)
        self._update_status()

    def _apply_spatial_selected_parameter_filter(self):
        search = safe_text(self.txtSearchElementSelectedParameters.Text).strip().lower()

        if not search:
            self.filtered_selected_spatial_parameters = list(self.selected_spatial_parameters)
        else:
            self.filtered_selected_spatial_parameters = [
                x for x in self.selected_spatial_parameters
                if search in x.Name.lower() or search in x.Status.lower()
            ]

        self._refresh_spatial_selected_parameter_list(self.filtered_selected_spatial_parameters)
        self._update_status()
        self._update_context_panels()
        self._update_preview_panel()
        self._update_action_buttons()

    def _get_selected_item(self):
        try:
            return self.lstSchedules.SelectedItem
        except Exception:
            return None

    def _get_selected_model_category_item(self):
        try:
            return self.lstModelCategories.SelectedItem
        except Exception:
            return None

    def _get_selected_annotation_category_item(self):
        try:
            return self.lstAnnotationCategories.SelectedItem
        except Exception:
            return None

    def _get_selected_element_category_item(self):
        try:
            return self.lstElementCategories.SelectedItem
        except Exception:
            return None

    def _get_selected_element_items(self):
        try:
            return list(self.lstElementItems.SelectedItems)
        except Exception:
            return []

    def _get_selected_spatial_category_item(self):
        if self.active_view != "Spatial":
            return None
        try:
            return self.lstElementCategories.SelectedItem
        except Exception:
            return None

    def _get_selected_spatial_items(self):
        if self.active_view != "Spatial":
            return []
        try:
            return list(self.lstElementItems.SelectedItems)
        except Exception:
            return []

    def _get_selected_model_available_parameters(self):
        try:
            return list(self.lstModelAvailableParameters.SelectedItems)
        except Exception:
            return []

    def _get_selected_model_selected_parameters(self):
        try:
            return list(self.lstModelSelectedParameters.SelectedItems)
        except Exception:
            return []

    def _get_selected_annotation_available_parameters(self):
        try:
            return list(self.lstAnnotationAvailableParameters.SelectedItems)
        except Exception:
            return []

    def _get_selected_annotation_selected_parameters(self):
        try:
            return list(self.lstAnnotationSelectedParameters.SelectedItems)
        except Exception:
            return []

    def _get_selected_element_available_parameters(self):
        try:
            return list(self.lstElementAvailableParameters.SelectedItems)
        except Exception:
            return []

    def _get_selected_element_selected_parameters(self):
        try:
            return list(self.lstElementSelectedParameters.SelectedItems)
        except Exception:
            return []

    def _get_selected_spatial_available_parameters(self):
        if self.active_view != "Spatial":
            return []
        try:
            return list(self.lstElementAvailableParameters.SelectedItems)
        except Exception:
            return []

    def _get_selected_spatial_selected_parameters(self):
        if self.active_view != "Spatial":
            return []
        try:
            return list(self.lstElementSelectedParameters.SelectedItems)
        except Exception:
            return []

    def _get_model_parameter_signature(self, parameter_item):
        metadata = getattr(parameter_item, "Metadata", {}) or {}
        used_params = metadata.get("UsedParams", [])
        return (
            safe_text(getattr(parameter_item, "Name", "")),
            safe_text(getattr(parameter_item, "Origin", "")),
            safe_text(getattr(parameter_item, "Editable", "")),
            tuple([safe_text(x) for x in used_params])
        )

    def _is_model_parameter_selected(self, parameter_item):
        target_signature = self._get_model_parameter_signature(parameter_item)
        for selected_item in self.selected_model_parameters:
            if self._get_model_parameter_signature(selected_item) == target_signature:
                return True
        return False

    def _is_annotation_parameter_selected(self, parameter_item):
        target_signature = self._get_model_parameter_signature(parameter_item)
        for selected_item in self.selected_annotation_parameters:
            if self._get_model_parameter_signature(selected_item) == target_signature:
                return True
        return False

    def _is_element_parameter_selected(self, parameter_item):
        target_signature = self._get_model_parameter_signature(parameter_item)
        for selected_item in self.selected_element_parameters:
            if self._get_model_parameter_signature(selected_item) == target_signature:
                return True
        return False

    def _is_spatial_parameter_selected(self, parameter_item):
        target_signature = self._get_model_parameter_signature(parameter_item)
        for selected_item in self.selected_spatial_parameters:
            if self._get_model_parameter_signature(selected_item) == target_signature:
                return True
        return False

    def _sync_model_parameter_views(self):
        self._apply_model_parameter_filter()
        self._apply_model_selected_parameter_filter()
        self._update_context_panels()
        self._update_preview_panel()
        self._update_action_buttons()

    def _sync_annotation_parameter_views(self):
        self._apply_annotation_parameter_filter()
        self._apply_annotation_selected_parameter_filter()
        self._update_context_panels()
        self._update_preview_panel()
        self._update_action_buttons()

    def _sync_element_parameter_views(self):
        self._apply_element_parameter_filter()
        self._apply_element_selected_parameter_filter()
        self._update_context_panels()
        self._update_preview_panel()
        self._update_action_buttons()

    def _sync_spatial_parameter_views(self):
        self._apply_spatial_parameter_filter()
        self._apply_spatial_selected_parameter_filter()
        self._update_context_panels()
        self._update_preview_panel()
        self._update_action_buttons()

    def _move_selected_parameter(self, listbox, master_list, direction, sync_callback):
        if listbox is None or master_list is None or direction == 0:
            return

        selected_item = None
        try:
            selected_item = listbox.SelectedItem
        except Exception:
            selected_item = None

        if selected_item is None:
            return

        try:
            current_index = master_list.index(selected_item)
        except Exception:
            return

        target_index = current_index + direction
        if target_index < 0 or target_index >= len(master_list):
            return

        master_list[current_index], master_list[target_index] = master_list[target_index], master_list[current_index]

        try:
            sync_callback()
        except Exception:
            pass

        try:
            listbox.SelectedItem = selected_item
        except Exception:
            pass

        try:
            listbox.ScrollIntoView(selected_item)
        except Exception:
            pass

    def _load_model_parameters_from_item(self, item):
        self.selected_model_parameters = []
        self.filtered_selected_model_parameters = []

        if item is None:
            self.current_model_parameters = []
            self._set_progress(0, "Completed")
        else:
            self._set_progress(15, "Scanning parameters")
            self.current_model_parameters = self._dedupe_preview_parameter_items(
                get_category_parameters(item.Category)
            )
            self._set_progress(100, "Parameters ready")

        self._sync_model_parameter_views()

    def _load_annotation_parameters_from_item(self, item):
        self.selected_annotation_parameters = []
        self.filtered_selected_annotation_parameters = []

        if item is None:
            self.current_annotation_parameters = []
            self._set_progress(0, "Completed")
        else:
            self._set_progress(15, "Scanning parameters")
            self.current_annotation_parameters = self._dedupe_preview_parameter_items(
                get_category_parameters(item.Category)
            )
            self._set_progress(100, "Parameters ready")

        self._sync_annotation_parameter_views()

    def on_schedule_search_changed(self, sender, args):
        self._apply_schedule_filter()

    def on_parameter_search_changed(self, sender, args):
        self._apply_parameter_filter()

    def on_model_category_search_changed(self, sender, args):
        self._apply_model_category_filter()

    def on_model_parameter_search_changed(self, sender, args):
        self._apply_model_parameter_filter()

    def on_model_selected_parameter_search_changed(self, sender, args):
        self._apply_model_selected_parameter_filter()

    def on_annotation_category_search_changed(self, sender, args):
        self._apply_annotation_category_filter()

    def on_annotation_parameter_search_changed(self, sender, args):
        self._apply_annotation_parameter_filter()

    def on_annotation_selected_parameter_search_changed(self, sender, args):
        self._apply_annotation_selected_parameter_filter()

    def on_element_category_search_changed(self, sender, args):
        if self.active_view == "Spatial":
            self._apply_spatial_category_filter()
            return
        self._apply_element_category_filter()

    def on_element_item_search_changed(self, sender, args):
        if self.active_view == "Spatial":
            self._apply_spatial_item_filter()
            return
        self._apply_element_item_filter()

    def on_element_parameter_search_changed(self, sender, args):
        if self.active_view == "Spatial":
            self._apply_spatial_parameter_filter()
            return
        self._apply_element_parameter_filter()

    def on_element_selected_parameter_search_changed(self, sender, args):
        if self.active_view == "Spatial":
            self._apply_spatial_selected_parameter_filter()
            return
        self._apply_element_selected_parameter_filter()

    def on_nav_clicked(self, sender, args):
        view_name = safe_text(getattr(sender, "Tag", "")).strip()
        if view_name:
            self._set_active_view(view_name)

    def on_schedule_selection_changed(self, sender, args):
        if getattr(self, "_suspend_selection_changed", False):
            return
        item = self._get_selected_item()
        self._load_parameters_from_item(item)

    def on_model_category_selection_changed(self, sender, args):
        if getattr(self, "_suspend_selection_changed", False):
            return
        item = self._get_selected_model_category_item()
        self._load_model_parameters_from_item(item)

    def _get_scope_from_combo(self, combo, fallback):
        valid_scopes = ("Whole Model", "Active View", "Current Selection")

        try:
            selected_item = combo.SelectedItem
            selected_text = normalize_text(getattr(selected_item, "Content", selected_item))
            if selected_text in valid_scopes:
                return selected_text
        except Exception:
            pass

        return fallback

    def _refresh_preview_for_scope_if_active(self, source_view):
        self._update_context_panels()
        self._update_preview_panel()
        self._update_action_buttons()
        self._update_status()

        if self.active_view == "Preview/Edit" and self._get_preview_source_view() == source_view:
            self._refresh_preview_grid()

    def _category_item_is_available(self, category_item, category_items):
        category_key = self._get_list_item_key(category_item)
        if not category_key:
            return False

        for item in list(category_items or []):
            if self._get_list_item_key(item) == category_key:
                return True

        return False

    def _clear_model_category_state(self):
        try:
            self.lstModelCategories.SelectedItem = None
        except Exception:
            pass

        self.current_model_parameters = []
        self.filtered_model_parameters = []
        self.selected_model_parameters = []
        self.filtered_selected_model_parameters = []
        self._refresh_model_parameter_grid([])
        self._refresh_model_selected_parameter_list([])

    def _clear_annotation_category_state(self):
        try:
            self.lstAnnotationCategories.SelectedItem = None
        except Exception:
            pass

        self.current_annotation_parameters = []
        self.filtered_annotation_parameters = []
        self.selected_annotation_parameters = []
        self.filtered_selected_annotation_parameters = []
        self._refresh_annotation_parameter_grid([])
        self._refresh_annotation_selected_parameter_list([])

    def _clear_element_category_state(self):
        try:
            self.lstElementCategories.SelectedItem = None
        except Exception:
            pass

        self.current_element_items = []
        self.filtered_element_items = []
        self.current_element_parameters = []
        self.filtered_element_parameters = []
        self.selected_element_parameters = []
        self.filtered_selected_element_parameters = []
        self._refresh_element_list()
        self._refresh_element_parameter_grid([])
        self._refresh_element_selected_parameter_list([])

    def _rebuild_model_categories_for_scope(self):
        self._set_progress(20, "Applying model scope")
        self.all_model_categories = collect_model_categories(self.model_scope)
        self._apply_model_category_filter()

        selected_item = self._get_selected_model_category_item()
        if not self._category_item_is_available(selected_item, self.all_model_categories):
            self._clear_model_category_state()

        self._set_progress(100, "Scope ready")

    def _rebuild_annotation_categories_for_scope(self):
        self._set_progress(20, "Applying annotation scope")
        self.all_annotation_categories = collect_annotation_categories(self.annotation_scope)
        self.filtered_annotation_categories = list(self.all_annotation_categories)
        self.annotation_categories_loaded = True
        self._apply_annotation_category_filter()

        selected_item = self._get_selected_annotation_category_item()
        if not self._category_item_is_available(selected_item, self.all_annotation_categories):
            self._clear_annotation_category_state()

        self._set_progress(100, "Scope ready")

    def _rebuild_element_categories_for_scope(self):
        self._set_progress(20, "Applying element scope")
        self.all_element_categories = collect_element_categories(self.elements_scope)
        self.filtered_element_categories = list(self.all_element_categories)
        self.element_categories_loaded = True
        self._apply_element_category_filter()

        selected_item = self._get_selected_element_category_item()
        if not self._category_item_is_available(selected_item, self.all_element_categories):
            self._clear_element_category_state()
        else:
            selected_parameters = list(self.selected_element_parameters)
            self._load_elements_from_category_item(selected_item)
            self.selected_element_parameters = selected_parameters
            self._sync_element_parameter_views()

        self._set_progress(100, "Scope ready")

    def on_model_scope_changed(self, sender, args):
        self.model_scope = self._get_scope_from_combo(self.cmbModelScope, self.model_scope)
        self._rebuild_model_categories_for_scope()
        self._refresh_preview_for_scope_if_active("Model Categories")

    def _get_checkbox_bool(self, checkbox):
        try:
            return bool(checkbox.IsChecked)
        except Exception:
            return False

    def on_model_export_type_id_changed(self, sender, args):
        self.model_export_type_id = self._get_checkbox_bool(self.chkModelExportTypeId)
        self._refresh_preview_for_scope_if_active("Model Categories")

    def on_annotation_export_type_id_changed(self, sender, args):
        self.annotation_export_type_id = self._get_checkbox_bool(self.chkAnnotationExportTypeId)
        self._refresh_preview_for_scope_if_active("Annotation Categories")

    def on_elements_export_type_id_changed(self, sender, args):
        self.elements_export_type_id = self._get_checkbox_bool(self.chkElementsExportTypeId)
        self._refresh_preview_for_scope_if_active("Elements")

    def on_annotation_scope_changed(self, sender, args):
        self.annotation_scope = self._get_scope_from_combo(self.cmbAnnotationScope, self.annotation_scope)
        self._rebuild_annotation_categories_for_scope()
        self._refresh_preview_for_scope_if_active("Annotation Categories")

    def on_elements_scope_changed(self, sender, args):
        self.elements_scope = self._get_scope_from_combo(self.cmbElementsScope, self.elements_scope)
        if self.active_view != "Spatial":
            self._rebuild_element_categories_for_scope()
        self._refresh_preview_for_scope_if_active("Elements")

    def on_annotation_category_selection_changed(self, sender, args):
        if getattr(self, "_suspend_selection_changed", False):
            return
        item = self._get_selected_annotation_category_item()
        self._load_annotation_parameters_from_item(item)

    def on_element_category_selection_changed(self, sender, args):
        if getattr(self, "_suspend_selection_changed", False):
            return
        if self.active_view == "Spatial":
            item = self._get_selected_spatial_category_item()
            self._load_spatial_from_category_item(item)
            return
        item = self._get_selected_element_category_item()
        self._load_elements_from_category_item(item)

    def on_element_item_selection_changed(self, sender, args):
        if getattr(self, "_suspend_selection_changed", False):
            return
        self._update_context_panels()
        self._update_preview_panel()
        self._update_action_buttons()
        self._update_status()

    def on_model_add_parameter_clicked(self, sender, args):
        selected_items = self._get_selected_model_available_parameters()
        if not selected_items:
            return

        for item in selected_items:
            if not self._is_model_parameter_selected(item):
                self.selected_model_parameters.append(item)

        self._sync_model_parameter_views()

    def on_model_remove_parameter_clicked(self, sender, args):
        selected_items = self._get_selected_model_selected_parameters()
        if not selected_items:
            return

        remove_signatures = [self._get_model_parameter_signature(x) for x in selected_items]
        self.selected_model_parameters = [
            x for x in self.selected_model_parameters
            if self._get_model_parameter_signature(x) not in remove_signatures
        ]

        self._sync_model_parameter_views()

    def on_model_selected_move_up_clicked(self, sender, args):
        self._move_selected_parameter(
            self.lstModelSelectedParameters,
            self.selected_model_parameters,
            -1,
            self._sync_model_parameter_views
        )

    def on_model_selected_move_down_clicked(self, sender, args):
        self._move_selected_parameter(
            self.lstModelSelectedParameters,
            self.selected_model_parameters,
            1,
            self._sync_model_parameter_views
        )

    def on_annotation_add_parameter_clicked(self, sender, args):
        selected_items = self._get_selected_annotation_available_parameters()
        if not selected_items:
            return

        for item in selected_items:
            if not self._is_annotation_parameter_selected(item):
                self.selected_annotation_parameters.append(item)

        self._sync_annotation_parameter_views()

    def _load_elements_from_category_item(self, item):
        self.current_element_items = []
        self.filtered_element_items = []
        self.current_element_parameters = []
        self.filtered_element_parameters = []
        self.selected_element_parameters = []
        self.filtered_selected_element_parameters = []

        if item is None:
            self._set_progress(0, "Completed")
        else:
            self._set_progress(15, "Collecting elements")
            self.current_element_items = get_element_items_for_category_scope(item.Category, self.elements_scope)
            self.filtered_element_items = list(self.current_element_items)
            self._set_progress(50, "Scanning parameters")
            self.current_element_parameters = self._dedupe_preview_parameter_items(
                get_category_parameters(item.Category)
            )
            self._set_progress(100, "Parameters ready")

        self._refresh_element_list()
        self._sync_element_parameter_views()

    def _load_spatial_from_category_item(self, item):
        self.current_spatial_items = []
        self.filtered_spatial_items = []
        self.current_spatial_parameters = []
        self.filtered_spatial_parameters = []
        self.selected_spatial_parameters = []
        self.filtered_selected_spatial_parameters = []

        if item is None:
            self._set_progress(0, "Completed")
        else:
            self._set_progress(15, "Collecting rooms/spaces")
            self.current_spatial_items = get_element_items_for_category(item.Category)
            self.filtered_spatial_items = list(self.current_spatial_items)
            self._set_progress(50, "Scanning parameters")
            self.current_spatial_parameters = self._dedupe_preview_parameter_items(
                get_category_parameters(item.Category)
            )
            self._set_progress(100, "Parameters ready")

        self._refresh_spatial_list()
        self._sync_spatial_parameter_views()

    def on_annotation_remove_parameter_clicked(self, sender, args):
        selected_items = self._get_selected_annotation_selected_parameters()
        if not selected_items:
            return

        remove_signatures = [self._get_model_parameter_signature(x) for x in selected_items]
        self.selected_annotation_parameters = [
            x for x in self.selected_annotation_parameters
            if self._get_model_parameter_signature(x) not in remove_signatures
        ]

        self._sync_annotation_parameter_views()

    def on_annotation_selected_move_up_clicked(self, sender, args):
        self._move_selected_parameter(
            self.lstAnnotationSelectedParameters,
            self.selected_annotation_parameters,
            -1,
            self._sync_annotation_parameter_views
        )

    def on_annotation_selected_move_down_clicked(self, sender, args):
        self._move_selected_parameter(
            self.lstAnnotationSelectedParameters,
            self.selected_annotation_parameters,
            1,
            self._sync_annotation_parameter_views
        )

    def on_element_add_parameter_clicked(self, sender, args):
        if self.active_view == "Spatial":
            selected_items = self._get_selected_spatial_available_parameters()
            if not selected_items:
                return

            for item in selected_items:
                if not self._is_spatial_parameter_selected(item):
                    self.selected_spatial_parameters.append(item)

            self._sync_spatial_parameter_views()
            return

        selected_items = self._get_selected_element_available_parameters()
        if not selected_items:
            return

        for item in selected_items:
            if not self._is_element_parameter_selected(item):
                self.selected_element_parameters.append(item)

        self._sync_element_parameter_views()

    def on_element_remove_parameter_clicked(self, sender, args):
        if self.active_view == "Spatial":
            selected_items = self._get_selected_spatial_selected_parameters()
            if not selected_items:
                return

            remove_signatures = [self._get_model_parameter_signature(x) for x in selected_items]
            self.selected_spatial_parameters = [
                x for x in self.selected_spatial_parameters
                if self._get_model_parameter_signature(x) not in remove_signatures
            ]

            self._sync_spatial_parameter_views()
            return

        selected_items = self._get_selected_element_selected_parameters()
        if not selected_items:
            return

        remove_signatures = [self._get_model_parameter_signature(x) for x in selected_items]
        self.selected_element_parameters = [
            x for x in self.selected_element_parameters
            if self._get_model_parameter_signature(x) not in remove_signatures
        ]

        self._sync_element_parameter_views()

    def on_element_selected_move_up_clicked(self, sender, args):
        self._move_selected_parameter(
            self.lstElementSelectedParameters,
            self.selected_element_parameters,
            -1,
            self._sync_element_parameter_views
        )

    def on_element_selected_move_down_clicked(self, sender, args):
        self._move_selected_parameter(
            self.lstElementSelectedParameters,
            self.selected_element_parameters,
            1,
            self._sync_element_parameter_views
        )

    def on_refresh_clicked(self, sender, args):
        self._set_progress(10, "Refreshing schedules")
        self.all_schedules = collect_schedules()
        self.filtered_schedules = list(self.all_schedules)
        self.current_parameters = []
        self.filtered_parameters = []
        self._set_progress(45, "Refreshing categories")
        self.all_model_categories = collect_model_categories(self.model_scope)
        self.filtered_model_categories = list(self.all_model_categories)
        self.current_model_parameters = []
        self.filtered_model_parameters = []
        self.selected_model_parameters = []
        self.filtered_selected_model_parameters = []
        self.annotation_categories_loaded = False
        self.element_categories_loaded = False
        self.spatial_categories_loaded = False
        self.all_annotation_categories = []
        self.filtered_annotation_categories = []
        self.current_annotation_parameters = []
        self.filtered_annotation_parameters = []
        self.selected_annotation_parameters = []
        self.filtered_selected_annotation_parameters = []
        self.all_element_categories = []
        self.filtered_element_categories = []
        self.current_element_items = []
        self.filtered_element_items = []
        self.current_element_parameters = []
        self.filtered_element_parameters = []
        self.selected_element_parameters = []
        self.filtered_selected_element_parameters = []
        self.all_spatial_categories = []
        self.filtered_spatial_categories = []
        self.current_spatial_items = []
        self.filtered_spatial_items = []
        self.current_spatial_parameters = []
        self.filtered_spatial_parameters = []
        self.selected_spatial_parameters = []
        self.filtered_selected_spatial_parameters = []

        self.txtSearchSchedules.Text = ""
        self.txtSearchParameters.Text = ""
        self.txtSearchModelCategories.Text = ""
        self.txtSearchModelParameters.Text = ""
        self.txtSearchModelSelectedParameters.Text = ""
        self.txtSearchAnnotationCategories.Text = ""
        self.txtSearchAnnotationParameters.Text = ""
        self.txtSearchAnnotationSelectedParameters.Text = ""
        self.txtSearchElementCategories.Text = ""
        self.txtSearchElementItems.Text = ""
        self.txtSearchElementParameters.Text = ""
        self.txtSearchElementSelectedParameters.Text = ""

        self._set_progress(75, "Updating interface")
        self._ensure_view_data_loaded(self.active_view)
        self._refresh_schedule_list()
        self._refresh_parameter_grid([])
        self._refresh_model_category_list()
        self._refresh_model_parameter_grid([])
        self._refresh_model_selected_parameter_list([])
        if self.active_view == "Annotation Categories":
            self._refresh_annotation_category_list()
            self._refresh_annotation_parameter_grid([])
            self._refresh_annotation_selected_parameter_list([])
        elif self.active_view == "Elements":
            self._refresh_element_category_list()
            self._refresh_element_list()
            self._refresh_element_parameter_grid([])
            self._refresh_element_selected_parameter_list([])
        elif self.active_view == "Spatial":
            self._refresh_spatial_category_list()
            self._refresh_spatial_list()
            self._refresh_spatial_parameter_grid([])
            self._refresh_spatial_selected_parameter_list([])
        self._update_context_panels()
        self._update_preview_panel()
        self._update_action_buttons()
        self._update_status()
        self._set_progress(100, "Refresh completed")

    def on_import_excel_clicked(self, sender, args):
        if self.active_view == "Preview/Edit":
            self._commit_preview_edits()
            return

        try:
            ctx = self.btnImportExcel.ContextMenu
            if ctx is not None:
                ctx.PlacementTarget = self.btnImportExcel
                ctx.IsOpen = True
                return
        except Exception:
            pass

        self.on_import_from_excel_clicked(sender, args)

    def on_import_from_excel_clicked(self, sender, args):
        xlsx_path = ask_input_xlsx_path()
        if not xlsx_path:
            return

        try:
            self._set_progress(20, "Opening workbook")
            self._set_progress(55, "Applying changes")
            run_import_apply(xlsx_path)
            self._set_progress(100, "Import completed")

        except Exception as ex:
            self._set_progress(0, "Completed")
            TaskDialog.Show(__title__, "Import Excel failed.\n\n{}".format(safe_text(ex)))

    def on_preview_import_clicked(self, sender, args):
        xlsx_path = ask_input_xlsx_path()
        if not xlsx_path:
            return

        excel_app = None
        workbooks = None
        workbook = None

        try:
            self._set_progress(15, "Opening workbook")
            excel_app = create_excel_application()
            excel_app.Visible = False
            excel_app.DisplayAlerts = False

            workbooks = excel_app.Workbooks
            workbook = workbooks.Open(xlsx_path, False, True)

            self._set_progress(40, "Reading import data")
            data_info = read_data_sheet_for_import_preview(workbook)

            self._set_progress(65, "Building preview import")
            table = self._build_preview_table_from_import_data(data_info)
            self.preview_import_mode = True
            self.preview_import_table = table
            self.preview_import_source_path = xlsx_path

            self._set_active_view("Preview/Edit")
            self._bind_preview_table(table)
            self._update_preview_panel()
            self._update_action_buttons()
            self._update_status()
            self._set_progress(100, "Preview ready")

        except Exception as ex:
            self._reset_preview_import_state()
            self._set_progress(0, "Completed")
            TaskDialog.Show(__title__, "Preview Import failed.\n\n{}".format(safe_text(ex)))

        finally:
            try:
                if workbook is not None:
                    workbook.Close(False)
            except Exception:
                pass

            try:
                if excel_app is not None:
                    excel_app.Quit()
            except Exception:
                pass

            release_com_object(workbook)
            release_com_object(workbooks)
            release_com_object(excel_app)

    def on_export_clicked(self, sender, args):
        if self.active_view == "Model Categories":
            self._export_selected_model_category()
            return

        if self.active_view == "Annotation Categories":
            self._export_selected_annotation_category()
            return

        if self.active_view == "Elements":
            self._export_selected_elements()
            return

        if self.active_view == "Spatial":
            self._export_selected_spatial()
            return

        if self.active_view == "Preview/Edit":
            source_view = self._get_preview_source_view()
            selected_item = self._get_selected_item()
            selected_model_category = self._get_selected_model_category_item()
            selected_annotation_category = self._get_selected_annotation_category_item()
            selected_element_category = self._get_selected_element_category_item()

            if source_view == "Schedules" and selected_item is not None:
                pass
            elif source_view == "Model Categories" and selected_model_category is not None and len(self.selected_model_parameters) > 0:
                self._export_selected_model_category()
                return
            elif source_view == "Annotation Categories" and selected_annotation_category is not None and len(self.selected_annotation_parameters) > 0:
                self._export_selected_annotation_category()
                return
            elif source_view == "Elements" and selected_element_category is not None and len(self._get_selected_element_items()) > 0 and len(self.selected_element_parameters) > 0:
                self._export_selected_elements()
                return
            elif source_view == "Spatial" and self._get_selected_spatial_category_for_preview() is not None and len(self._get_selected_spatial_items_for_preview()) > 0 and len(self.selected_spatial_parameters) > 0:
                self._export_selected_spatial_from_preview()
                return

        selected_item = self._get_selected_item()
        if selected_item is None:
            TaskDialog.Show(__title__, "Select one schedule first.")
            return

        full_path = ask_output_xlsx_path_for_name(get_schedule_name(selected_item.Schedule))
        if not full_path:
            return

        try:
            self._set_progress(15, "Preparing export")
            full_path, has_valid_ids, id_source, data_row_count, schedule_element_count, skipped_group_rows = export_schedule_to_xlsx(
                selected_item.Schedule,
                full_path
            )
            self._set_progress(100, "Export completed")
        except Exception as ex:
            self._set_progress(0, "Completed")
            error_text = safe_text(ex)

            if "cannot access" in error_text.lower():
                TaskDialog.Show(
                    __title__,
                    "Excel export failed.\n\nExcel cannot access the target file.\nClose it if it is open and try again.\n\nDetails:\n{}".format(error_text)
                )
            else:
                TaskDialog.Show(__title__, "Excel export failed.\n\n{}".format(error_text))

            return

        open_path_with_default_app(full_path)

        summary_label = selected_item.Name
        if not has_valid_ids:
            summary_label = "{} | Warning: no import IDs".format(summary_label)

        self._set_completed_export_summary(summary_label, data_row_count, len(self.current_parameters))

    def _export_selected_model_category(self):
        selected_item = self._get_selected_model_category_item()
        if selected_item is None:
            TaskDialog.Show(__title__, "Select one model category first.")
            return

        parameter_items = self._with_required_model_context_parameters(
            selected_item,
            self.selected_model_parameters
        )
        if not parameter_items:
            TaskDialog.Show(__title__, "Select at least one parameter for the model category.")
            return

        full_path = ask_output_xlsx_path_for_name("{}_ModelCategory".format(selected_item.Name))
        if not full_path:
            return

        try:
            self._set_progress(15, "Preparing export")
            full_path, exported_row_count, category_element_count = export_category_to_xlsx(
                selected_item,
                parameter_items,
                full_path,
                source_elements=get_category_elements_for_scope(selected_item.Category, self.model_scope),
                include_type_id=self.model_export_type_id
            )
            self._set_progress(100, "Export completed")
        except Exception as ex:
            self._set_progress(0, "Completed")
            error_text = safe_text(ex)

            if "cannot access" in error_text.lower():
                TaskDialog.Show(
                    __title__,
                    "Excel export failed.\n\nExcel cannot access the target file.\nClose it if it is open and try again.\n\nDetails:\n{}".format(error_text)
                )
            else:
                TaskDialog.Show(__title__, "Excel export failed.\n\n{}".format(error_text))

            return

        open_path_with_default_app(full_path)
        self._set_completed_export_summary(selected_item.Name, exported_row_count, len(parameter_items))

    def _export_selected_elements(self):
        selected_category = self._get_selected_element_category_item()
        if selected_category is None:
            TaskDialog.Show(__title__, "Select one element category first.")
            return

        selected_elements = self._get_selected_element_items()
        if not selected_elements:
            TaskDialog.Show(__title__, "Select at least one element.")
            return

        parameter_items = self._dedupe_preview_parameter_items(self.selected_element_parameters)
        if not parameter_items:
            TaskDialog.Show(__title__, "Select at least one parameter for the selected elements.")
            return

        full_path = ask_output_xlsx_path_for_name("{}_Elements".format(selected_category.Name))
        if not full_path:
            return

        source_elements = [x.Element for x in selected_elements if getattr(x, "Element", None) is not None]
        source_elements = filter_elements_by_scope(source_elements, self.elements_scope)

        try:
            self._set_progress(15, "Preparing export")
            full_path, exported_row_count, category_element_count = export_category_to_xlsx(
                selected_category,
                parameter_items,
                full_path,
                source_elements=source_elements,
                sheet_suffix="Elements",
                keep_empty_rows=True,
                include_type_id=self.elements_export_type_id
            )
            self._set_progress(100, "Export completed")
        except Exception as ex:
            self._set_progress(0, "Completed")
            error_text = safe_text(ex)

            if "cannot access" in error_text.lower():
                TaskDialog.Show(
                    __title__,
                    "Excel export failed.\n\nExcel cannot access the target file.\nClose it if it is open and try again.\n\nDetails:\n{}".format(error_text)
                )
            else:
                TaskDialog.Show(__title__, "Excel export failed.\n\n{}".format(error_text))

            return

        open_path_with_default_app(full_path)
        self._set_completed_export_summary(
            "{} elements".format(selected_category.Name),
            exported_row_count,
            len(parameter_items)
        )

    def _export_selected_spatial_from_preview(self):
        selected_category = self._get_selected_spatial_category_for_preview()
        if selected_category is None:
            TaskDialog.Show(__title__, "Select Rooms or Spaces first.")
            return

        selected_items = self._get_selected_spatial_items_for_preview()
        if not selected_items:
            TaskDialog.Show(__title__, "Select at least one room or space.")
            return

        parameter_items = self._dedupe_preview_parameter_items(self.selected_spatial_parameters)
        if not parameter_items:
            TaskDialog.Show(__title__, "Select at least one parameter for the selected rooms/spaces.")
            return

        full_path = ask_output_xlsx_path_for_name("{}_Spatial".format(selected_category.Name))
        if not full_path:
            return

        source_elements = [x.Element for x in selected_items if getattr(x, "Element", None) is not None]

        try:
            self._set_progress(15, "Preparing export")
            full_path, exported_row_count, category_element_count = export_category_to_xlsx(
                selected_category,
                parameter_items,
                full_path,
                source_elements=source_elements,
                sheet_suffix="Spatial",
                keep_empty_rows=True
            )
            self._set_progress(100, "Export completed")
        except Exception as ex:
            self._set_progress(0, "Completed")
            error_text = safe_text(ex)

            if "cannot access" in error_text.lower():
                TaskDialog.Show(
                    __title__,
                    "Excel export failed.\n\nExcel cannot access the target file.\nClose it if it is open and try again.\n\nDetails:\n{}".format(error_text)
                )
            else:
                TaskDialog.Show(__title__, "Excel export failed.\n\n{}".format(error_text))

            return

        open_path_with_default_app(full_path)
        self._set_completed_export_summary(
            "{} spatial".format(selected_category.Name),
            exported_row_count,
            len(parameter_items)
        )

    def _export_selected_spatial(self):
        selected_category = self._get_selected_spatial_category_item()
        if selected_category is None:
            TaskDialog.Show(__title__, "Select Rooms or Spaces first.")
            return

        selected_items = self._get_selected_spatial_items()
        if not selected_items:
            TaskDialog.Show(__title__, "Select at least one room or space.")
            return

        parameter_items = self._dedupe_preview_parameter_items(self.selected_spatial_parameters)
        if not parameter_items:
            TaskDialog.Show(__title__, "Select at least one parameter for the selected rooms/spaces.")
            return

        full_path = ask_output_xlsx_path_for_name("{}_Spatial".format(selected_category.Name))
        if not full_path:
            return

        source_elements = [x.Element for x in selected_items if getattr(x, "Element", None) is not None]

        try:
            self._set_progress(15, "Preparing export")
            full_path, exported_row_count, category_element_count = export_category_to_xlsx(
                selected_category,
                parameter_items,
                full_path,
                source_elements=source_elements,
                sheet_suffix="Spatial",
                keep_empty_rows=True
            )
            self._set_progress(100, "Export completed")
        except Exception as ex:
            self._set_progress(0, "Completed")
            error_text = safe_text(ex)

            if "cannot access" in error_text.lower():
                TaskDialog.Show(
                    __title__,
                    "Excel export failed.\n\nExcel cannot access the target file.\nClose it if it is open and try again.\n\nDetails:\n{}".format(error_text)
                )
            else:
                TaskDialog.Show(__title__, "Excel export failed.\n\n{}".format(error_text))

            return

        open_path_with_default_app(full_path)
        self._set_completed_export_summary(
            "{} spatial".format(selected_category.Name),
            exported_row_count,
            len(parameter_items)
        )

    def _export_selected_annotation_category(self):
        selected_item = self._get_selected_annotation_category_item()
        if selected_item is None:
            TaskDialog.Show(__title__, "Select one annotation category first.")
            return

        parameter_items = self._dedupe_preview_parameter_items(self.selected_annotation_parameters)
        if not parameter_items:
            TaskDialog.Show(__title__, "Select at least one parameter for the annotation category.")
            return

        full_path = ask_output_xlsx_path_for_name("{}_AnnotationCategory".format(selected_item.Name))
        if not full_path:
            return

        try:
            self._set_progress(15, "Preparing export")
            full_path, exported_row_count, category_element_count = export_category_to_xlsx(
                selected_item,
                parameter_items,
                full_path,
                source_elements=get_category_elements_for_scope(selected_item.Category, self.annotation_scope),
                include_type_id=self.annotation_export_type_id
            )
            self._set_progress(100, "Export completed")
        except Exception as ex:
            self._set_progress(0, "Completed")
            error_text = safe_text(ex)

            if "cannot access" in error_text.lower():
                TaskDialog.Show(
                    __title__,
                    "Excel export failed.\n\nExcel cannot access the target file.\nClose it if it is open and try again.\n\nDetails:\n{}".format(error_text)
                )
            else:
                TaskDialog.Show(__title__, "Excel export failed.\n\n{}".format(error_text))

            return

        open_path_with_default_app(full_path)
        self._set_completed_export_summary(selected_item.Name, exported_row_count, len(parameter_items))

    def on_close_clicked(self, sender, args):
        self.Close()

    def on_go_schedules_clicked(self, sender, args):
        source_view = self._get_preview_source_view()
        if source_view == "Preview/Edit" or source_view not in self.view_panels:
            source_view = "Model Categories"
        self._set_active_view(source_view)

    def _load_parameters_from_item(self, item):
        if item is None:
            self.current_parameters = []
        else:
            self.current_parameters = get_schedule_parameters(item.Schedule)

        self._apply_parameter_filter()
        self._update_preview_panel()
        self._update_context_panels()
        self._update_action_buttons()

if not os.path.exists(XAML_PATH):
    TaskDialog.Show(__title__, "Missing XAML file:\n{}".format(XAML_FILENAME))
    sys.exit()

window = ScheduleBrowserWindow(XAML_PATH)
window.ShowDialog()









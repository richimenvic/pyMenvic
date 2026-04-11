# -*- coding: utf-8 -*-

__title__ = "Schedule Browser"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
==========================================================
pyMENVIC | SCHEDULE BROWSER
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
- Escribe UniqueId y ElementId en columnas técnicas cuando es posible
- Soporta fallback por elementos reales del schedule si el CSV no trae ElementId
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
- Si el CSV no trae ElementId, intenta resolver IDs desde los elementos reales del schedule
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

from pyrevit import forms
from pyrevit import script
from Autodesk.Revit import DB
from Autodesk.Revit.UI import TaskDialog

import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel
from System.Runtime.InteropServices import Marshal
from Microsoft.Win32 import SaveFileDialog, OpenFileDialog



logger = script.get_logger()

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

XAML_FILENAME = "window.xaml"
THIS_DIR = os.path.dirname(__file__)
XAML_PATH = os.path.join(THIS_DIR, XAML_FILENAME)

TEMP_FOLDER = os.path.expandvars("%temp%")
MAX_SAMPLE_ELEMENTS = 20
EXPORT_DELIMITER = ","

class ScheduleItem(object):
    def __init__(self, schedule):
        self.Schedule = schedule
        self.Name = get_schedule_name(schedule)


class ParameterItem(object):
    def __init__(self, name, origin, editable):
        self.Name = name
        self.Origin = origin
        self.Editable = editable
        self.Status = build_status(origin, editable)


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

def ask_output_xlsx_path(schedule):
    schedule_name = sanitize_filename(get_schedule_name(schedule))

    dialog = SaveFileDialog()
    dialog.Title = "Save Excel Export"
    dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
    dialog.FileName = "{}.xlsx".format(schedule_name)
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

            try:
                if param.SetValueString(value):
                    return True
            except Exception:
                pass

            param.Set(int(float(value)))
            return True

        if storage_type == DB.StorageType.Double:
            if value == "":
                return False

            try:
                if param.SetValueString(value):
                    return True
            except Exception:
                pass

            param.Set(float(value))
            return True

        if storage_type == DB.StorageType.ElementId:
            if value == "":
                param.Set(DB.ElementId.InvalidElementId)
            else:
                param.Set(DB.ElementId(int(float(value))))
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
    failed = 0
    duplicate_count = 0
    unresolved_count = 0
    missing_param_count = 0

    duplicate_lines = []
    unresolved_lines = []
    failed_lines = []

    try:
        excel_app = Excel.ApplicationClass()
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
                                    safe_text(element.Id.IntegerValue),
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
                                safe_text(element.Id.IntegerValue),
                                field_name,
                                excel_value
                            )
                        )

            if not row_changed:
                skipped += 1

        t.Commit()

        message = []
        message.append("Import completed.")
        message.append("")
        message.append("Updated values: {}".format(updated))
        message.append("Skipped: {}".format(skipped))
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
        return parameter_id.IntegerValue
    except Exception:
        return None


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
                if pid.IntegerValue == parameter_id_value:
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

    return sorted(results, key=lambda x: (x.Editable, x.Name.lower()))


def build_schedule_export_options():
    options = DB.ViewScheduleExportOptions()
    try:
        options.FieldDelimiter = EXPORT_DELIMITER
    except Exception:
        pass
    return options


def read_csv_rows(csv_path):
    rows = []
    with open(csv_path, "r") as fp:
        reader = csv.reader(fp, delimiter=EXPORT_DELIMITER)
        for row in reader:
            rows.append([normalize_text(x) for x in row])
    return rows


def release_com_object(obj):
    try:
        if obj is not None:
            Marshal.ReleaseComObject(obj)
    except Exception:
        pass


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
                return safe_text(eid.IntegerValue)
    except Exception:
        pass

    return ""


def get_parameter_from_metadata(element, metadata):
    if element is None or metadata is None:
        return None, None

    used_params = metadata.get("UsedParams", [])
    if not used_params:
        return None, None

    parameter_id_value = None
    try:
        parameter_id_value = int(used_params[0])
    except Exception:
        return None, None

    param = find_parameter_on_element(element, parameter_id_value)
    if param is not None:
        return param, "Instance"

    type_element = get_type_element(element)
    param = find_parameter_on_element(type_element, parameter_id_value)
    if param is not None:
        return param, "Type"

    return None, None


def read_data_sheet_for_import_preview(workbook):
    worksheet = get_worksheet_by_name(workbook, "Data")
    if worksheet is None:
        raise Exception("Missing worksheet: Data")

    last_row, last_col = get_last_used_row_col(worksheet)
    if last_row < 3 or last_col < 2:
        raise Exception("Worksheet 'Data' does not contain valid rows.")

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
                "Values": row_values
            })

    return {
        "Worksheet": worksheet,
        "LastRow": last_row,
        "LastCol": last_col,
        "MetadataByCol": metadata_by_col,
        "Rows": data_rows
    }


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
    values = row_info.get("Values", {})
    unique_id_text = normalize_text(values.get(1, ""))
    element_id_text = normalize_text(values.get(2, ""))

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

    try:
        current_id = current_element.Id.IntegerValue
    except Exception:
        current_id = None

    if field_name == normalize_text("View Name"):
        collector = DB.FilteredElementCollector(doc).OfClass(DB.View)
        for elem in collector:
            try:
                if elem is None:
                    continue
                if current_id is not None and elem.Id.IntegerValue == current_id:
                    continue
                if normalize_text(elem.Name) == new_value:
                    return True, "Already used by View Id {}".format(elem.Id.IntegerValue)
            except Exception:
                pass

    elif field_name == normalize_text("Sheet Number"):
        collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)
        for elem in collector:
            try:
                if elem is None:
                    continue
                if current_id is not None and elem.Id.IntegerValue == current_id:
                    continue

                param = elem.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER)
                if param is None:
                    continue

                existing_value = normalize_text(param.AsString())
                if existing_value == new_value:
                    return True, "Already used by Sheet Id {}".format(elem.Id.IntegerValue)
            except Exception:
                pass

    elif field_name == normalize_text("Level Name"):
        collector = DB.FilteredElementCollector(doc).OfClass(DB.Level)
        for elem in collector:
            try:
                if elem is None:
                    continue
                if current_id is not None and elem.Id.IntegerValue == current_id:
                    continue
                if normalize_text(elem.Name) == new_value:
                    return True, "Already used by Level Id {}".format(elem.Id.IntegerValue)
            except Exception:
                pass

    elif field_name == normalize_text("Grid Name"):
        collector = DB.FilteredElementCollector(doc).OfClass(DB.Grid)
        for elem in collector:
            try:
                if elem is None:
                    continue
                if current_id is not None and elem.Id.IntegerValue == current_id:
                    continue
                if normalize_text(elem.Name) == new_value:
                    return True, "Already used by Grid Id {}".format(elem.Id.IntegerValue)
            except Exception:
                pass

    return False, ""

def run_import_preview(xlsx_path):
    excel_app = None
    workbooks = None
    workbook = None

    try:
        excel_app = Excel.ApplicationClass()
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
                                    safe_text(element.Id.IntegerValue),
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
                            safe_text(element.Id.IntegerValue),
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

def export_schedule_to_xlsx(schedule, full_path):
    original_visible_fields = get_visible_schedule_fields(schedule)
    csv_path = export_schedule_to_temp_csv(schedule)
    csv_rows = read_csv_rows(csv_path)

    if not csv_rows:
        raise Exception("The schedule export returned no rows.")

    header_row_index = find_header_row(csv_rows)
    if header_row_index < 0:
        raise Exception("Could not detect the header row in the exported CSV.")

    csv_headers = csv_rows[header_row_index]
    csv_headers = [normalize_text(x) for x in csv_headers]

    csv_data_rows = []
    skipped_group_rows = 0

    for row in csv_rows[header_row_index + 1:]:
        cleaned = [normalize_text(x) for x in row]

        if not any(cleaned):
            continue

        if not row_has_real_schedule_content(cleaned, csv_headers):
            skipped_group_rows += 1
            continue

        csv_data_rows.append(cleaned)

    if not csv_headers:
        raise Exception("No headers were found in the exported CSV.")

    element_id_csv_index = find_element_id_column_index(csv_headers)

    id_source = "None"
    schedule_element_count = 0
    has_valid_element_ids = False

    if element_id_csv_index >= 0:
        element_id_values = build_element_id_values_from_rows(csv_data_rows, element_id_csv_index)
        unique_id_values, has_valid_element_ids = build_unique_id_map_from_rows(csv_data_rows, element_id_csv_index)
        id_source = "CSV"
    else:
        element_id_values, unique_id_values, has_valid_element_ids, schedule_element_count = build_id_data_from_schedule_elements(
            schedule,
            len(csv_data_rows)
        )

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

    orig_index = 0
    for i, header in enumerate(csv_headers):
        if i == element_id_csv_index:
            continue

        final_headers.append(header)

        if orig_index < len(original_visible_fields):
            md = dict(original_visible_fields[orig_index]["metadata"])
            export_columns.append({
                "ExcelIndex": len(export_columns) + 1,
                "Name": header,
                "Status": md.get("Status", "Unknown"),
                "Origin": md.get("Origin", "Special"),
                "Editable": md.get("Editable", "Unknown"),
                "ScheduleId": md.get("ScheduleId", ""),
                "UsedParams": md.get("UsedParams", []),
                "FieldIndex": md.get("FieldIndex", ""),
                "ColumnRole": "ScheduleField",
                "Hidden": False
            })
            orig_index += 1
        else:
            export_columns.append({
                "ExcelIndex": len(export_columns) + 1,
                "Name": header,
                "Status": "Unknown",
                "Origin": "Special",
                "Editable": "Unknown",
                "ScheduleId": "",
                "UsedParams": [],
                "FieldIndex": "",
                "ColumnRole": "ScheduleField",
                "Hidden": False
            })

    final_rows = []
    for r, row in enumerate(csv_data_rows):
        new_row = []

        unique_id = unique_id_values[r] if r < len(unique_id_values) else ""
        new_row.append(unique_id)

        element_id_value = element_id_values[r] if r < len(element_id_values) else ""
        new_row.append(element_id_value)

        for i, cell in enumerate(row):
            if i == element_id_csv_index:
                continue
            new_row.append(cell)

        final_rows.append(new_row)

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
    schema_sheet = None

    try:
        excel_app = Excel.ApplicationClass()
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

        workbooks = excel_app.Workbooks
        workbook = workbooks.Add()

        data_sheet = workbook.Worksheets[1]
        data_sheet.Name = "Data"

        total_cols = len(final_headers)

        for c in range(1, total_cols + 1):
            data_sheet.Cells[1, c].Value2 = metadata_row[c - 1]
            data_sheet.Cells[2, c].Value2 = final_headers[c - 1]

        for r in range(0, len(final_rows)):
            excel_row = r + 3
            row_vals = final_rows[r]

            for c in range(1, total_cols + 1):
                value = row_vals[c - 1] if c - 1 < len(row_vals) else ""
                cell = data_sheet.Cells[excel_row, c]
                cell.Value2 = value

                editable = export_columns[c - 1].get("Editable", "Unknown")
                role = export_columns[c - 1].get("ColumnRole", "")

                if role in ("UniqueId", "ElementId"):
                    cell.Locked = True
                else:
                    cell.Locked = editable != "Yes"

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

            data_range.Font.Color = 0x000000

            if role == "ElementId":
                data_range.Interior.Color = 0x1450BE
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

        data_sheet.Rows["2:2"].AutoFilter()
        fit_export_columns(data_sheet, export_columns, 8, 40, last_data_row)

        try:
            data_sheet.Columns[4].ColumnWidth = 24
        except Exception:
            pass

        try:
            data_sheet.Protect(
                "pyMenvic",
                True,
                True,
                True,
                False,
                True,
                True,
                False,
                False,
                False,
                False,
                False,
                False,
                True,
                True,
                False
            )
        except Exception:
            data_sheet.Protect("pyMenvic", True, True)

        schema_sheet = workbook.Worksheets.Add()
        schema_sheet.Name = "Schema"

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

        workbook.SaveAs(full_path)
        workbook.Close(True)
        excel_app.Quit()

        try:
            os.remove(csv_path)
        except Exception:
            pass

        return full_path, has_valid_element_ids, id_source, len(csv_data_rows), schedule_element_count, skipped_group_rows

    finally:
        release_com_object(schema_sheet)
        release_com_object(data_sheet)
        release_com_object(workbook)
        release_com_object(workbooks)
        release_com_object(excel_app)
class ScheduleBrowserWindow(forms.WPFWindow):
    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)

        self.all_schedules = collect_schedules()
        self.filtered_schedules = list(self.all_schedules)
        self.current_parameters = []
        self.filtered_parameters = []

        self._bind_events()
        self._configure_export_state()
        self._refresh_schedule_list()
        self._refresh_parameter_grid([])
        self._update_status()

    def _bind_events(self):
        self.txtSearchSchedules.TextChanged += self.on_schedule_search_changed
        self.txtSearchParameters.TextChanged += self.on_parameter_search_changed
        self.lstSchedules.SelectionChanged += self.on_schedule_selection_changed
        self.lstSchedules.MouseUp += self.on_schedule_list_mouse_up
        self.btnRefresh.Click += self.on_refresh_clicked
        self.btnImportExcel.Click += self.on_import_excel_clicked
        self.btnExport.Click += self.on_export_clicked
        self.btnClose.Click += self.on_close_clicked

    def _configure_export_state(self):
        self.btnExport.IsEnabled = True
        self.btnExport.ToolTip = "Export selected schedule to Excel"

    def _refresh_schedule_list(self):
        self.lstSchedules.ItemsSource = None
        self.lstSchedules.ItemsSource = self.filtered_schedules

    def _refresh_parameter_grid(self, parameter_items):
        self.dgParameters.ItemsSource = None
        self.dgParameters.ItemsSource = parameter_items

    def _update_status(self):
        total_count = len(self.all_schedules)
        field_count = len(self.filtered_parameters)

        self.lblStatus.Text = "Schedules: {} | Fields: {}".format(
            total_count,
            field_count
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

    def _get_selected_item(self):
        try:
            return self.lstSchedules.SelectedItem
        except Exception:
            return None

    def on_schedule_search_changed(self, sender, args):
        self._apply_schedule_filter()

    def on_parameter_search_changed(self, sender, args):
        self._apply_parameter_filter()

    def on_schedule_selection_changed(self, sender, args):
        item = self._get_selected_item()
        self._load_parameters_from_item(item)

    def on_refresh_clicked(self, sender, args):
        self.all_schedules = collect_schedules()
        self.filtered_schedules = list(self.all_schedules)
        self.current_parameters = []
        self.filtered_parameters = []

        self.txtSearchSchedules.Text = ""
        self.txtSearchParameters.Text = ""

        self._refresh_schedule_list()
        self._refresh_parameter_grid([])
        self._update_status()

    def on_import_excel_clicked(self, sender, args):
        xlsx_path = ask_input_xlsx_path()
        if not xlsx_path:
            return

        action = forms.CommandSwitchWindow.show(
            [
                "PREVIEW",
                "IMPORT TO REVIT",
                "CANCEL"
            ],
            message="Choose import action"
        )

        if not action or action == "CANCEL":
            return

        try:
            if action == "IMPORT TO REVIT":
                run_import_apply(xlsx_path)
            else:
                run_import_preview(xlsx_path)

        except Exception as ex:
            TaskDialog.Show(__title__, "Import Excel failed.\n\n{}".format(safe_text(ex)))

    def on_export_clicked(self, sender, args):
        selected_item = self._get_selected_item()
        if selected_item is None:
            TaskDialog.Show(__title__, "Select one schedule first.")
            return

        full_path = ask_output_xlsx_path(selected_item.Schedule)
        if not full_path:
            return

        try:
            full_path, has_valid_ids, id_source, data_row_count, schedule_element_count, skipped_group_rows = export_schedule_to_xlsx(
                selected_item.Schedule,
                full_path
            )
        except Exception as ex:
            error_text = safe_text(ex)

            if "0x800A03EC" in error_text or "cannot access" in error_text.lower():
                TaskDialog.Show(
                    __title__,
                    "Excel export failed.\n\nThe Excel file is open.\nClose it before exporting again."
                )
            else:
                TaskDialog.Show(__title__, "Excel export failed.\n\n{}".format(error_text))

            return

        message = "Excel export completed.\n\n{}".format(full_path)

        if skipped_group_rows > 0:
            message += "\n\nFiltered rows:\n{} group/header rows were excluded from export.".format(skipped_group_rows)

        if id_source == "CSV":
            if not has_valid_ids:
                message += "\n\nWarning:\nElementId was found in the export, but UniqueId could not be resolved from the rows."
        else:
            message += "\n\nWarning:\nNo reliable ElementId / UniqueId source was found for this schedule."
            if schedule_element_count > 0:
                message += "\nThe fallback by schedule row order was disabled to avoid assigning IDs to the wrong elements."
                message += "\nSchedule elements detected: {}".format(schedule_element_count)

        try:
            subprocess.Popen([full_path], shell=True)
        except Exception:
            pass

        TaskDialog.Show(__title__, message)

    def on_close_clicked(self, sender, args):
        self.Close()

    def _load_parameters_from_item(self, item):
        if item is None:
            self.current_parameters = []
        else:
            self.current_parameters = get_schedule_parameters(item.Schedule)

        self._apply_parameter_filter()

    def on_schedule_list_mouse_up(self, sender, args):
        item = self._get_selected_item()
        self._load_parameters_from_item(item)
if not os.path.exists(XAML_PATH):
    TaskDialog.Show(__title__, "Missing XAML file:\n{}".format(XAML_FILENAME))
    sys.exit()

window = ScheduleBrowserWindow(XAML_PATH)
window.ShowDialog()




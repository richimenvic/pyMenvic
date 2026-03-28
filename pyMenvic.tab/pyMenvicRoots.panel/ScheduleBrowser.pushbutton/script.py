# -*- coding: utf-8 -*-

__title__ = "Schedule Browser"
__author__ = "Ricardo J. Mendieta"

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


logger = script.get_logger()

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

XAML_FILENAME = "window.xaml"
THIS_DIR = os.path.dirname(__file__)
XAML_PATH = os.path.join(THIS_DIR, XAML_FILENAME)

TEMP_FOLDER = os.path.expandvars("%temp%")
MAX_SAMPLE_ELEMENTS = 20
EXPORT_DELIMITER = ","
TEMP_ELEMENTID_FIELD_NAME = "Element ID"


class ScheduleItem(object):
    def __init__(self, schedule):
        self.Schedule = schedule
        self.Name = get_schedule_name(schedule)
        self.IsChecked = False


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
    if "editable" in text:
        return 0xD9EAD3
    if "locked" in text:
        return 0xF4CCCC
    return 0xE6E6E6


def auto_fit_used_columns(worksheet, column_count):
    for i in range(1, column_count + 1):
        try:
            worksheet.Columns[i].AutoFit()
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


def export_schedule_to_temp_csv(schedule):
    temp_field_id = add_temp_elementid_field(schedule)

    schedule_name = get_schedule_name(schedule)
    file_name = "{}_{}.csv".format(
        sanitize_filename(schedule_name),
        str(int(time.time() * 1000))
    )
    full_path = os.path.join(TEMP_FOLDER, file_name)

    try:
        options = build_schedule_export_options()
        schedule.Export(TEMP_FOLDER, file_name, options)
    finally:
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


def export_schedule_to_xlsx(schedule):
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
    for row in csv_rows[header_row_index + 1:]:
        cleaned = [normalize_text(x) for x in row]
        if any(cleaned):
            csv_data_rows.append(cleaned)

    if not csv_headers:
        raise Exception("No headers were found in the exported CSV.")

    element_id_csv_index = find_element_id_column_index(csv_headers)
    unique_id_values, has_valid_element_ids = build_unique_id_map_from_rows(csv_data_rows, element_id_csv_index)

    # Build final headers
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

    # Build final rows
    final_rows = []
    for r, row in enumerate(csv_data_rows):
        new_row = []

        unique_id = unique_id_values[r] if r < len(unique_id_values) else ""
        new_row.append(unique_id)

        element_id_value = ""
        if 0 <= element_id_csv_index < len(row):
            element_id_value = row[element_id_csv_index]
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

    file_name = "{}_{}.xlsx".format(
        sanitize_filename(get_schedule_name(schedule)),
        str(int(time.time() * 1000))
    )
    full_path = os.path.join(TEMP_FOLDER, file_name)

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

        header_row = data_sheet.Range[data_sheet.Cells[2, 1], data_sheet.Cells[2, total_cols]]
        header_row.Font.Bold = True
        header_row.Font.Color = 0xFFFFFF
        header_row.Interior.Color = 0x5B7D95

        last_data_row = max(2, len(final_rows) + 2)

        for c in range(1, total_cols + 1):
            role = export_columns[c - 1].get("ColumnRole", "")
            hidden = export_columns[c - 1].get("Hidden", False)

            if hidden:
                data_sheet.Columns[c].Hidden = True

            if role in ("UniqueId", "ElementId"):
                fill_color = 0xEDEDED
            else:
                fill_color = get_status_fill_color(export_columns[c - 1].get("Status", "Unknown"))

            data_sheet.Range[
                data_sheet.Cells[2, c],
                data_sheet.Cells[last_data_row, c]
            ].Interior.Color = fill_color

        data_sheet.Columns[2].ColumnWidth = 12

        used_range = data_sheet.Range[data_sheet.Cells[2, 1], data_sheet.Cells[last_data_row, total_cols]]
        used_range.Borders.LineStyle = 1

        data_sheet.Rows["2:2"].AutoFilter()
        auto_fit_used_columns(data_sheet, total_cols)
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
        auto_fit_used_columns(schema_sheet, len(schema_headers))

        workbook.SaveAs(full_path)
        workbook.Close(True)
        excel_app.Quit()

        try:
            os.remove(csv_path)
        except Exception:
            pass

        return full_path, has_valid_element_ids, element_id_csv_index >= 0

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
        self.btnExport.Click += self.on_export_clicked
        self.btnClose.Click += self.on_close_clicked
        self.chkSelectAll.Click += self.on_select_all_clicked

    def _configure_export_state(self):
        self.btnExport.IsEnabled = True
        self.btnExport.ToolTip = "Export selected schedule to formatted Excel (.xlsx)"

    def _refresh_schedule_list(self):
        self.lstSchedules.ItemsSource = None
        self.lstSchedules.ItemsSource = self.filtered_schedules

    def _refresh_parameter_grid(self, parameter_items):
        self.dgParameters.ItemsSource = None
        self.dgParameters.ItemsSource = parameter_items

    def _update_status(self):
        checked_count = len([x for x in self.all_schedules if x.IsChecked])
        total_count = len(self.all_schedules)
        field_count = len(self.filtered_parameters)

        self.lblStatus.Text = "Schedules: {} | Checked: {} | Fields: {}".format(
            total_count,
            checked_count,
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

    def on_export_clicked(self, sender, args):
        selected_item = self._get_selected_item()
        if selected_item is None:
            TaskDialog.Show(__title__, "Select one schedule first.")
            return

        try:
            full_path, has_valid_ids, has_element_id_column = export_schedule_to_xlsx(selected_item.Schedule)
        except Exception as ex:
            TaskDialog.Show(__title__, "Excel export failed.\n\n{}".format(safe_text(ex)))
            return

        message = "Excel export completed.\n\n{}".format(full_path)

        if not has_element_id_column:
            message += "\n\nWarning:\nElementId could not be injected or detected."
            message += "\nUniqueId may be empty."

        elif not has_valid_ids:
            message += "\n\nWarning:\nElementId was found, but UniqueId could not be resolved from the rows."

        try:
            subprocess.Popen([full_path], shell=True)
        except Exception:
            pass

        TaskDialog.Show(__title__, message)

    def on_close_clicked(self, sender, args):
        self.Close()

    def on_select_all_clicked(self, sender, args):
        state = bool(self.chkSelectAll.IsChecked)

        for item in self.all_schedules:
            item.IsChecked = state

        self.lstSchedules.Items.Refresh()
        self._update_status()

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
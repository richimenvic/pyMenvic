# -*- coding: utf-8 -*-

import re

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction
from pyrevit import script

from importers import ImportResult


ENABLE_EXPERIMENTAL_SCHEDULE_IMPORT = True
DEBUG_OUTPUT = False
SCHEDULE_EXPERIMENTAL_CREATE_UNIQUE = False
SCHEDULE_USE_ELECTRICAL_PANEL_NORMALIZER = False
SCHEDULE_MIN_COL_WIDTH = 0.35
SCHEDULE_MAX_COL_WIDTH = 3.50
SCHEDULE_DEFAULT_COL_WIDTH = 0.80
SCHEDULE_EXCEL_WIDTH_SCALE = 0.055
SCHEDULE_MAX_TOTAL_WIDTH = 22.0
SCHEDULE_MIN_COL_WIDTH_FT = SCHEDULE_MIN_COL_WIDTH
SCHEDULE_DEFAULT_COL_WIDTH_FT = SCHEDULE_DEFAULT_COL_WIDTH
SCHEDULE_MAX_COL_WIDTH_FT = SCHEDULE_MAX_COL_WIDTH
SCHEDULE_ABSOLUTE_MIN_COL_WIDTH_FT = 0.35
SCHEDULE_MAX_TEXT_CHARS = 250
SCHEDULE_DEBUG_CELL_LIMIT = 25
SCHEDULE_DEBUG_SIMPLE_HEADER = False
SCHEDULE_DEBUG_VERBOSE = False
DEBUG_VERBOSE_SCHEDULE = False
SCHEDULE_SIMPLE_COL_WIDTH_FT = 1.0
SCHEDULE_DEBUG_LIMIT_REAL_IMPORT = False
SCHEDULE_DEBUG_LIMIT_ROWS = 10
SCHEDULE_DEBUG_LIMIT_COLS = 8
SCHEDULE_DEBUG_LIMIT_COL_WIDTH_FT = 2.50
SCHEDULE_RENDER_FIELD_WIDTH_FT = 0.30
SCHEDULE_DROP_VISUAL_SPACER_COLUMNS = False
SCHEDULE_CLIP_LONG_TEXT = False
SCHEDULE_MAX_CELL_CHARS = 28
SCHEDULE_DEBUG_APPEND_SIMPLE_MARKER = False
SCHEDULE_COL_WIDTH_MIN_FT = SCHEDULE_MIN_COL_WIDTH
SCHEDULE_COL_WIDTH_SHORT_FT = 0.65
SCHEDULE_COL_WIDTH_GENERAL_FT = 1.40
SCHEDULE_COL_WIDTH_DESCRIPTION_FT = 3.00
SCHEDULE_COL_WIDTH_MAX_FT = SCHEDULE_MAX_COL_WIDTH
SCHEDULE_NORMALIZE_ELECTRICAL_PANEL = SCHEDULE_USE_ELECTRICAL_PANEL_NORMALIZER
SCHEDULE_NORMALIZED_PANEL_HEADERS = [
    u"Breaker",
    u"Cable",
    u"Conduit",
    u"Circuit",
    u"Description",
    u"Phase 1",
    u"Phase 2",
    u"Phase 3",
    u"Amp",
    u"Pole",
    u"Duct",
    u"AWG",
]
SCHEDULE_NORMALIZED_PANEL_WIDTHS_FT = [0.8, 3.0, 1.0, 0.6, 3.5, 0.8, 0.8, 0.8, 0.7, 0.7, 0.7, 0.7]
SCHEDULE_NORMALIZED_PANEL_MAX_CHARS = 40


def _debug(message):
    if not DEBUG_OUTPUT:
        return
    text = _safe_unicode(message)
    try:
        output = script.get_output()
        output.print_md("Table Importer Schedule Debug: %s" % text)
    except Exception:
        try:
            print("Table Importer Schedule Debug: %s" % text)
        except Exception:
            pass


def _safe_unicode(value):
    if value is None:
        return u""
    try:
        if isinstance(value, unicode):
            return value
    except Exception:
        pass
    try:
        return unicode(value)
    except Exception:
        try:
            return unicode(value.ToString())
        except Exception:
            return u""


def _ctx(context, name, default=None):
    try:
        value = context.get(name)
        if value is not None:
            return value
    except Exception:
        pass
    return default


def _clean_text(context, value):
    cleaner = _ctx(context, "clean_display_text")
    try:
        if cleaner is not None:
            return cleaner(value)
    except Exception:
        pass
    return _safe_unicode(value).strip()


def _not_supported(reason):
    raise Exception("Schedule View header import is not supported in this Revit/API context. %s" % _safe_unicode(reason))


def _get_cell_value(table_data, row_index, col_index, context):
    helper = _ctx(context, "get_cell_value")
    if helper is not None:
        try:
            return _clean_text(context, helper(table_data, row_index, col_index))
        except Exception:
            pass
    try:
        row = table_data[row_index]
        if col_index < len(row):
            return _clean_text(context, row[col_index])
    except Exception:
        pass
    return u""


def _trim_cell_text(text):
    text = _safe_unicode(text)
    if SCHEDULE_CLIP_LONG_TEXT and len(text) > SCHEDULE_MAX_CELL_CHARS:
        return text[:SCHEDULE_MAX_CELL_CHARS - 3] + u"..."
    if len(text) > SCHEDULE_MAX_TEXT_CHARS:
        return text[:SCHEDULE_MAX_TEXT_CHARS - 3] + u"..."
    return text


def _trim_normalized_panel_text(text):
    text = _safe_unicode(text).strip()
    if len(text) > SCHEDULE_NORMALIZED_PANEL_MAX_CHARS:
        return text[:SCHEDULE_NORMALIZED_PANEL_MAX_CHARS - 3] + u"..."
    return text


def _get_table_data(entry, table_data, context):
    if table_data is not None:
        try:
            row_count = len(table_data)
        except Exception:
            row_count = 0
        column_count = 0
        for row in table_data or []:
            try:
                if len(row) > column_count:
                    column_count = len(row)
            except Exception:
                pass
        return table_data, row_count, column_count

    reader = _ctx(context, "read_table_data_for_entry")
    if reader is None:
        _not_supported("Excel table data reader is unavailable.")
    result = reader(entry)
    return result


def _debug_table_preview(entry, table_data, row_count, column_count, context):
    _debug("worksheet: %s" % _safe_unicode(getattr(entry, "Worksheet", "")))
    _debug("region: %s" % _safe_unicode(getattr(entry, "Region", "")))
    _debug("table_data rows: %s" % row_count)
    _debug("table_data columns: %s" % column_count)
    max_rows = min(5, int(row_count or 0))
    max_cols = min(8, int(column_count or 0))
    for row_index in range(max_rows):
        values = []
        for col_index in range(max_cols):
            values.append(_get_cell_value(table_data, row_index, col_index, context))
        _debug("excel row %s: %s" % (row_index, u" | ".join(values)))


def _count_empty_leading_columns(table_data, row_count, column_count, context):
    empty_leading = 0
    for col_index in range(int(column_count or 0)):
        has_value = False
        for row_index in range(int(row_count or 0)):
            if _get_cell_value(table_data, row_index, col_index, context).strip():
                has_value = True
                break
        if has_value:
            break
        empty_leading += 1
    return empty_leading


def _trim_schedule_leading_empty_columns(table_data, row_count, column_count, context):
    empty_leading = _count_empty_leading_columns(table_data, row_count, column_count, context)
    if empty_leading <= 0:
        _debug("Schedule leading empty columns trimmed: 0")
        return table_data, row_count, column_count
    if empty_leading >= column_count:
        _debug("Schedule leading empty columns trimmed: all columns were empty; keeping original data")
        return table_data, row_count, column_count

    trimmed = []
    for row in table_data:
        try:
            trimmed.append(list(row)[empty_leading:])
        except Exception:
            trimmed.append([])
    trimmed_columns = max(1, int(column_count) - empty_leading)
    _debug("Schedule leading empty columns trimmed: %s" % empty_leading)
    _debug("Schedule columns after trim: %s" % trimmed_columns)
    return trimmed, row_count, trimmed_columns


def _count_empty_trailing_columns(table_data, row_count, column_count, context):
    empty_trailing = 0
    for col_index in range(int(column_count or 0) - 1, -1, -1):
        has_value = False
        for row_index in range(int(row_count or 0)):
            if _get_cell_value(table_data, row_index, col_index, context).strip():
                has_value = True
                break
        if has_value:
            break
        empty_trailing += 1
    return empty_trailing


def _trim_schedule_trailing_empty_columns(table_data, row_count, column_count, context):
    empty_trailing = _count_empty_trailing_columns(table_data, row_count, column_count, context)
    if empty_trailing <= 0:
        _debug("Schedule trailing empty columns trimmed: 0")
        return table_data, row_count, column_count
    if empty_trailing >= column_count:
        _debug("Schedule trailing empty columns trimmed: all columns were empty; keeping original data")
        return table_data, row_count, column_count

    keep_count = int(column_count) - empty_trailing
    trimmed = []
    for row in table_data:
        try:
            trimmed.append(list(row)[:keep_count])
        except Exception:
            trimmed.append([])
    _debug("Schedule trailing empty columns trimmed: %s" % empty_trailing)
    _debug("Schedule columns after trailing trim: %s" % keep_count)
    return trimmed, row_count, keep_count


def _count_empty_leading_rows(table_data, row_count, column_count, context):
    empty_leading = 0
    for row_index in range(int(row_count or 0)):
        has_value = False
        for col_index in range(int(column_count or 0)):
            if _get_cell_value(table_data, row_index, col_index, context).strip():
                has_value = True
                break
        if has_value:
            break
        empty_leading += 1
    return empty_leading


def _trim_schedule_leading_empty_rows(table_data, row_count, column_count, context):
    empty_leading = _count_empty_leading_rows(table_data, row_count, column_count, context)
    if empty_leading <= 0:
        _debug("Schedule leading empty rows trimmed: 0")
        return table_data, row_count, column_count
    if empty_leading >= row_count:
        _debug("Schedule leading empty rows trimmed: all rows were empty; keeping original data")
        return table_data, row_count, column_count
    trimmed = list(table_data)[empty_leading:]
    trimmed_rows = max(1, int(row_count) - empty_leading)
    _debug("Schedule leading empty rows trimmed: %s" % empty_leading)
    _debug("Schedule rows after trim: %s" % trimmed_rows)
    return trimmed, trimmed_rows, column_count


def _count_empty_trailing_rows(table_data, row_count, column_count, context):
    empty_trailing = 0
    for row_index in range(int(row_count or 0) - 1, -1, -1):
        has_value = False
        for col_index in range(int(column_count or 0)):
            if _get_cell_value(table_data, row_index, col_index, context).strip():
                has_value = True
                break
        if has_value:
            break
        empty_trailing += 1
    return empty_trailing


def _trim_schedule_trailing_empty_rows(table_data, row_count, column_count, context):
    empty_trailing = _count_empty_trailing_rows(table_data, row_count, column_count, context)
    if empty_trailing <= 0:
        _debug("Schedule trailing empty rows trimmed: 0")
        return table_data, row_count, column_count
    if empty_trailing >= row_count:
        _debug("Schedule trailing empty rows trimmed: all rows were empty; keeping original data")
        return table_data, row_count, column_count
    keep_count = int(row_count) - empty_trailing
    trimmed = list(table_data)[:keep_count]
    _debug("Schedule trailing empty rows trimmed: %s" % empty_trailing)
    _debug("Schedule rows after trailing trim: %s" % keep_count)
    return trimmed, keep_count, column_count


def _trim_schedule_empty_edges(table_data, row_count, column_count, context):
    original_rows = row_count
    original_cols = column_count
    table_data, row_count, column_count = _trim_schedule_leading_empty_rows(table_data, row_count, column_count, context)
    table_data, row_count, column_count = _trim_schedule_trailing_empty_rows(table_data, row_count, column_count, context)
    table_data, row_count, column_count = _trim_schedule_leading_empty_columns(table_data, row_count, column_count, context)
    table_data, row_count, column_count = _trim_schedule_trailing_empty_columns(table_data, row_count, column_count, context)
    _debug("Schedule cleanup: original rows=%s columns=%s; cleaned rows=%s columns=%s" % (
        original_rows,
        original_cols,
        row_count,
        column_count,
    ))
    return table_data, row_count, column_count


def _get_simple_header_data():
    return [
        [u"A", u"B", u"C"],
        [u"1", u"2", u"3"],
        [u"TEST", u"TEST", u"TEST"],
    ], 3, 3


def _is_circuit_code(value):
    return re.match(r"^c\d+[\w\-]*$", _safe_unicode(value).strip().lower()) is not None


def _is_numeric_power_value(value):
    text = _safe_unicode(value).strip()
    if not text:
        return False
    text = text.replace(u",", u"")
    return re.match(r"^\d+(\.\d+)?$", text) is not None


def _looks_like_breaker(value):
    text = _safe_unicode(value).strip().lower()
    if not text:
        return False
    return re.search(r"\b\d+\s*p\b", text) is not None and re.search(r"\b\d+(\.\d+)?\s*a\b", text) is not None


def _looks_like_cable(value):
    text = _safe_unicode(value).strip().lower()
    return "awg" in text or " thw" in text or " cu" in text or " pe " in (" " + text + " ")


def _looks_like_conduit(value):
    text = _safe_unicode(value).strip().lower()
    return "emt" in text or "pvc" in text or u"ø" in text or u"Ø" in text or '"' in text


def _looks_like_pole(value):
    return re.search(r"\b\d+\s*p\b", _safe_unicode(value).strip().lower()) is not None


def _looks_like_amp(value):
    text = _safe_unicode(value).strip().lower()
    return re.search(r"\b\d+(\.\d+)?\s*a\b", text) is not None or re.match(r"^\d+(\.\d+)?$", text) is not None


def _split_breaker_parts(value):
    text = _safe_unicode(value).strip()
    pole = u""
    amp = u""
    pole_match = re.search(r"\b(\d+\s*p)\b", text, re.IGNORECASE)
    if pole_match:
        pole = pole_match.group(1).replace(u" ", u"").upper()
    amp_match = re.search(r"\b(\d+(?:\.\d+)?\s*a)\b", text, re.IGNORECASE)
    if amp_match:
        amp = re.sub(r"\s+", u" ", amp_match.group(1)).upper()
    return pole, amp


def _normalize_excel_panel_row(values):
    circuit_index = None
    circuit_value = u""
    for index, value in enumerate(values):
        if _is_circuit_code(value):
            circuit_index = index
            circuit_value = value
            break
    if circuit_index is None:
        return None

    before = values[:circuit_index]
    after = values[circuit_index + 1:]
    breaker = u""
    cable = u""
    conduit = u""
    for value in before:
        if not breaker and _looks_like_breaker(value):
            breaker = value
        elif not cable and _looks_like_cable(value):
            cable = value
        elif not conduit and _looks_like_conduit(value):
            conduit = value

    phase_start = None
    for index, value in enumerate(after):
        if _is_numeric_power_value(value):
            phase_start = index
            break
    if phase_start is None:
        description_parts = after
        phase_values = []
        remainder = []
    else:
        description_parts = after[:phase_start]
        phase_values = []
        remainder_start = phase_start
        for index in range(phase_start, len(after)):
            if len(phase_values) < 3 and _is_numeric_power_value(after[index]):
                phase_values.append(after[index])
                remainder_start = index + 1
            else:
                break
        remainder = after[remainder_start:]

    description = u" ".join([_safe_unicode(value).strip() for value in description_parts if _safe_unicode(value).strip()])
    amp = u""
    pole = u""
    duct = u""
    awg = u""
    for value in remainder:
        if _looks_like_breaker(value):
            breaker_pole, breaker_amp = _split_breaker_parts(value)
            if not pole and breaker_pole:
                pole = breaker_pole
            if not amp and breaker_amp:
                amp = breaker_amp
        elif not awg and _looks_like_cable(value):
            awg = value
        elif not duct and _looks_like_conduit(value):
            duct = value
        elif not pole and _looks_like_pole(value):
            pole = value
        elif not amp and _looks_like_amp(value):
            amp = value

    while len(phase_values) < 3:
        phase_values.append(u"")

    return [
        breaker,
        cable,
        conduit,
        circuit_value,
        description,
        phase_values[0],
        phase_values[1],
        phase_values[2],
        amp,
        pole,
        duct,
        awg,
    ]


def _normalize_electrical_panel_schedule(table_data, row_count, column_count, context):
    normalized = [list(SCHEDULE_NORMALIZED_PANEL_HEADERS)]
    for row_index in range(int(row_count or 0)):
        values = []
        for col_index in range(int(column_count or 0)):
            value = _get_cell_value(table_data, row_index, col_index, context)
            if value.strip():
                values.append(value.strip())
        if not values:
            continue
        panel_row = _normalize_excel_panel_row(values)
        if panel_row is None:
            continue
        normalized.append([_trim_normalized_panel_text(value) for value in panel_row])

    if len(normalized) <= 1:
        _debug("Schedule normalized panel import found no C# circuit rows; using current visual header logic")
        return None, 0, 0

    _debug("Schedule normalized electrical panel rows: %s" % (len(normalized) - 1))
    return normalized, len(normalized), len(SCHEDULE_NORMALIZED_PANEL_HEADERS)


def _is_normalized_panel_data(table_data, column_count, context):
    if int(column_count or 0) != len(SCHEDULE_NORMALIZED_PANEL_HEADERS):
        return False
    try:
        for col_index, header in enumerate(SCHEDULE_NORMALIZED_PANEL_HEADERS):
            if _get_cell_value(table_data, 0, col_index, context).strip().lower() != _safe_unicode(header).lower():
                return False
        return True
    except Exception:
        return False


def _is_key_schedule_text(value):
    text = _safe_unicode(value).strip()
    if not text:
        return False
    lowered = text.lower()
    if len(text) > 8:
        return True
    if re.match(r"^c\d+[\w\-]*$", lowered):
        return True
    if re.search(r"\d", text) and re.search(r"(w|kw|va|kva|a|awg|mm|pvc|emt|tfm|cable|cond|prote|breaker|term)", lowered):
        return True
    if re.search(r"[a-zA-Z]", text) and len(text) >= 3:
        return True
    return False


def _column_non_empty_values(table_data, row_count, col_index, context):
    values = []
    for row_index in range(int(row_count or 0)):
        value = _get_cell_value(table_data, row_index, col_index, context)
        if value.strip():
            values.append(value)
    return values


def _drop_schedule_spacer_columns(table_data, row_count, column_count, context):
    if not SCHEDULE_DROP_VISUAL_SPACER_COLUMNS:
        return table_data, row_count, column_count
    keep_indexes = []
    dropped_indexes = []
    for col_index in range(int(column_count or 0)):
        values = _column_non_empty_values(table_data, row_count, col_index, context)
        non_empty_count = len(values)
        keep = True
        if non_empty_count <= 0:
            keep = False
        elif non_empty_count < 2:
            keep = False
            for value in values:
                if _is_key_schedule_text(value):
                    keep = True
                    break
        elif non_empty_count <= max(1, int(row_count or 0) / 5):
            keep = False
            for value in values:
                if _is_key_schedule_text(value):
                    keep = True
                    break
        if keep:
            keep_indexes.append(col_index)
        else:
            dropped_indexes.append(col_index)

    if not keep_indexes:
        _debug("Schedule spacer column cleanup would drop all columns; keeping original data")
        return table_data, row_count, column_count

    cleaned = []
    for row_index in range(int(row_count or 0)):
        row_values = []
        for col_index in keep_indexes:
            row_values.append(_get_cell_value(table_data, row_index, col_index, context))
        cleaned.append(row_values)
    _debug("Schedule spacer columns dropped: %s" % ", ".join([_safe_unicode(index) for index in dropped_indexes]))
    _debug("Schedule useful columns kept: %s" % ", ".join([_safe_unicode(index) for index in keep_indexes]))
    return cleaned, row_count, len(keep_indexes)


def _limit_schedule_real_data(table_data, row_count, column_count, context):
    if not SCHEDULE_DEBUG_LIMIT_REAL_IMPORT:
        return table_data, row_count, column_count
    limited_rows = min(int(row_count or 0), int(SCHEDULE_DEBUG_LIMIT_ROWS))
    limited_cols = min(int(column_count or 0), int(SCHEDULE_DEBUG_LIMIT_COLS))
    limited = []
    for row_index in range(limited_rows):
        row_values = []
        for col_index in range(limited_cols):
            row_values.append(_trim_cell_text(_get_cell_value(table_data, row_index, col_index, context)))
        limited.append(row_values)
    if SCHEDULE_DEBUG_APPEND_SIMPLE_MARKER and limited_rows > 0 and limited_cols > 0:
        try:
            if not _safe_unicode(limited[0][0]).strip():
                limited[0][0] = u"VISIBLE"
                _debug("Schedule visible marker written at row=0 col=0")
        except Exception:
            pass
    _debug("Schedule limited real import enabled: rows=%s cols=%s width=%.2f" % (
        limited_rows,
        limited_cols,
        float(SCHEDULE_DEBUG_LIMIT_COL_WIDTH_FT),
    ))
    return limited, limited_rows, limited_cols


def _get_element_id_value(element_id, context):
    helper = _ctx(context, "get_element_id_value")
    if helper is not None:
        try:
            return helper(element_id)
        except Exception:
            pass
    try:
        return element_id.Value
    except Exception:
        pass
    try:
        return element_id.IntegerValue
    except Exception:
        pass
    try:
        return int(str(element_id))
    except Exception:
        return None


def _make_element_id(value, context):
    helper = _ctx(context, "make_element_id")
    if helper is not None:
        try:
            return helper(value)
        except Exception:
            return None
    return None


def _sanitize_name(context, value):
    helper = _ctx(context, "sanitize_revit_view_name")
    if helper is not None:
        try:
            return helper(value)
        except Exception:
            pass
    name = _clean_text(context, value)
    if not name:
        name = u"Imported Excel Schedule"
    return name


def clean_schedule_base_name(name):
    return re.sub(r"(_SCH_\d+)+$", "", _safe_unicode(name)).strip()


def _make_unique_name(context, base_name, existing_names):
    helper = _ctx(context, "make_unique_name")
    if helper is not None:
        try:
            return helper(base_name, existing_names)
        except Exception:
            pass
    existing = {}
    for name in existing_names or []:
        existing[_safe_unicode(name).lower()] = True
    if base_name.lower() not in existing:
        return base_name
    index = 2
    while True:
        candidate = u"%s %s" % (base_name, index)
        if candidate.lower() not in existing:
            return candidate
        index += 1


def _make_unique_schedule_test_name(context, base_name, existing_names):
    base = clean_schedule_base_name(_sanitize_name(context, base_name))
    if not base:
        base = u"Imported Excel Schedule"
    if base.upper().endswith("_SCH"):
        root = base
    else:
        root = u"%s_SCH" % base
    existing = {}
    for name in existing_names or []:
        existing[_safe_unicode(name).strip().lower()] = True
    index = 1
    while True:
        candidate = u"%s_%03d" % (root, index)
        if candidate.lower() not in existing:
            return candidate
        index += 1


def _get_existing_view_names(context, doc, exclude_view=None):
    names = []
    exclude_id = None
    if exclude_view is not None:
        try:
            exclude_id = _get_element_id_value(exclude_view.Id, context)
        except Exception:
            exclude_id = None
    try:
        for view in FilteredElementCollector(doc).OfClass(DB.View):
            try:
                if view.IsTemplate:
                    continue
                if exclude_id is not None and _get_element_id_value(view.Id, context) == exclude_id:
                    continue
                names.append(_clean_text(context, view.Name))
            except Exception:
                pass
    except Exception:
        helper = _ctx(context, "get_existing_revit_view_names")
        if helper is not None:
            try:
                names = helper(doc)
            except Exception:
                names = []
    return names


def _is_schedule_view(view):
    try:
        if isinstance(view, DB.ViewSchedule):
            return True
    except Exception:
        pass
    try:
        return view is not None and view.ViewType == DB.ViewType.Schedule
    except Exception:
        pass
    return False


def _get_schedule_category_ids(doc):
    category_items = []
    candidates = []
    for name in ("OST_GenericModel", "OST_Furniture", "OST_Doors", "OST_Walls", "OST_Rooms"):
        try:
            candidates.append(getattr(DB.BuiltInCategory, name))
        except Exception:
            pass
    for built_in_category in candidates:
        try:
            category = doc.Settings.Categories.get_Item(built_in_category)
            if category is not None and category.Id is not None:
                category_items.append((category.Id, _safe_unicode(category.Name)))
        except Exception:
            pass
    try:
        category_items.append((DB.ElementId.InvalidElementId, "InvalidElementId"))
    except Exception:
        pass
    return category_items


def _rename_schedule(schedule, entry, doc, context):
    desired_name = _sanitize_name(context, getattr(entry, "ViewName", ""))
    existing_names = _get_existing_view_names(context, doc, schedule)
    desired_name = _make_unique_name(context, desired_name, existing_names)
    try:
        if _clean_text(context, schedule.Name).lower() != desired_name.lower():
            schedule.Name = desired_name
    except Exception:
        schedule.Name = desired_name
    try:
        entry.ViewName = _clean_text(context, schedule.Name)
    except Exception:
        entry.ViewName = desired_name


def _create_schedule(doc, entry, context):
    desired_name = _sanitize_name(context, getattr(entry, "ViewName", ""))
    existing_names = _get_existing_view_names(context, doc, None)
    if SCHEDULE_EXPERIMENTAL_CREATE_UNIQUE:
        desired_name = _make_unique_schedule_test_name(context, desired_name, existing_names)
    else:
        desired_name = _make_unique_name(context, desired_name, existing_names)
    last_error = None
    for category_id, category_name in _get_schedule_category_ids(doc):
        try:
            _debug("schedule creation category id=%s name=%s" % (_safe_unicode(_get_element_id_value(category_id, context)), _safe_unicode(category_name)))
            schedule = DB.ViewSchedule.CreateSchedule(doc, category_id)
            try:
                schedule._table_importer_category_name = category_name
            except Exception:
                pass
            schedule.Name = desired_name
            try:
                entry.ViewName = _clean_text(context, schedule.Name)
            except Exception:
                entry.ViewName = desired_name
            schedule_id_value = _get_element_id_value(schedule.Id, context)
            if schedule_id_value is not None:
                entry.RevitViewId = _safe_unicode(schedule_id_value)
            return schedule
        except Exception as ex:
            last_error = ex
    _not_supported("Could not create a ViewSchedule: %s" % _safe_unicode(last_error))


def _get_or_create_schedule(doc, entry, context):
    if SCHEDULE_EXPERIMENTAL_CREATE_UNIQUE:
        _debug("SCHEDULE_EXPERIMENTAL_CREATE_UNIQUE enabled: creating a new schedule for this run")
        return _create_schedule(doc, entry, context), True

    if getattr(entry, "RevitViewId", None):
        element_id = _make_element_id(entry.RevitViewId, context)
        if element_id is None:
            raise Exception("Invalid RevitViewId '%s'." % _safe_unicode(entry.RevitViewId))
        schedule = doc.GetElement(element_id)
        if schedule is None:
            raise Exception("Missing Revit schedule for RevitViewId '%s'." % _safe_unicode(entry.RevitViewId))
        if not _is_schedule_view(schedule):
            raise Exception("Existing view is not a Schedule View.")
        _rename_schedule(schedule, entry, doc, context)
        return schedule, False
    return _create_schedule(doc, entry, context), True


def _section_count(section, count_name, first_name, last_name):
    try:
        return int(getattr(section, count_name))
    except Exception:
        pass
    try:
        return max(0, int(getattr(section, last_name)) - int(getattr(section, first_name)) + 1)
    except Exception:
        return 0


def _section_index(section, name, fallback):
    try:
        return int(getattr(section, name))
    except Exception:
        return fallback


def _row_count(section):
    return _section_count(section, "NumberOfRows", "FirstRowNumber", "LastRowNumber")


def _col_count(section):
    return _section_count(section, "NumberOfColumns", "FirstColumnNumber", "LastColumnNumber")


def _get_header_section(schedule):
    try:
        schedule_table_data = schedule.GetTableData()
        section = schedule_table_data.GetSectionData(DB.SectionType.Header)
        _debug("using schedule section: Header")
        return section
    except Exception as ex:
        _not_supported("Could not access schedule header section: %s" % _safe_unicode(ex))


def _insert_row(section):
    before = _row_count(section)
    insert_index = _section_index(section, "LastRowNumber", before - 1) + 1
    if DEBUG_VERBOSE_SCHEDULE:
        _debug("insert header row attempt: before=%s insert_index=%s" % (before, insert_index))
    try:
        section.InsertRow(insert_index)
        if DEBUG_VERBOSE_SCHEDULE:
            _debug("insert header row succeeded with index %s" % insert_index)
    except Exception:
        try:
            section.InsertRow(before)
            if DEBUG_VERBOSE_SCHEDULE:
                _debug("insert header row succeeded with fallback index %s" % before)
        except Exception as ex:
            _debug("insert header row failed: %s" % _safe_unicode(ex))
            _not_supported("Could not insert schedule header row: %s" % _safe_unicode(ex))
    if _row_count(section) <= before:
        _not_supported("Schedule header row count did not increase.")


def _insert_column(section):
    before = _col_count(section)
    insert_index = _section_index(section, "LastColumnNumber", before - 1) + 1
    if DEBUG_VERBOSE_SCHEDULE:
        _debug("insert header column attempt: before=%s insert_index=%s" % (before, insert_index))
    try:
        section.InsertColumn(insert_index)
        if DEBUG_VERBOSE_SCHEDULE:
            _debug("insert header column succeeded with index %s" % insert_index)
    except Exception:
        try:
            section.InsertColumn(before)
            if DEBUG_VERBOSE_SCHEDULE:
                _debug("insert header column succeeded with fallback index %s" % before)
        except Exception as ex:
            _debug("insert header column failed: %s" % _safe_unicode(ex))
            _not_supported("Could not insert schedule header column: %s" % _safe_unicode(ex))
    if _col_count(section) <= before:
        _not_supported("Schedule header column count did not increase.")


def _ensure_header_size(section, target_rows, target_cols):
    target_rows = max(1, int(target_rows or 1))
    target_cols = max(1, int(target_cols or 1))
    _debug("header rows before insertion: %s" % _row_count(section))
    _debug("header columns before insertion: %s" % _col_count(section))
    guard = 0
    while _row_count(section) < target_rows:
        guard += 1
        if guard > target_rows + 20:
            _not_supported("Could not size schedule header rows.")
        _insert_row(section)
    guard = 0
    while _col_count(section) < target_cols:
        guard += 1
        if guard > target_cols + 20:
            _not_supported("Could not size schedule header columns.")
        _insert_column(section)
    _debug("header rows after insertion: %s" % _row_count(section))
    _debug("header columns after insertion: %s" % _col_count(section))


def _set_column_width(section, column_index, width):
    if width < SCHEDULE_MIN_COL_WIDTH:
        width = SCHEDULE_MIN_COL_WIDTH
    if width > SCHEDULE_MAX_COL_WIDTH:
        width = SCHEDULE_MAX_COL_WIDTH
    before_width = None
    try:
        before_width = section.GetColumnWidth(column_index)
    except Exception as width_ex:
        _debug("column %s current width unavailable: %s" % (column_index, _safe_unicode(width_ex)))
    _debug("column width attempt: column=%s current=%s target=%.3f" % (column_index, _safe_unicode(before_width), float(width)))
    try:
        section.SetColumnWidth(column_index, width)
        try:
            after_width = section.GetColumnWidth(column_index)
        except Exception:
            after_width = u"<unavailable>"
        _debug("column width after: column=%s width=%s" % (column_index, _safe_unicode(after_width)))
        return True
    except Exception as ex:
        _debug("SetColumnWidth failed for column %s target=%.3f: %s" % (column_index, float(width), _safe_unicode(ex)))
        pass
    try:
        section.SetColumnWidth(column_index, float(width))
        try:
            after_width = section.GetColumnWidth(column_index)
        except Exception:
            after_width = u"<unavailable>"
        _debug("column width after fallback: column=%s width=%s" % (column_index, _safe_unicode(after_width)))
        return True
    except Exception as ex:
        _debug("SetColumnWidth fallback failed for column %s: %s" % (column_index, _safe_unicode(ex)))
        return False


def _get_excel_column_widths(table_data, column_count):
    try:
        column_widths = getattr(table_data, "column_widths", None)
    except Exception:
        column_widths = None
    if not column_widths:
        return None
    widths = []
    has_width = False
    for col_index in range(int(column_count or 0)):
        width = None
        try:
            width = column_widths[col_index]
        except Exception:
            width = None
        if width is None:
            widths.append(SCHEDULE_DEFAULT_COL_WIDTH)
            continue
        try:
            width = float(width) * SCHEDULE_EXCEL_WIDTH_SCALE
            has_width = True
        except Exception:
            width = SCHEDULE_DEFAULT_COL_WIDTH
        if width < SCHEDULE_MIN_COL_WIDTH:
            width = SCHEDULE_MIN_COL_WIDTH
        if width > SCHEDULE_MAX_COL_WIDTH:
            width = SCHEDULE_MAX_COL_WIDTH
        widths.append(width)
    if not has_width:
        return None
    _debug("Schedule Excel column widths used: %s" % ", ".join(["%.2f" % width for width in widths]))
    return widths


def _get_schedule_column_widths(table_data, row_count, column_count, context):
    if SCHEDULE_USE_ELECTRICAL_PANEL_NORMALIZER and _is_normalized_panel_data(table_data, column_count, context):
        _debug("Schedule normalized panel fixed widths: %s" % ", ".join(["%.2f" % width for width in SCHEDULE_NORMALIZED_PANEL_WIDTHS_FT]))
        return _limit_schedule_total_width(list(SCHEDULE_NORMALIZED_PANEL_WIDTHS_FT))
    excel_widths = _get_excel_column_widths(table_data, column_count)
    if excel_widths:
        return _limit_schedule_total_width(excel_widths)
    widths = []
    for col_index in range(int(column_count or 0)):
        total_len = 0
        non_empty = 0
        max_len = 0
        for row_index in range(int(row_count or 0)):
            value = _get_cell_value(table_data, row_index, col_index, context)
            if value.strip():
                length = len(value)
                total_len += length
                non_empty += 1
                if length > max_len:
                    max_len = length
        avg_len = 0
        if non_empty:
            avg_len = float(total_len) / float(non_empty)
        if non_empty <= 0:
            width = SCHEDULE_DEFAULT_COL_WIDTH
        elif avg_len > 15 or max_len > 22:
            width = min(SCHEDULE_MAX_COL_WIDTH, max(SCHEDULE_DEFAULT_COL_WIDTH, float(max_len) * 0.055))
        elif avg_len <= 6 and max_len <= 10:
            width = max(SCHEDULE_MIN_COL_WIDTH, min(SCHEDULE_DEFAULT_COL_WIDTH, float(max_len + 2) * 0.07))
        else:
            width = max(SCHEDULE_DEFAULT_COL_WIDTH, float(max_len + 2) * 0.06)
        if width < SCHEDULE_MIN_COL_WIDTH:
            width = SCHEDULE_MIN_COL_WIDTH
        if width > SCHEDULE_MAX_COL_WIDTH:
            width = SCHEDULE_MAX_COL_WIDTH
        widths.append(width)
    _debug("Schedule final widths: %s" % ", ".join(["%.2f" % width for width in widths]))
    return _limit_schedule_total_width(widths)


def _limit_schedule_total_width(widths):
    if not widths:
        return widths
    total = 0.0
    for width in widths:
        try:
            total += float(width)
        except Exception:
            total += float(SCHEDULE_DEFAULT_COL_WIDTH)
    if total <= float(SCHEDULE_MAX_TOTAL_WIDTH):
        return widths

    try:
        scale = float(SCHEDULE_MAX_TOTAL_WIDTH) / total
    except Exception:
        scale = 1.0

    limited = []
    for width in widths:
        try:
            value = float(width) * scale
        except Exception:
            value = float(SCHEDULE_DEFAULT_COL_WIDTH) * scale
        if value < SCHEDULE_MIN_COL_WIDTH:
            value = SCHEDULE_MIN_COL_WIDTH
        if value > SCHEDULE_MAX_COL_WIDTH:
            value = SCHEDULE_MAX_COL_WIDTH
        limited.append(value)
    _debug("Schedule total width limited from %.2f to %.2f ft" % (total, sum(limited)))
    return limited


def _set_readable_column_widths(section, column_count, widths=None):
    first_col = _section_index(section, "FirstColumnNumber", 0)
    failures = 0
    target_width = SCHEDULE_DEFAULT_COL_WIDTH_FT
    if SCHEDULE_DEBUG_SIMPLE_HEADER:
        target_width = SCHEDULE_SIMPLE_COL_WIDTH_FT
    if SCHEDULE_DEBUG_LIMIT_REAL_IMPORT:
        target_width = SCHEDULE_DEBUG_LIMIT_COL_WIDTH_FT
    if widths:
        _debug("schedule header column width target=content-based for %s column(s)" % column_count)
    else:
        _debug("schedule header column width target=%.2f for %s column(s)" % (float(target_width), column_count))
    for offset in range(column_count):
        width = target_width
        if widths and offset < len(widths):
            width = widths[offset]
        if not _set_column_width(section, first_col + offset, width):
            failures += 1
    if failures >= column_count and column_count > 0:
        _not_supported("Could not set readable schedule header column widths.")
    _debug("schedule header column width attempts complete: %s success, %s failed" % (column_count - failures, failures))


def _set_cell_text(section, row_index, col_index, text, context, debug_label=None):
    text = _trim_cell_text(_clean_text(context, text))
    if debug_label and SCHEDULE_DEBUG_VERBOSE:
        _debug("%s cell write attempt row=%s col=%s value='%s'" % (debug_label, row_index, col_index, text))
    try:
        section.SetCellText(row_index, col_index, text)
        if debug_label and SCHEDULE_DEBUG_VERBOSE:
            _debug("%s cell write succeeded row=%s col=%s" % (debug_label, row_index, col_index))
        return True
    except Exception as ex:
        if debug_label:
            _debug("%s cell write failed row=%s col=%s: %s" % (debug_label, row_index, col_index, _safe_unicode(ex)))
        return False


def _get_cell_text(section, row_index, col_index):
    try:
        return True, section.GetCellText(row_index, col_index)
    except Exception as ex:
        return False, _safe_unicode(ex)


def _debug_readback(section, row_index, col_index, expected, label):
    ok, value = _get_cell_text(section, row_index, col_index)
    if ok:
        _debug("%s readback row=%s col=%s expected='%s' readback='%s'" % (
            label,
            row_index,
            col_index,
            _safe_unicode(expected),
            _safe_unicode(value),
        ))
        return _clean_text({}, value) == _safe_unicode(expected).strip()
    _debug("%s readback failed row=%s col=%s: %s" % (label, row_index, col_index, _safe_unicode(value)))
    return False


def _clear_visible_cells(section, row_count, column_count, context):
    first_row = _section_index(section, "FirstRowNumber", 0)
    first_col = _section_index(section, "FirstColumnNumber", 0)
    actual_rows = max(_row_count(section), row_count)
    actual_cols = max(_col_count(section), column_count)
    for row_offset in range(actual_rows):
        for col_offset in range(actual_cols):
            _set_cell_text(section, first_row + row_offset, first_col + col_offset, u"", context, None)


def _has_merged_cells(table_data):
    try:
        merged_ranges = getattr(table_data, "merged_ranges", None)
    except Exception:
        merged_ranges = None
    try:
        return bool(merged_ranges)
    except Exception:
        return False


def _try_set_bool_property(target, property_name, value):
    try:
        if hasattr(target, property_name):
            try:
                before_value = getattr(target, property_name)
            except Exception:
                before_value = u"<unavailable>"
            try:
                setattr(target, property_name, value)
                try:
                    after_value = getattr(target, property_name)
                except Exception:
                    after_value = u"<unavailable>"
                _debug("visibility setting %s: before=%s after=%s" % (property_name, _safe_unicode(before_value), _safe_unicode(after_value)))
                return True
            except Exception as set_ex:
                _debug("visibility setting %s found but could not be set: %s" % (property_name, _safe_unicode(set_ex)))
                return False
    except Exception:
        pass
    _debug("visibility setting %s: not found" % property_name)
    return False


def _try_enable_schedule_title_header_visibility(schedule):
    _debug("checking Schedule title/header visibility settings")
    targets = [schedule]
    try:
        definition = schedule.Definition
        targets.append(definition)
        _debug("ScheduleDefinition found for visibility inspection")
    except Exception as ex:
        _debug("ScheduleDefinition unavailable for visibility inspection: %s" % _safe_unicode(ex))
    for target in targets:
        for property_name in ("ShowTitle", "ShowHeaders", "IsTitleVisible", "IsHeaderVisible", "ShowHeader", "ShowTitleAndHeaders"):
            _try_set_bool_property(target, property_name, True)


def _definition_field_count(definition):
    try:
        return int(definition.GetFieldCount())
    except Exception as ex:
        _debug("definition field count unavailable: %s" % _safe_unicode(ex))
        return -1


def _get_schedulable_field_name(field):
    for method_name in ("GetName", "GetFieldName"):
        try:
            method = getattr(field, method_name)
            return _safe_unicode(method())
        except Exception:
            pass
    for property_name in ("Name", "FieldName"):
        try:
            return _safe_unicode(getattr(field, property_name))
        except Exception:
            pass
    return _safe_unicode(field)


def _is_preferred_schedulable_field(name):
    lowered = _safe_unicode(name).lower()
    for token in ("type name", "family and type", "comments", "mark", "count"):
        if token in lowered:
            return True
    return False


def _add_visible_schedule_field_if_needed(schedule):
    try:
        definition = schedule.Definition
    except Exception as ex:
        _debug("schedule definition unavailable for field setup: %s" % _safe_unicode(ex))
        return False

    before_count = _definition_field_count(definition)
    _debug("definition field count before: %s" % before_count)
    schedulable_fields = []
    try:
        raw_fields = definition.GetSchedulableFields()
        for field in raw_fields:
            schedulable_fields.append(field)
    except Exception as ex:
        _debug("GetSchedulableFields failed: %s" % _safe_unicode(ex))
    _debug("schedulable fields available count: %s" % len(schedulable_fields))

    if before_count > 0:
        _debug("definition already has visible/body field(s); no field added")
        try:
            existing_field = definition.GetField(0)
            _minimize_render_field(existing_field)
        except Exception as ex:
            _debug("existing render field could not be minimized: %s" % _safe_unicode(ex))
        return True

    selected = None
    selected_name = u""
    for field in schedulable_fields:
        name = _get_schedulable_field_name(field)
        if _is_preferred_schedulable_field(name):
            selected = field
            selected_name = name
            break
    if selected is None and schedulable_fields:
        selected = schedulable_fields[0]
        selected_name = _get_schedulable_field_name(selected)

    if selected is None:
        _debug("no schedulable field available to add")
        return False

    try:
        added_field = definition.AddField(selected)
        try:
            added_name = _safe_unicode(added_field.GetName())
        except Exception:
            added_name = selected_name
        _debug("added visible schedule field: %s" % _safe_unicode(added_name))
    except Exception as ex:
        _debug("AddField failed for '%s': %s" % (_safe_unicode(selected_name), _safe_unicode(ex)))
        return False

    after_count = _definition_field_count(definition)
    _debug("definition field count after: %s" % after_count)
    _minimize_render_field(added_field)
    return after_count > before_count


def _minimize_render_field(field):
    if field is None:
        return
    try:
        field.ColumnHeading = "."
        _debug("render field header set to '.'")
    except Exception as ex:
        _debug("render field header could not be renamed: %s" % _safe_unicode(ex))
    for property_name in ("GridColumnWidth", "SheetColumnWidth"):
        try:
            setattr(field, property_name, SCHEDULE_RENDER_FIELD_WIDTH_FT)
            try:
                after_value = getattr(field, property_name)
            except Exception:
                after_value = u"<unavailable>"
            _debug("render field %s set to %s" % (property_name, _safe_unicode(after_value)))
            return
        except Exception as ex:
            _debug("render field %s could not be set: %s" % (property_name, _safe_unicode(ex)))


def _activate_schedule_view(schedule, context):
    getter = _ctx(context, "get_revit_uidocument")
    uidoc = None
    if getter is not None:
        try:
            uidoc = getter()
        except Exception:
            uidoc = None
    if uidoc is None:
        _debug("ActiveUIDocument unavailable; schedule view not activated")
        return False
    try:
        uidoc.ActiveView = schedule
        _debug("activated schedule view: %s" % _safe_unicode(schedule.Name))
        return True
    except Exception as ex:
        _debug("could not activate schedule view '%s': %s" % (_safe_unicode(getattr(schedule, "Name", "")), _safe_unicode(ex)))
        return False


def _write_plain_text_header(doc, schedule, table_data, row_count, column_count, context):
    _try_enable_schedule_title_header_visibility(schedule)
    _debug("Merged cells are not supported yet in Schedule View import.")
    section = _get_header_section(schedule)
    _ensure_header_size(section, row_count, column_count)
    widths = _get_schedule_column_widths(table_data, row_count, column_count, context)
    _set_readable_column_widths(section, column_count, widths)

    first_row = _section_index(section, "FirstRowNumber", 0)
    first_col = _section_index(section, "FirstColumnNumber", 0)
    _debug("first valid header cell: row=%s col=%s" % (first_row, first_col))
    _clear_visible_cells(section, row_count, column_count, context)

    first_non_empty = None
    for test_row in range(row_count):
        for test_col in range(column_count):
            test_value = _get_cell_value(table_data, test_row, test_col, context)
            if test_value.strip():
                first_non_empty = (test_row, test_col, test_value)
                break
        if first_non_empty is not None:
            break
    if first_non_empty is None:
        _not_supported("Schedule was created, but no non-empty Excel values were found. See pyRevit output.")

    test_row, test_col, test_value = first_non_empty
    if not _set_cell_text(section, first_row + test_row, first_col + test_col, test_value, context, "test"):
        _not_supported("Schedule Header cell writing is not supported with current API method.")
    test_readback_ok = _debug_readback(section, first_row + test_row, first_col + test_col, test_value, "test")
    _set_cell_text(section, first_row + test_row, first_col + test_col, u"", context, None)

    written = 0
    non_empty_written = 0
    readback_success = 0
    readback_logged = 0
    for row_index in range(row_count):
        for col_index in range(column_count):
            text = _get_cell_value(table_data, row_index, col_index, context)
            debug_label = None
            if DEBUG_VERBOSE_SCHEDULE and (row_index * column_count + col_index) < SCHEDULE_DEBUG_CELL_LIMIT:
                debug_label = "excel"
            if _set_cell_text(section, first_row + row_index, first_col + col_index, text, context, debug_label):
                written += 1
                if _clean_text(context, text):
                    non_empty_written += 1
                    if readback_logged < 5:
                        if _debug_readback(section, first_row + row_index, first_col + col_index, text, "excel"):
                            readback_success += 1
                        readback_logged += 1
    if written <= 0 and row_count > 0 and column_count > 0:
        _not_supported("Could not write text into schedule header cells.")
    _debug("schedule header cell writes succeeded: %s" % written)
    _debug("non-empty Excel values written: %s" % non_empty_written)
    if non_empty_written <= 0:
        _not_supported("Schedule was created, but no header cells were written. See pyRevit output.")
    if not test_readback_ok and readback_success <= 0:
        _not_supported("Schedule header cell text was written, but readback did not confirm persistence. See pyRevit output.")
    try:
        doc.Regenerate()
        _debug("doc.Regenerate succeeded after schedule header write")
    except Exception as regen_ex:
        _debug("doc.Regenerate failed after schedule header write: %s" % _safe_unicode(regen_ex))
    return non_empty_written


def import_schedule_view(entry, table_data, doc, context):
    if not ENABLE_EXPERIMENTAL_SCHEDULE_IMPORT:
        return ImportResult("Skipped", "Schedule View import is experimental / not stable yet.", skipped=1)

    _debug("starting Schedule View import")
    table_data, row_count, column_count = _get_table_data(entry, table_data, context)
    _debug_table_preview(entry, table_data, row_count, column_count, context)
    if row_count <= 0 or column_count <= 0:
        return ImportResult("Skipped", "No readable Excel data.", skipped=1)
    original_rows = row_count
    original_cols = column_count
    _debug("Schedule original data rows=%s columns=%s" % (original_rows, original_cols))
    if SCHEDULE_DEBUG_SIMPLE_HEADER:
        _debug("SCHEDULE_DEBUG_SIMPLE_HEADER enabled: importing fixed 3x3 visible test data")
        table_data, row_count, column_count = _get_simple_header_data()
    else:
        if SCHEDULE_USE_ELECTRICAL_PANEL_NORMALIZER:
            normalized_data, normalized_rows, normalized_cols = _normalize_electrical_panel_schedule(table_data, row_count, column_count, context)
            if normalized_data is not None:
                table_data = normalized_data
                row_count = normalized_rows
                column_count = normalized_cols
            else:
                _debug("Electrical panel normalizer found no circuit rows; using generic selected range")
        else:
            _debug("Schedule generic mode: using selected Excel range without normalization")
    table_data, row_count, column_count = _trim_schedule_empty_edges(table_data, row_count, column_count, context)
    table_data, row_count, column_count = _trim_schedule_trailing_empty_columns(table_data, row_count, column_count, context)
    # TODO: entry.ViewScale is stored by the UI, but Schedule View scaling needs
    # a separate, verified rule. Keep current scale behavior unchanged for now.
    _debug("Schedule final import data rows=%s columns=%s" % (row_count, column_count))
    if row_count <= 0 or column_count <= 0:
        return ImportResult("Skipped", "No non-empty Excel data after Schedule cleanup.", skipped=1)

    transaction = Transaction(doc, "Import Table Importer Schedule View")
    transaction.Start()
    try:
        schedule, was_created = _get_or_create_schedule(doc, entry, context)
        try:
            _debug("schedule name: %s" % _safe_unicode(schedule.Name))
        except Exception:
            _debug("schedule name: <unavailable>")
        _add_visible_schedule_field_if_needed(schedule)
        non_empty_written = _write_plain_text_header(doc, schedule, table_data, row_count, column_count, context)
        if non_empty_written <= 0:
            _not_supported("Schedule was created, but no header cells were written. See pyRevit output.")
        schedule_id_value = _get_element_id_value(schedule.Id, context)
        if schedule_id_value is not None:
            entry.RevitViewId = _safe_unicode(schedule_id_value)
        try:
            entry.Status = "Updated"
        except Exception:
            pass
        transaction.Commit()
        transaction = None
        _activate_schedule_view(schedule, context)
        _debug("Schedule View import committed. Open schedule view: %s" % _safe_unicode(getattr(entry, "ViewName", "")))
        schedule_name = _safe_unicode(getattr(entry, "ViewName", ""))
        message = "Schedule View imported: %s" % schedule_name
        if _has_merged_cells(table_data):
            message = "%s. Merged cells are not fully supported in Schedule View." % message
        if was_created:
            return ImportResult("Updated", message, created=1)
        return ImportResult("Updated", message, updated=1)
    except Exception as ex:
        try:
            if transaction is not None:
                transaction.RollBack()
        except Exception:
            pass
        message = _safe_unicode(ex)
        _debug("Schedule import failed/skipped: %s" % message)
        return ImportResult("Skipped", message, skipped=1)

import json
import os


PROJECT_STANDARDS_DIAGNOSTIC_MODE = False
LOCKED_FILE_MESSAGE = (
    "Cannot save the Excel file because it is currently open or locked. "
    "Please close the file and try again, or choose another file name."
)


def _ctx(context, name):
    return context[name]


def _safe_text(context, value):
    return _ctx(context, "safe_text")(value)


def _normalize_text(context, value):
    return _ctx(context, "normalize_text")(value)


def _get_category_line_weight(context, category, graphics_style_type):
    if category is None:
        return ""

    try:
        return _safe_text(context, category.GetLineWeight(graphics_style_type))
    except Exception:
        return ""


def _build_plain_sheet(headers, data_rows):
    rows = [list(headers or [])]
    for data_row in data_rows or []:
        rows.append(list(data_row))
    return rows


def _build_metadata_cell(sheet_name, column_name, storage_type="String", is_read_only=True):
    metadata = {
        "UniqueId": "ProjectStandards.{0}.{1}".format(
            sheet_name.replace(" ", ""),
            column_name.replace(" ", "").replace("(", "").replace(")", "").replace(",", ""),
        ),
        "Name": column_name,
        "IsReadOnly": bool(is_read_only),
        "ExportType": sheet_name,
        "ParamStorageType": storage_type,
        "StartIndex": 2,
    }
    return json.dumps(metadata, ensure_ascii=False)


def _build_metadata_sheet(sheet_name, column_specs, data_rows):
    metadata_row = []
    description_row = []
    for column_spec in column_specs:
        column_name = column_spec.get("name", "")
        metadata_row.append(_build_metadata_cell(
            sheet_name,
            column_name,
            column_spec.get("storage_type", "String"),
            column_spec.get("is_read_only", True),
        ))
        description_row.append(column_spec.get("description", column_name))

    rows = [metadata_row, description_row]
    for data_row in data_rows or []:
        rows.append(list(data_row))
    return rows


def _get_integer_id_text(context, element_or_id):
    if element_or_id is None:
        return ""
    try:
        return _safe_text(context, element_or_id.IntegerValue)
    except Exception:
        pass
    try:
        return _safe_text(context, element_or_id.Id.IntegerValue)
    except Exception:
        return ""


def _get_color_label_spaced(context, color_value):
    if color_value is None:
        return ""
    try:
        return "{0}, {1}, {2}".format(color_value.Red, color_value.Green, color_value.Blue)
    except Exception:
        return _safe_text(context, color_value)


def _get_category_bucket_key(context, category):
    if category is None:
        return ""

    DB = _ctx(context, "DB")

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

    category_type_text = _normalize_text(context, getattr(category, "CategoryType", ""))
    if "annotation" in category_type_text:
        return "annotation"
    if "analytical" in category_type_text:
        return "analytical"
    if "model" in category_type_text:
        return "model"
    return ""


def _make_hierarchical_child_name(context, child_name):
    return "       |---- {0}".format(_safe_text(context, child_name))


def _collect_object_style_rows(context, bucket_key):
    DB = _ctx(context, "DB")
    doc = _ctx(context, "doc")
    rows = []

    try:
        categories = list(doc.Settings.Categories)
    except Exception:
        categories = []

    categories.sort(key=lambda item: _normalize_text(context, getattr(item, "Name", "")))

    for category in categories:
        if _get_category_bucket_key(context, category) != bucket_key:
            continue

        include_cut = bucket_key == "model"
        parent_row = [
            _get_integer_id_text(context, category),
            _safe_text(context, getattr(category, "Name", "")),
            _get_category_line_weight(context, category, DB.GraphicsStyleType.Projection),
        ]
        if include_cut:
            parent_row.append(_get_category_line_weight(context, category, DB.GraphicsStyleType.Cut))
        parent_row.append(_get_color_label_spaced(context, getattr(category, "LineColor", None)))
        rows.append(parent_row)

        try:
            subcategories = list(category.SubCategories)
        except Exception:
            subcategories = []

        subcategories.sort(key=lambda item: _normalize_text(context, getattr(item, "Name", "")))
        for subcategory in subcategories:
            child_row = [
                _get_integer_id_text(context, subcategory),
                _make_hierarchical_child_name(context, getattr(subcategory, "Name", "")),
                _get_category_line_weight(context, subcategory, DB.GraphicsStyleType.Projection),
            ]
            if include_cut:
                child_row.append(_get_category_line_weight(context, subcategory, DB.GraphicsStyleType.Cut))
            child_row.append(_get_color_label_spaced(context, getattr(subcategory, "LineColor", None)))
            rows.append(child_row)

    return rows


def _build_object_style_sheet_rows(context, bucket_key):
    include_cut = bucket_key == "model"
    if bucket_key == "model":
        sheet_name = "Model Objects"
    elif bucket_key == "annotation":
        sheet_name = "Annotation Objects"
    else:
        sheet_name = "Analytical Model Objects"

    column_specs = [
        {
            "name": "Id",
            "description": "Category Id\nString\nInstance",
        },
        {
            "name": "Name",
            "description": "Category Name\nString\nInstance",
        },
        {
            "name": "Projection",
            "description": "Line Weight\nProjection\nString\nInstance",
            "is_read_only": False,
        },
    ]
    if include_cut:
        column_specs.append({
            "name": "Cut",
            "description": "Line Weight\nCut\nString\nInstance",
            "is_read_only": False,
        })
    column_specs.append({
        "name": "Color",
        "description": "Line Color(R, G, B)\nString\nInstance",
        "is_read_only": False,
    })
    return _build_metadata_sheet(sheet_name, column_specs, _collect_object_style_rows(context, bucket_key))


def _build_project_information_rows(context):
    doc = _ctx(context, "doc")
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
            _safe_text(context, getattr(definition, "Name", "")),
            _ctx(context, "get_parameter_preview_value")(param),
            _ctx(context, "get_definition_group_label")(definition),
            _safe_text(context, getattr(param, "StorageType", "")),
            _ctx(context, "get_bool_label")(getattr(param, "IsReadOnly", True)),
        ])

    return rows


def _build_project_parameters_rows(context):
    doc = _ctx(context, "doc")
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
            shared_guid = _safe_text(context, guid_value)
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
            _safe_text(context, getattr(definition, "Name", "")),
            _ctx(context, "get_definition_type_label")(definition),
            _ctx(context, "get_definition_group_label")(definition),
            _ctx(context, "get_binding_type_label")(binding),
            _ctx(context, "get_binding_categories_label")(binding),
            _ctx(context, "get_bool_label")(shared_value),
            shared_guid,
            _ctx(context, "get_bool_label")(visible_value),
            _ctx(context, "get_bool_label")(user_modifiable_value),
        ])

    return rows


def _build_line_styles_rows(context):
    DB = _ctx(context, "DB")
    doc = _ctx(context, "doc")
    column_specs = [
        {
            "name": "Id",
            "description": "Category Id\nString\nInstance",
        },
        {
            "name": "Name",
            "description": "Category Name\nString\nInstance",
        },
        {
            "name": "Projection",
            "description": "Line Weight\nProjection\nString\nInstance",
            "is_read_only": False,
        },
        {
            "name": "Color",
            "description": "Line Color(R, G, B)\nString\nInstance",
            "is_read_only": False,
        },
        {
            "name": "Line Pattern",
            "description": "Line Pattern\nString\nInstance",
            "is_read_only": False,
        },
    ]
    data_rows = []

    try:
        lines_category = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
    except Exception:
        lines_category = None

    if lines_category is None:
        return _build_metadata_sheet("Line Styles", column_specs, data_rows)

    data_rows.append([
        _get_integer_id_text(context, lines_category),
        _safe_text(context, getattr(lines_category, "Name", "Lines")) or "Lines",
        _get_category_line_weight(context, lines_category, DB.GraphicsStyleType.Projection),
        _get_color_label_spaced(context, getattr(lines_category, "LineColor", None)),
        "",
    ])

    try:
        subcategories = list(lines_category.SubCategories)
    except Exception:
        subcategories = []

    subcategories.sort(key=lambda item: _normalize_text(context, getattr(item, "Name", "")))

    try:
        for subcategory in subcategories:
            data_rows.append([
                _get_integer_id_text(context, subcategory),
                _make_hierarchical_child_name(context, getattr(subcategory, "Name", "")),
                _get_category_line_weight(context, subcategory, DB.GraphicsStyleType.Projection),
                _get_color_label_spaced(context, getattr(subcategory, "LineColor", None)),
                _ctx(context, "get_line_pattern_name")(subcategory.GetLinePatternId(DB.GraphicsStyleType.Projection)),
            ])
    except Exception:
        pass

    return _build_metadata_sheet("Line Styles", column_specs, data_rows)


def _build_families_rows(context):
    rows = [[
        "Category",
        "Family",
        "Type",
    ]]

    last_category = None
    last_family = None
    for category_name, family_name, type_name in _collect_family_listing_rows(context):
        if category_name != last_category:
            rows.append([category_name, "", ""])
            last_category = category_name
            last_family = None

        if family_name != last_family:
            rows.append(["", family_name, ""])
            last_family = family_name

        if type_name:
            rows.append(["", "", type_name])

    return rows


def _collect_family_listing_rows(context):
    DB = _ctx(context, "DB")
    doc = _ctx(context, "doc")
    listing_rows = []

    try:
        families = list(DB.FilteredElementCollector(doc).OfClass(DB.Family))
    except Exception:
        families = []

    families.sort(key=lambda item: (
        _normalize_text(context, getattr(getattr(item, "FamilyCategory", None), "Name", "")),
        _normalize_text(context, getattr(item, "Name", "")),
    ))

    for family in families:
        category_name = _safe_text(context, getattr(getattr(family, "FamilyCategory", None), "Name", ""))
        family_name = _safe_text(context, getattr(family, "Name", ""))

        try:
            symbol_ids = list(family.GetFamilySymbolIds())
        except Exception:
            symbol_ids = []

        type_names = []
        for symbol_id in symbol_ids:
            try:
                symbol = doc.GetElement(symbol_id)
                type_names.append(_safe_text(context, getattr(symbol, "Name", "")))
            except Exception:
                pass

        type_names = [type_name for type_name in type_names if type_name]
        type_names.sort(key=lambda value: _normalize_text(context, value))

        for type_name in type_names:
            listing_rows.append([category_name, family_name, type_name])

    return listing_rows


def _build_family_listing_datasource_rows(context):
    rows = [["Category", "Family", "Type"]]
    rows.extend(_collect_family_listing_rows(context))
    return rows


def _build_param_values_rows():
    rows = []
    for index in range(1, 17):
        rows.append([str(index), str(index)])
    return rows


def _build_instructions_rows(context, selected_section_keys):
    selected_labels = [
        "Project Information",
        "Project Parameters",
        "Model Objects",
        "Annotation Objects",
        "Analytical Model Objects",
        "Line Styles",
        "Families",
    ]
    selected_export_labels = []
    for section_key in selected_section_keys:
        if section_key == "object_styles":
            selected_export_labels.extend([
                "Model Objects",
                "Annotation Objects",
                "Analytical Model Objects",
            ])
        else:
            selected_export_labels.append(_ctx(context, "get_project_standards_section_label")(section_key))

    return [
        ["SheetLink - pyMENVIC"],
        ["Project Standards Export"],
        [""],
        ["Workbook sheets:"],
        [", ".join(selected_labels)],
        [""],
        ["Selected sections:"],
        [", ".join(selected_export_labels) if selected_export_labels else "None"],
        [""],
        ["Notes:"],
        ["This workbook documents project standards exported from the current Revit model."],
        ["Project Parameters lists parameter definitions and their bindings, not per-element values."],
        ["Editability and import back into Revit are not part of this standards export workflow yet."],
    ]


def _get_rows_for_section(context, section_key):
    if section_key == "project_information":
        return _build_project_information_rows(context)
    if section_key == "project_parameters":
        return _build_project_parameters_rows(context)
    if section_key == "object_styles":
        return [["Info"], ["Objects export through split worksheets."]]
    if section_key == "line_styles":
        return _build_line_styles_rows(context)
    if section_key == "families":
        return _build_families_rows(context)
    return [["Info"], ["Section not available"]]


def _format_sheet(worksheet, row_count, col_count):
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


def _format_instructions_sheet(worksheet, row_count, col_count):
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


def _show_message(context, message):
    try:
        _ctx(context, "TaskDialog").Show(_ctx(context, "__title__"), message)
    except Exception:
        pass


def _can_overwrite_file(full_path):
    if not full_path or not os.path.exists(full_path):
        return True

    handle = None
    try:
        handle = open(full_path, "r+b")
        handle.seek(0, os.SEEK_END)
        return True
    except Exception:
        return False
    finally:
        try:
            if handle is not None:
                handle.close()
        except Exception:
            pass


def export_project_standards_to_xlsx(full_path, selected_section_keys, open_after_export, context):
    if not full_path or not selected_section_keys:
        return False

    if not _can_overwrite_file(full_path):
        _show_message(context, LOCKED_FILE_MESSAGE)
        return False

    Excel = _ctx(context, "Excel")
    excel_app = None
    workbooks = None
    workbook = None
    worksheets_to_release = []

    try:
        excel_app = Excel.ApplicationClass()
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        workbooks = excel_app.Workbooks
        workbook = workbooks.Add()

        sheet_entries = []
        for section_key in selected_section_keys:
            if section_key == "object_styles":
                sheet_entries.append({
                    "name": "Model Objects",
                    "rows": _build_object_style_sheet_rows(context, "model"),
                    "hidden": False,
                })
                sheet_entries.append({
                    "name": "Annotation Objects",
                    "rows": _build_object_style_sheet_rows(context, "annotation"),
                    "hidden": False,
                })
                sheet_entries.append({
                    "name": "Analytical Model Objects",
                    "rows": _build_object_style_sheet_rows(context, "analytical"),
                    "hidden": False,
                })
            else:
                sheet_entries.append({
                    "name": _ctx(context, "get_project_standards_section_label")(section_key),
                    "rows": _get_rows_for_section(context, section_key),
                    "hidden": False,
                })

        if not sheet_entries:
            return False

        sheet_entries.append({
            "name": "FamilyListingDataSource",
            "rows": _build_family_listing_datasource_rows(context),
            "hidden": True,
        })
        sheet_entries.append({
            "name": "ParamValues",
            "rows": _build_param_values_rows(),
            "hidden": True,
        })

        sheet_entries.append({
            "name": "Instructions",
            "rows": _build_instructions_rows(context, selected_section_keys),
            "hidden": False,
            "instructions": True,
        })

        total_sheet_count = len(sheet_entries)
        _ctx(context, "ensure_workbook_sheet_count")(workbook, total_sheet_count)

        try:
            while workbook.Worksheets.Count > total_sheet_count:
                workbook.Worksheets[workbook.Worksheets.Count].Delete()
        except Exception:
            pass

        first_data_sheet = workbook.Worksheets[1]
        first_data_sheet.Name = _ctx(context, "make_excel_sheet_name")(sheet_entries[0]["name"], "Data")
        worksheets_to_release.append(first_data_sheet)

        data_rows = sheet_entries[0]["rows"]
        _ctx(context, "write_matrix_to_range")(first_data_sheet, 1, 1, data_rows)
        _format_sheet(first_data_sheet, len(data_rows), len(data_rows[0]) if data_rows else 0)

        sheet_index = 2
        for sheet_entry in sheet_entries[1:]:
            worksheet = workbook.Worksheets[sheet_index]
            worksheet.Name = _ctx(context, "make_excel_sheet_name")(sheet_entry["name"], "Sheet{}".format(sheet_index))
            worksheets_to_release.append(worksheet)

            rows = sheet_entry["rows"]
            _ctx(context, "write_matrix_to_range")(worksheet, 1, 1, rows)
            if sheet_entry.get("instructions", False):
                _format_instructions_sheet(worksheet, len(rows), len(rows[0]) if rows else 0)
            else:
                _format_sheet(worksheet, len(rows), len(rows[0]) if rows else 0)

            try:
                if sheet_entry.get("hidden", False):
                    worksheet.Visible = 0
            except Exception:
                pass
            sheet_index += 1

        try:
            first_data_sheet.Activate()
        except Exception:
            pass

        try:
            if os.path.exists(full_path):
                os.remove(full_path)
        except Exception:
            pass

        workbook.SaveAs(full_path)
        workbook.Close(True)
        excel_app.Quit()

        if open_after_export:
            _ctx(context, "open_path_with_default_app")(full_path)

        return True
    finally:
        for worksheet in worksheets_to_release:
            _ctx(context, "release_com_object")(worksheet)
        _ctx(context, "release_com_object")(workbook)
        _ctx(context, "release_com_object")(workbooks)
        _ctx(context, "release_com_object")(excel_app)

import os
import subprocess


def release_com_object(obj, marshal):
    try:
        if obj is not None:
            marshal.ReleaseComObject(obj)
    except Exception:
        pass


def make_excel_sheet_name(name, fallback_name, normalize_text):
    invalid_chars = ['\\', '/', '?', '*', '[', ']', ':']
    sheet_name = normalize_text(name)

    if not sheet_name:
        sheet_name = fallback_name

    for invalid_char in invalid_chars:
        sheet_name = sheet_name.replace(invalid_char, "_")

    sheet_name = sheet_name.strip()
    if not sheet_name:
        sheet_name = fallback_name

    return sheet_name[:31]


def write_matrix_to_range(worksheet, start_row, start_col, matrix, array_type, object_type):
    if worksheet is None or not matrix:
        return

    row_count = len(matrix)
    if row_count == 0:
        return

    col_count = len(matrix[0])
    if col_count == 0:
        return

    values = array_type.CreateInstance(object_type, row_count, col_count)

    for row_index in range(row_count):
        row_vals = matrix[row_index]
        for col_index in range(col_count):
            cell_value = None
            if col_index < len(row_vals):
                cell_value = row_vals[col_index]
            values[row_index, col_index] = cell_value

    target_range = worksheet.Range[
        worksheet.Cells[start_row, start_col],
        worksheet.Cells[start_row + row_count - 1, start_col + col_count - 1]
    ]
    target_range.Value2 = values


def ensure_workbook_sheet_count(workbook, required_count):
    if workbook is None:
        return

    try:
        worksheets = workbook.Worksheets
        while worksheets.Count < required_count:
            worksheets.Add()
    except Exception:
        pass


def open_path_with_default_app(full_path):
    try:
        os.startfile(full_path)
        return True
    except Exception:
        pass

    try:
        subprocess.Popen([full_path], shell=False)
        return True
    except Exception:
        return False

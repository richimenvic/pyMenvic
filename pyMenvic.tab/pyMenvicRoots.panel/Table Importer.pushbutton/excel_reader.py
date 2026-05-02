# -*- coding: utf-8 -*-

import os
import System
from System.Runtime.InteropServices import Marshal


def get_last_modified(file_path):
    """Return formatted last modified date."""
    if not file_path or not os.path.exists(file_path):
        return ""

    try:
        timestamp = os.path.getmtime(file_path)
        return System.DateTime.FromFileTimeUtc(
            long(timestamp * 10000000) + 116444736000000000
        ).ToLocalTime().ToString("yyyy-MM-dd HH:mm")
    except Exception:
        return ""


def get_file_name_without_extension(file_path):
    if not file_path:
        return ""

    try:
        return os.path.splitext(os.path.basename(file_path))[0]
    except Exception:
        return ""


def get_excel_worksheets(file_path):
    """
    Reads worksheet names from an Excel file using COM.
    Requires Microsoft Excel installed.
    """
    if not file_path or not os.path.exists(file_path):
        return []

    excel = None
    workbook = None
    worksheets = []

    try:
        excel_type = System.Type.GetTypeFromProgID("Excel.Application")
        excel = System.Activator.CreateInstance(excel_type)

        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(file_path)

        for i in range(1, workbook.Worksheets.Count + 1):
            sheet = workbook.Worksheets.Item[i]
            worksheets.append(sheet.Name)

        workbook.Close(False)
        excel.Quit()

        return worksheets

    except Exception:
        try:
            if workbook:
                workbook.Close(False)
        except Exception:
            pass

        try:
            if excel:
                excel.Quit()
        except Exception:
            pass

        return []

    finally:
        try:
            if workbook:
                Marshal.ReleaseComObject(workbook)
        except Exception:
            pass

        try:
            if excel:
                Marshal.ReleaseComObject(excel)
        except Exception:
            pass


def get_used_range_address(file_path, worksheet_name):
    """
    Returns the used range address for a worksheet.
    Example: A1:F25
    """
    if not file_path or not os.path.exists(file_path):
        return "Used Range"

    excel = None
    workbook = None

    try:
        excel_type = System.Type.GetTypeFromProgID("Excel.Application")
        excel = System.Activator.CreateInstance(excel_type)

        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(file_path)
        sheet = workbook.Worksheets.Item[worksheet_name]

        used_range = sheet.UsedRange
        address = used_range.Address
        address = address.replace("$", "")

        workbook.Close(False)
        excel.Quit()

        return address

    except Exception:
        try:
            if workbook:
                workbook.Close(False)
        except Exception:
            pass

        try:
            if excel:
                excel.Quit()
        except Exception:
            pass

        return "Used Range"

    finally:
        try:
            if workbook:
                Marshal.ReleaseComObject(workbook)
        except Exception:
            pass

        try:
            if excel:
                Marshal.ReleaseComObject(excel)
        except Exception:
            pass
# -*- coding: utf-8 -*-

from importers import ImportResult


def import_legend_view(entry, table_data, doc, context):
    callback = None
    try:
        callback = context.get("import_to_legend_view")
    except Exception:
        callback = None

    if callback is None:
        return ImportResult("Failed", "Legend View importer callback is unavailable.", failed=1)

    result = callback(entry)
    if result == "Created":
        return ImportResult("Updated", "Legend View created.", created=1)
    return ImportResult("Updated", "Legend View updated.", updated=1)

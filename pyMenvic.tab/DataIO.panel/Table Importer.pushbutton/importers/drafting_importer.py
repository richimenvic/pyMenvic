# -*- coding: utf-8 -*-

from importers import ImportResult


def import_drafting_view(entry, table_data, doc, context, cleanup_legacy=False):
    # The Drafting View engine is intentionally kept behind the existing
    # callback for this refactor so its geometry/output remains unchanged.
    callback = None
    try:
        callback = context.get("import_to_drafting_view")
    except Exception:
        callback = None

    if callback is None:
        return ImportResult("Failed", "Drafting View importer callback is unavailable.", failed=1)

    result = callback(entry, cleanup_legacy)
    if result == "Created":
        return ImportResult("Updated", "Drafting View created.", created=1)
    return ImportResult("Updated", "Drafting View updated.", updated=1)

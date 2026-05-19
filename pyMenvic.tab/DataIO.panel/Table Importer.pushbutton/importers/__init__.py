# -*- coding: utf-8 -*-


class ImportResult(object):
    def __init__(self, status, message="", created=0, updated=0, skipped=0, failed=0):
        self.status = status
        self.message = message
        self.created = int(created or 0)
        self.updated = int(updated or 0)
        self.skipped = int(skipped or 0)
        self.failed = int(failed or 0)


def import_row_to_revit(entry, view_type, table_data, doc, context, cleanup_legacy=False):
    if view_type == "Drafting View":
        import drafting_importer
        return drafting_importer.import_drafting_view(entry, table_data, doc, context, cleanup_legacy)

    if view_type == "Schedule View":
        import schedule_importer
        return schedule_importer.import_schedule_view(entry, table_data, doc, context)

    if view_type == "Legend View":
        import legend_importer
        return legend_importer.import_legend_view(entry, table_data, doc, context)

    return ImportResult("Skipped", "Unsupported view type: %s" % view_type, skipped=1)

# -*- coding: utf-8 -*-
__title__ = "Sheets by Rev"

import re
from pyrevit import revit, DB, script, forms


doc = revit.doc
output = script.get_output()
logger = script.get_logger()


class RevisionOption(object):
    def __init__(self, name, revision_id, sort_key):
        self.name = name
        self.revision_id = revision_id
        self.sort_key = sort_key

    def __repr__(self):
        return self.name


def natural_key(value):
    """Sort strings naturally: 2 < 10, A2 < A10."""
    if value is None:
        return []
    value = str(value)
    return [int(token) if token.isdigit() else token.lower()
            for token in re.split(r'(\d+)', value)]


def element_id_value(element_id):
    """
    Compatible ElementId integer value for Revit 2024-2026.
    Revit 2024+ may expose Value.
    Older versions use IntegerValue.
    """
    if element_id is None:
        return None

    try:
        return element_id.Value
    except Exception:
        pass

    try:
        return element_id.IntegerValue
    except Exception:
        pass

    return None


def same_element_id(id_a, id_b):
    return element_id_value(id_a) == element_id_value(id_b)


def safe_as_string(param, default=""):
    if not param:
        return default

    try:
        value = param.AsString()
        if value:
            return value
    except Exception:
        pass

    try:
        value = param.AsValueString()
        if value:
            return value
    except Exception:
        pass

    return default


def get_built_in_parameter(element, *parameter_names):
    for parameter_name in parameter_names:
        built_in_parameter = getattr(DB.BuiltInParameter, parameter_name, None)
        if built_in_parameter is None:
            continue

        try:
            param = element.get_Parameter(built_in_parameter)
            if param:
                return param
        except Exception:
            continue

    return None


def get_parameter_value(element, built_in_names=None, lookup_names=None, default=""):
    built_in_names = built_in_names or []
    lookup_names = lookup_names or []

    param = get_built_in_parameter(element, *built_in_names)
    if param:
        value = safe_as_string(param, default="")
        if value:
            return value

    for lookup_name in lookup_names:
        try:
            param = element.LookupParameter(lookup_name)
        except Exception:
            param = None

        if param:
            value = safe_as_string(param, default="")
            if value:
                return value

    return default


def get_revision_sequence_number(revision):
    """
    SequenceNumber is safe even when revision numbering is Per Sheet.
    """
    if not revision:
        return None

    try:
        seq = revision.SequenceNumber
        if seq is not None:
            return seq
    except Exception:
        pass

    value = get_parameter_value(
        revision,
        built_in_names=[
            'PROJECT_REVISION_SEQUENCE_NUM',
            'PROJECT_REVISION_REVISION_NUM'
        ],
        lookup_names=[
            'Sequence Number',
            'Revision Sequence',
            'Revision Number'
        ],
        default=""
    )

    try:
        return int(value)
    except Exception:
        return None


def get_revision_number(revision):
    """
    Do not access revision.RevisionNumber directly without try/except.
    It fails when revision numbering is set to Per Sheet.
    """
    if not revision:
        return ""

    try:
        number = revision.RevisionNumber
        if number:
            return str(number)
    except Exception:
        pass

    value = get_parameter_value(
        revision,
        built_in_names=['PROJECT_REVISION_REVISION_NUM'],
        lookup_names=['Revision Number'],
        default=""
    )

    if value:
        return value

    seq = get_revision_sequence_number(revision)
    if seq is not None:
        return "Seq {}".format(seq)

    return ""


def get_revision_sort_key(revision):
    seq = get_revision_sequence_number(revision)
    if seq is not None:
        return [seq]

    return natural_key(get_revision_number(revision))


def get_revision_date(revision):
    if not revision:
        return ""

    return get_parameter_value(
        revision,
        built_in_names=['PROJECT_REVISION_REVISION_DATE'],
        lookup_names=['Revision Date']
    )


def get_revision_description(revision):
    if not revision:
        return ""

    return get_parameter_value(
        revision,
        built_in_names=['PROJECT_REVISION_DESCRIPTION'],
        lookup_names=['Revision Description']
    )


def build_revision_label(revision):
    rev_number = get_revision_number(revision)
    rev_description = get_revision_description(revision)
    rev_date = get_revision_date(revision)

    parts = [part for part in [rev_number, rev_description, rev_date] if part]
    return " - ".join(parts) if parts else "<Unnamed Revision>"


def load_revisions():
    revisions = DB.FilteredElementCollector(doc) \
                  .OfCategory(DB.BuiltInCategory.OST_Revisions) \
                  .WhereElementIsNotElementType() \
                  .ToElements()

    revision_items = []

    for revision in revisions:
        revision_items.append(
            RevisionOption(
                build_revision_label(revision),
                revision.Id,
                get_revision_sort_key(revision)
            )
        )

    return sorted(
        revision_items,
        key=lambda item: (item.sort_key, item.name.lower())
    )


def select_revision(revision_items):
    if not revision_items:
        forms.alert('No revisions found in the current project.', exitscript=True)

    return forms.SelectFromList.show(
        revision_items,
        title='Select Revision',
        button_name='Select',
        name_attr='name',
        multiselect=False
    )


def build_sheet_lookup():
    all_sheets = DB.FilteredElementCollector(doc) \
                   .OfCategory(DB.BuiltInCategory.OST_Sheets) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

    lookup = {}

    for sheet in all_sheets:
        try:
            lookup[sheet.SheetNumber] = sheet
        except Exception:
            continue

    return lookup


def clean_sheet_name(sheet_name):
    if not sheet_name:
        return 'Unknown Sheet'

    if ' - ' in sheet_name:
        return sheet_name.split(' - ', 1)[-1].strip()

    return sheet_name.strip()


def get_sheet_info(revision_cloud, sheet_by_number):
    owner_view = doc.GetElement(revision_cloud.OwnerViewId)

    if not owner_view:
        return 'Unknown Sheet', 'Unknown Number', None

    if isinstance(owner_view, DB.ViewSheet):
        return clean_sheet_name(owner_view.Name), owner_view.SheetNumber, owner_view

    if isinstance(owner_view, DB.View):
        sheet_number = get_parameter_value(
            owner_view,
            built_in_names=['VIEWER_SHEET_NUMBER'],
            lookup_names=['Sheet Number']
        )

        if sheet_number:
            sheet = sheet_by_number.get(sheet_number)
            if sheet:
                return clean_sheet_name(sheet.Name), sheet.SheetNumber, owner_view

    return 'Unknown Sheet', 'Unknown Number', owner_view


def get_cloud_comment(revision_cloud):
    try:
        comment_param = revision_cloud.get_Parameter(
            DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS
        )
        return safe_as_string(comment_param)
    except Exception:
        return ""


def collect_clouds_by_sheet(selected_revision_id, sheet_by_number):
    revision_clouds = DB.FilteredElementCollector(doc) \
                        .OfCategory(DB.BuiltInCategory.OST_RevisionClouds) \
                        .WhereElementIsNotElementType() \
                        .ToElements()

    data_by_sheet = {}

    for revision_cloud in revision_clouds:
        try:
            cloud_revision_id = revision_cloud.RevisionId
        except Exception:
            continue

        if not same_element_id(cloud_revision_id, selected_revision_id):
            continue

        sheet_name, sheet_number, owner_view = get_sheet_info(
            revision_cloud,
            sheet_by_number
        )

        if sheet_number not in data_by_sheet:
            data_by_sheet[sheet_number] = {
                'name': sheet_name,
                'views': {}
            }

        view_name = owner_view.Name if owner_view else 'Unknown View'

        data_by_sheet[sheet_number]['views'].setdefault(
            view_name,
            []
        ).append(revision_cloud)

    return data_by_sheet


def print_report(data_by_sheet):
    total_clouds = 0

    sorted_sheets = sorted(
        data_by_sheet.items(),
        key=lambda item: natural_key(item[0])
    )

    for sheet_number, sheet_data in sorted_sheets:
        output.print_md(
            '### **Sheet:** {} - {}\n'.format(
                sheet_number,
                sheet_data['name']
            )
        )

        view_items = sorted(
            sheet_data['views'].items(),
            key=lambda item: natural_key(item[0])
        )

        for view_name, clouds in view_items:
            output.print_md('#### **View:** {}\n'.format(view_name))

            try:
                clouds.sort(key=lambda cloud: element_id_value(cloud.Id) or 0)
            except Exception as exc:
                logger.warning(
                    'Could not sort clouds in view "%s": %s',
                    view_name,
                    exc
                )

            rows = []
            headers = ['CloudId', 'Sheet', 'View', 'Comment']

            for revision_cloud in clouds:
                total_clouds += 1

                rows.append([
                    output.linkify([revision_cloud.Id]),
                    sheet_number,
                    view_name,
                    get_cloud_comment(revision_cloud)
                ])

            if rows:
                output.print_table(table_data=rows, columns=headers)
            else:
                output.print_md('_No revision clouds in this view._\n')

    output.print_md(
        '\n**SEARCH COMPLETED - {} revision clouds found.**'.format(
            total_clouds
        )
    )


def main():
    revision_items = load_revisions()

    selected_item = select_revision(revision_items)

    if not selected_item:
        script.exit()

    selected_revision_id = selected_item.revision_id

    sheet_by_number = build_sheet_lookup()

    data_by_sheet = collect_clouds_by_sheet(
        selected_revision_id,
        sheet_by_number
    )

    print_report(data_by_sheet)


if __name__ == '__main__':
    main()

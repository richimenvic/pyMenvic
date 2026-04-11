# -*- coding: utf-8 -*-
__title__ = "Sheets by Rev"

import re
from pyrevit import revit, DB, script, forms


doc = revit.doc
output = script.get_output()
logger = script.get_logger()


def natural_key(value):
    """Sort strings naturally: 2 < 10, A2 < A10."""
    if value is None:
        return []
    value = str(value)
    return [int(token) if token.isdigit() else token.lower()
            for token in re.split(r'(\d+)', value)]


def safe_as_string(param, default=""):
    if not param:
        return default
    value = param.AsString()
    return value if value is not None else default


def get_built_in_parameter(element, *parameter_names):
    for parameter_name in parameter_names:
        built_in_parameter = getattr(DB.BuiltInParameter, parameter_name, None)
        if built_in_parameter is None:
            continue

        param = element.get_Parameter(built_in_parameter)
        if param:
            return param

    return None


def get_parameter_value(element, built_in_names=None, lookup_names=None, default=""):
    built_in_names = built_in_names or []
    lookup_names = lookup_names or []

    param = get_built_in_parameter(element, *built_in_names)
    if param:
        value = param.AsString() or param.AsValueString()
        if value:
            return value

    for lookup_name in lookup_names:
        param = element.LookupParameter(lookup_name)
        if param:
            value = param.AsString() or param.AsValueString()
            if value:
                return value

    return default


def get_revision_number(revision):
    if not revision:
        return ""

    number = getattr(revision, 'RevisionNumber', None)
    if number:
        return str(number)

    return get_parameter_value(
        revision,
        built_in_names=['PROJECT_REVISION_REVISION_NUM'],
        lookup_names=['Revision Number']
    )


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
        revision_items.append({
            'name': build_revision_label(revision),
            'revision_id': revision.Id,
            'sort_key': natural_key(get_revision_number(revision)),
        })

    return sorted(revision_items, key=lambda item: (item['sort_key'], item['name'].lower()))


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
        lookup[sheet.SheetNumber] = sheet
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
    comment_param = revision_cloud.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    return safe_as_string(comment_param)


def collect_clouds_by_sheet(selected_revision_id, sheet_by_number):
    revision_clouds = DB.FilteredElementCollector(doc) \
                        .OfCategory(DB.BuiltInCategory.OST_RevisionClouds) \
                        .WhereElementIsNotElementType() \
                        .ToElements()

    data_by_sheet = {}

    for revision_cloud in revision_clouds:
        if revision_cloud.RevisionId != selected_revision_id:
            continue

        sheet_name, sheet_number, owner_view = get_sheet_info(revision_cloud, sheet_by_number)
        if sheet_number not in data_by_sheet:
            data_by_sheet[sheet_number] = {'name': sheet_name, 'views': {}}

        view_name = owner_view.Name if owner_view else 'Unknown View'
        data_by_sheet[sheet_number]['views'].setdefault(view_name, []).append(revision_cloud)

    return data_by_sheet


def print_report(data_by_sheet):
    total_clouds = 0
    sorted_sheets = sorted(data_by_sheet.items(), key=lambda item: natural_key(item[0]))

    for sheet_number, sheet_data in sorted_sheets:
        output.print_md('### **Sheet:** {} - {}\n'.format(sheet_number, sheet_data['name']))

        view_items = sorted(sheet_data['views'].items(), key=lambda item: natural_key(item[0]))
        for view_name, clouds in view_items:
            output.print_md('#### **View:** {}\n'.format(view_name))

            try:
                clouds.sort(key=lambda cloud: cloud.Id.IntegerValue)
            except AttributeError as exc:
                logger.warning('Could not sort clouds in view "%s": %s', view_name, exc)

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

    output.print_md('\n**SEARCH COMPLETED - {} revision clouds found.**'.format(total_clouds))


def main():
    revision_items = load_revisions()
    selected_item = select_revision(revision_items)
    if not selected_item:
        script.exit()

    selected_revision_id = selected_item['revision_id']
    sheet_by_number = build_sheet_lookup()
    data_by_sheet = collect_clouds_by_sheet(selected_revision_id, sheet_by_number)
    print_report(data_by_sheet)


if __name__ == '__main__':
    main()

# -*- coding: utf-8 -*-

from pyrevit import forms, script
from pyrevit.revit import tabs
from pyrevit.userconfig import user_config


def _safe_bool(value):
    return bool(value)


def _get_theme():
    return tabs.get_tabcoloring_theme(user_config)


def _theme_sort_state(theme):
    if theme and hasattr(theme, 'SortDocTabs'):
        return _safe_bool(theme.SortDocTabs)
    return False


def _print_report(output, active_state, sort_state, styled_docs_count, lines):
    output.print_md('## MENVIC | TAB MANAGER')
    output.print_md('- Active: {0}'.format(active_state))
    output.print_md('- Sort Document Tabs: {0}'.format(sort_state))
    output.print_md('- Styled documents count: {0}'.format(styled_docs_count))

    if lines:
        output.print_md('')
        for line in lines:
            output.print_md(line)


def _format_hex(color_obj):
    if color_obj is None:
        return 'N/A'

    for attr_name in ['ToHexString', 'to_hex', 'AsHex', 'Hex', 'hex']:
        if hasattr(color_obj, attr_name):
            attr_val = getattr(color_obj, attr_name)
            try:
                if callable(attr_val):
                    result = attr_val()
                else:
                    result = attr_val
                if result:
                    return str(result)
            except Exception:
                pass

    return str(color_obj)


def _list_document_colors():
    detail_lines = []
    slots = tabs.get_styled_slots() or []

    for slot in slots:
        slot_id = getattr(slot, 'slot_id', None)
        if slot_id is None:
            slot_id = getattr(slot, 'SlotId', 'N/A')

        is_family = getattr(slot, 'is_family', None)
        if is_family is None:
            is_family = getattr(slot, 'IsFamily', False)

        color_val = getattr(slot, 'color', None)
        if color_val is None:
            color_val = getattr(slot, 'Color', None)

        color_hex = _format_hex(color_val)

        detail_lines.append('- Slot {0} | Color: {1} | Family: {2}'.format(slot_id, color_hex, bool(is_family)))

    return detail_lines, len(slots)


def main():
    output = script.get_output()

    options = [
        'Enable Color + Sort Tabs',
        'Disable Tab Coloring',
        'Reset Theme',
        'List Document Colors',
        'Cancel'
    ]

    selected = forms.CommandSwitchWindow.show(
        options,
        message='Select Tab Manager action:',
        title='MENVIC | TAB MANAGER'
    )

    if not selected or selected == 'Cancel':
        script.exit()

    details = []

    if selected == 'Enable Color + Sort Tabs':
        theme = _get_theme()
        if theme and hasattr(theme, 'SortDocTabs'):
            theme.SortDocTabs = True
            tabs.set_tabcoloring_theme(user_config, theme)
        user_config.colorize_docs = True
        user_config.save_changes()
        tabs.init_doc_colorizer(user_config)
        details.append('- Action: Enabled document tab coloring and sorting.')

    elif selected == 'Disable Tab Coloring':
        user_config.colorize_docs = False
        user_config.save_changes()
        tabs.init_doc_colorizer(user_config)
        details.append('- Action: Disabled document tab coloring.')

    elif selected == 'Reset Theme':
        tabs.reset_doc_colorizer()
        details.append('- Action: Reset tab coloring theme.')

    elif selected == 'List Document Colors':
        doc_state = tabs.get_doc_colorizer_state()
        details.append('- Colorizer state: {0}'.format(doc_state))

    active_state = _safe_bool(tabs.get_doc_colorizer_state())
    current_theme = _get_theme()
    sort_state = _theme_sort_state(current_theme)

    listed_count = 0
    if selected == 'List Document Colors':
        list_lines, listed_count = _list_document_colors()
        details.extend(list_lines)

    styled_slots = tabs.get_styled_slots() or []
    styled_docs_count = len(styled_slots)

    if selected == 'List Document Colors' and listed_count != styled_docs_count:
        details.append('- Note: listed slots {0}, current styled slots {1}.'.format(listed_count, styled_docs_count))

    _print_report(output, active_state, sort_state, styled_docs_count, details)


if __name__ == '__main__':
    main()

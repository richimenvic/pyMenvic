# -*- coding: utf-8 -*-

__title__ = "Tab Manager"
__author__ = "Ricardo J. Mendieta"

from pyrevit import script
from pyrevit.revit import tabs
from pyrevit.userconfig import user_config


def _safe_bool(value):
    try:
        return bool(value)
    except:
        return False


def _get_theme():
    try:
        return tabs.get_tabcoloring_theme(user_config)
    except:
        return None


def _get_sort_state(theme):
    if theme and hasattr(theme, "SortDocTabs"):
        return _safe_bool(theme.SortDocTabs)
    return False


def _save_config():
    try:
        user_config.save_changes()
    except:
        pass


def _set_icon(state):
    try:
        script.toggle_icon(state)
    except:
        pass


def _slots_count():
    try:
        slots = tabs.get_styled_slots() or []
        return len(slots)
    except:
        return 0


def _enable_tabs():
    theme = _get_theme()
    if theme and hasattr(theme, "SortDocTabs"):
        theme.SortDocTabs = True
        tabs.set_tabcoloring_theme(user_config, theme)

    user_config.colorize_docs = True
    _save_config()
    tabs.init_doc_colorizer(user_config)
    _set_icon(True)
    return "Enabled color and sort."


def _stop_tabs():
    user_config.colorize_docs = False
    _save_config()
    tabs.init_doc_colorizer(user_config)
    _set_icon(False)
    return "Disabled tab coloring."


def _report(message):
    output = script.get_output()
    theme = _get_theme()
    output.print_md("## MENVIC | TAB MANAGER")
    output.print_md("- Action: {0}".format(message))
    output.print_md("- Active: {0}".format(_safe_bool(tabs.get_doc_colorizer_state())))
    output.print_md("- Sort Document Tabs: {0}".format(_get_sort_state(theme)))
    output.print_md("- Styled documents count: {0}".format(_slots_count()))


def main():
    if _safe_bool(tabs.get_doc_colorizer_state()):
        message = _stop_tabs()
    else:
        message = _enable_tabs()
    _report(message)


if __name__ == "__main__":
    main()

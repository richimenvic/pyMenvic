# -*- coding: utf-8 -*-

__title__ = "Tabs by Document"
__author__ = "Ricardo J. Mendieta"

import os
import sys

from pyrevit import script
from pyrevit.revit import tabs
from pyrevit.userconfig import user_config

try:
    from lib.core.tab_sorter import sort_tabs_by_document
except ImportError:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    while True:
        if os.path.basename(current_dir).lower() == "pymenvic.extension":
            lib_dir = os.path.join(current_dir, "lib")
            if lib_dir not in sys.path:
                sys.path.append(lib_dir)
            break
        parent_dir = os.path.dirname(current_dir)
        if parent_dir == current_dir:
            break
        current_dir = parent_dir
    from core.tab_sorter import sort_tabs_by_document


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


def _is_active():
    try:
        return _safe_bool(tabs.get_doc_colorizer_state())
    except:
        return _safe_bool(getattr(user_config, "colorize_docs", False))


def _enable_and_sort_tabs():
    theme = _get_theme()
    if theme and hasattr(theme, "SortDocTabs"):
        theme.SortDocTabs = True
        tabs.set_tabcoloring_theme(user_config, theme)

    user_config.colorize_docs = True
    _save_config()
    tabs.init_doc_colorizer(user_config)
    sort_tabs_by_document()
    _set_icon(True)


def _disable_tab_coloring():
    user_config.colorize_docs = False
    _save_config()
    tabs.init_doc_colorizer(user_config)
    _set_icon(False)


def _print_error(ex):
    output = script.get_output()
    output.print_md("## MENVIC | TABS BY DOCUMENT")
    output.print_md("- Failed: {0}".format(str(ex).split("\n")[0]))


def main():
    try:
        if _is_active():
            _disable_tab_coloring()
        else:
            _enable_and_sort_tabs()
    except Exception as ex:
        _print_error(ex)


if __name__ == "__main__":
    main()

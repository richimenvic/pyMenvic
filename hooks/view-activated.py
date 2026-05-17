# -*- coding: utf-8 -*-

import os
import sys

from pyrevit.revit import events, tabs
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


def _should_sort_tabs():
    if not _safe_bool(getattr(user_config, "colorize_docs", False)):
        return False

    try:
        theme = tabs.get_tabcoloring_theme(user_config)
        if hasattr(theme, "SortDocTabs"):
            return _safe_bool(theme.SortDocTabs)
    except:
        pass

    return False


def _sort_if_enabled():
    try:
        if _should_sort_tabs():
            sort_tabs_by_document()
    except:
        pass


try:
    events.execute_in_revit_context(_sort_if_enabled)
except:
    _sort_if_enabled()

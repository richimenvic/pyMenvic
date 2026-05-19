# -*- coding: utf-8 -*-

import os
import sys

from pyrevit.revit import tabs
from pyrevit.userconfig import user_config

try:
    from System.Threading import Thread
except:
    Thread = None

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


PENDING_ENVVAR = "PYMENVIC_TABS_SORT_PENDING"
PENDING_TICKS = "12"
DELAYED_PASS_MS = 180


def _safe_bool(value):
    try:
        return bool(value)
    except:
        return False


def _sleep(ms):
    if Thread is None:
        return
    try:
        Thread.Sleep(ms)
    except:
        pass


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


def _safe_sort_tabs():
    try:
        return sort_tabs_by_document()
    except:
        return 0


try:
    if _should_sort_tabs():
        os.environ[PENDING_ENVVAR] = PENDING_TICKS

        # First pass keeps the old behavior for already-visible tabs.
        _safe_sort_tabs()

        # Revit can append a newly opened view tab after the first event pass.
        # A very short second pass fixes that without adding a permanent watcher.
        _sleep(DELAYED_PASS_MS)
        _safe_sort_tabs()
except:
    pass

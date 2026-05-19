# -*- coding: utf-8 -*-

import os
import sys

from pyrevit.revit import tabs
from pyrevit.userconfig import user_config
from pyrevit.revit import ui

try:
    from System import Action
    from System.Windows.Threading import DispatcherPriority
except:
    Action = None
    DispatcherPriority = None

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
PENDING_TICKS = "20"


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


def _safe_sort_tabs():
    try:
        return sort_tabs_by_document()
    except:
        return 0


def _queue_dispatcher_sort():
    if Action is None or DispatcherPriority is None:
        return False
    try:
        main_window = ui.get_mainwindow()
        dispatcher = main_window.Dispatcher if main_window is not None else None
        if dispatcher is None:
            return False
        dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, Action(_safe_sort_tabs))
        dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(_safe_sort_tabs))
        dispatcher.BeginInvoke(DispatcherPriority.Background, Action(_safe_sort_tabs))
        return True
    except:
        return False


try:
    if _should_sort_tabs():
        # Keep several idling passes as a lightweight fallback.
        os.environ[PENDING_ENVVAR] = PENDING_TICKS

        # Immediate pass handles tabs already visible at event time.
        _safe_sort_tabs()

        # Dispatcher passes run after Revit/WPF finishes visual tab creation.
        _queue_dispatcher_sort()
except:
    pass

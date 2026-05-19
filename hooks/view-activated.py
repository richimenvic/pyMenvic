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
PYMENVIC_SORT_ENVVAR = "PYMENVIC_TABS_BY_DOCUMENT_ENABLED"
PYMENVIC_SORT_CONFIG = "pymenvic_sort_doc_tabs"
HOOK_HIT_ENVVAR = "PYMENVIC_TABS_HOOK_HIT"
HOOK_SHOULD_ENVVAR = "PYMENVIC_TABS_HOOK_SHOULD_SORT"
HOOK_IMMEDIATE_MOVES_ENVVAR = "PYMENVIC_TABS_HOOK_IMMEDIATE_MOVES"
HOOK_DISPATCHER_ENVVAR = "PYMENVIC_TABS_HOOK_DISPATCHER"
HOOK_ERROR_ENVVAR = "PYMENVIC_TABS_HOOK_ERROR"


def _safe_bool(value):
    try:
        return bool(value)
    except:
        return False


def _safe_int(value, default_value):
    try:
        return int(value)
    except:
        return default_value


def _set_env(name, value):
    try:
        os.environ[name] = str(value)
    except:
        pass


def _should_sort_tabs():
    try:
        if os.environ.get(PYMENVIC_SORT_ENVVAR, "") == "1":
            return True
    except:
        pass

    if _safe_bool(getattr(user_config, PYMENVIC_SORT_CONFIG, False)):
        return True

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
    except Exception as ex:
        _set_env(HOOK_ERROR_ENVVAR, str(ex).split("\n")[0])
        return 0


def _dispatcher_sort():
    if Action is None or DispatcherPriority is None:
        return False
    try:
        main_window = ui.get_mainwindow()
        dispatcher = main_window.Dispatcher if main_window is not None else None
        if dispatcher is None:
            return False

        # Use Invoke instead of BeginInvoke. In pyRevit hooks, async delegates can be
        # lost when the hook engine exits. Probe confirmed synchronous Invoke works.
        dispatcher.Invoke(DispatcherPriority.ApplicationIdle, Action(_safe_sort_tabs))
        dispatcher.Invoke(DispatcherPriority.ContextIdle, Action(_safe_sort_tabs))
        dispatcher.Invoke(DispatcherPriority.Background, Action(_safe_sort_tabs))
        return True
    except Exception as ex:
        _set_env(HOOK_ERROR_ENVVAR, str(ex).split("\n")[0])
        return False


try:
    hit_count = _safe_int(os.environ.get(HOOK_HIT_ENVVAR, "0"), 0) + 1
    _set_env(HOOK_HIT_ENVVAR, hit_count)
    _set_env(HOOK_ERROR_ENVVAR, "")

    should_sort = _should_sort_tabs()
    _set_env(HOOK_SHOULD_ENVVAR, "1" if should_sort else "0")

    if should_sort:
        # Keep several idling passes as a lightweight fallback.
        os.environ[PENDING_ENVVAR] = PENDING_TICKS

        # Immediate pass handles tabs already visible at event time.
        immediate_moves = _safe_sort_tabs()
        _set_env(HOOK_IMMEDIATE_MOVES_ENVVAR, immediate_moves)

        # Dispatcher passes run after Revit/WPF finishes visual tab creation.
        dispatcher_ok = _dispatcher_sort()
        _set_env(HOOK_DISPATCHER_ENVVAR, "1" if dispatcher_ok else "0")
    else:
        _set_env(HOOK_IMMEDIATE_MOVES_ENVVAR, "skipped")
        _set_env(HOOK_DISPATCHER_ENVVAR, "skipped")
except Exception as ex:
    _set_env(HOOK_ERROR_ENVVAR, str(ex).split("\n")[0])

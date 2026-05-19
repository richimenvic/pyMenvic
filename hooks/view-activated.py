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
    from lib.core import tab_state
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
    from core import tab_state


PENDING_ENVVAR = "PYMENVIC_TABS_SORT_PENDING"
PENDING_TICKS = "20"


def _safe_int(value, default_value):
    try:
        return int(value)
    except:
        return default_value


def _safe_sort_tabs():
    try:
        return sort_tabs_by_document()
    except Exception as ex:
        tab_state.update_state(HOOK_ERROR=str(ex).split("\n")[0])
        return 0


def _dispatcher_sort():
    if Action is None or DispatcherPriority is None:
        return False
    try:
        main_window = ui.get_mainwindow()
        dispatcher = main_window.Dispatcher if main_window is not None else None
        if dispatcher is None:
            return False
        dispatcher.Invoke(DispatcherPriority.ApplicationIdle, Action(_safe_sort_tabs))
        dispatcher.Invoke(DispatcherPriority.ContextIdle, Action(_safe_sort_tabs))
        dispatcher.Invoke(DispatcherPriority.Background, Action(_safe_sort_tabs))
        return True
    except Exception as ex:
        tab_state.update_state(HOOK_ERROR=str(ex).split("\n")[0])
        return False


try:
    state = tab_state.read_state()
    hit_count = _safe_int(state.get("HOOK_HIT", "0"), 0) + 1
    tab_state.update_state(HOOK_HIT=hit_count, HOOK_ERROR="")

    should_sort = tab_state.is_enabled(user_config, tabs)
    tab_state.update_state(HOOK_SHOULD_SORT="1" if should_sort else "0")

    if should_sort:
        os.environ[PENDING_ENVVAR] = PENDING_TICKS
        immediate_moves = _safe_sort_tabs()
        dispatcher_ok = _dispatcher_sort()
        tab_state.update_state(
            HOOK_IMMEDIATE_MOVES=immediate_moves,
            HOOK_DISPATCHER="1" if dispatcher_ok else "0",
        )
    else:
        tab_state.update_state(HOOK_IMMEDIATE_MOVES="skipped", HOOK_DISPATCHER="skipped")
except Exception as ex:
    tab_state.update_state(HOOK_ERROR=str(ex).split("\n")[0])

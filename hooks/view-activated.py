# -*- coding: utf-8 -*-

import os
import sys

from pyrevit.revit import tabs, ui
from pyrevit.userconfig import user_config

try:
    from System import TimeSpan
    from System.Windows.Threading import DispatcherTimer
except:
    TimeSpan = None
    DispatcherTimer = None

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


_DELAY_MS = 450
_MAX_TICKS = 4
_TIMER_REF = None


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


def _start_delayed_sort():
    global _TIMER_REF

    if TimeSpan is None or DispatcherTimer is None:
        _sort_if_enabled()
        return

    try:
        dispatcher = ui.get_mainwindow().Dispatcher
        timer = DispatcherTimer(dispatcher)
        timer.Interval = TimeSpan.FromMilliseconds(_DELAY_MS)
        state = {"ticks": 0}

        def _on_tick(sender, args):
            try:
                state["ticks"] += 1
                _sort_if_enabled()
                if state["ticks"] >= _MAX_TICKS:
                    sender.Stop()
            except:
                try:
                    sender.Stop()
                except:
                    pass

        timer.Tick += _on_tick
        timer.Start()
        _TIMER_REF = timer
    except:
        _sort_if_enabled()


_start_delayed_sort()

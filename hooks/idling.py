# -*- coding: utf-8 -*-

import os
import sys
import time

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


PENDING_ENVVAR = "PYMENVIC_TABS_SORT_PENDING"
LAST_RUN_ENVVAR = "PYMENVIC_TABS_SORT_LAST_RUN"
MIN_INTERVAL_SECONDS = 0.35


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


def _safe_float(value, default_value):
    try:
        return float(value)
    except:
        return default_value


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


def _consume_pending_sort():
    pending = _safe_int(os.environ.get(PENDING_ENVVAR, "0"), 0)
    if pending <= 0:
        return

    if not _should_sort_tabs():
        os.environ[PENDING_ENVVAR] = "0"
        return

    now = time.time()
    last_run = _safe_float(os.environ.get(LAST_RUN_ENVVAR, "0"), 0.0)
    if now - last_run < MIN_INTERVAL_SECONDS:
        return

    try:
        sort_tabs_by_document()
    except:
        pass

    os.environ[LAST_RUN_ENVVAR] = str(now)
    pending -= 1
    if pending < 0:
        pending = 0
    os.environ[PENDING_ENVVAR] = str(pending)


try:
    _consume_pending_sort()
except:
    pass

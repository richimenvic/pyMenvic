# -*- coding: utf-8 -*-

import os
import sys
import time

from pyrevit import HOST_APP
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


PYMENVIC_SORT_ENVVAR = "PYMENVIC_TABS_BY_DOCUMENT_ENABLED"
PYMENVIC_SORT_CONFIG = "pymenvic_sort_doc_tabs"

PENDING_ENVVAR = "PYMENVIC_TABS_SORT_PENDING"
LAST_RUN_ENVVAR = "PYMENVIC_TABS_SORT_LAST_RUN"

STARTUP_LOADED_ENVVAR = "PYMENVIC_TABS_STARTUP_LOADED"
SUBSCRIBED_ENVVAR = "PYMENVIC_TABS_EVENT_SUBSCRIBED"
STARTUP_ERROR_ENVVAR = "PYMENVIC_TABS_STARTUP_ERROR"

HOOK_HIT_ENVVAR = "PYMENVIC_TABS_HOOK_HIT"
HOOK_SHOULD_ENVVAR = "PYMENVIC_TABS_HOOK_SHOULD_SORT"

IDLING_HIT_ENVVAR = "PYMENVIC_TABS_IDLING_HIT"
IDLING_SHOULD_ENVVAR = "PYMENVIC_TABS_IDLING_SHOULD_SORT"
IDLING_MOVES_ENVVAR = "PYMENVIC_TABS_IDLING_MOVES"
IDLING_ERROR_ENVVAR = "PYMENVIC_TABS_IDLING_ERROR"

PENDING_TICKS = 20
MIN_INTERVAL_SECONDS = 0.30
DELAY_BEFORE_SORT_TICKS = 3


def _set_env(name, value):
    try:
        os.environ[name] = str(value)
    except:
        pass


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


def _on_view_activated(sender, args):
    try:
        hit_count = _safe_int(os.environ.get(HOOK_HIT_ENVVAR, "0"), 0) + 1
        _set_env(HOOK_HIT_ENVVAR, hit_count)

        should_sort = _should_sort_tabs()
        _set_env(HOOK_SHOULD_ENVVAR, "1" if should_sort else "0")
        if should_sort:
            _set_env(PENDING_ENVVAR, PENDING_TICKS)
    except Exception as ex:
        _set_env(STARTUP_ERROR_ENVVAR, str(ex).split("\n")[0])


def _on_idling(sender, args):
    hit_count = _safe_int(os.environ.get(IDLING_HIT_ENVVAR, "0"), 0) + 1
    _set_env(IDLING_HIT_ENVVAR, hit_count)
    _set_env(IDLING_ERROR_ENVVAR, "")

    try:
        should_sort = _should_sort_tabs()
        _set_env(IDLING_SHOULD_ENVVAR, "1" if should_sort else "0")
        if not should_sort:
            _set_env(PENDING_ENVVAR, "0")
            _set_env(IDLING_MOVES_ENVVAR, "skipped")
            return

        pending = _safe_int(os.environ.get(PENDING_ENVVAR, "0"), 0)
        if pending <= 0:
            return

        pending -= 1
        _set_env(PENDING_ENVVAR, pending)
        if pending > (PENDING_TICKS - DELAY_BEFORE_SORT_TICKS):
            _set_env(IDLING_MOVES_ENVVAR, "waiting")
            return

        now = time.time()
        last_run = _safe_float(os.environ.get(LAST_RUN_ENVVAR, "0"), 0.0)
        if now - last_run < MIN_INTERVAL_SECONDS:
            return

        moves = sort_tabs_by_document()
        _set_env(IDLING_MOVES_ENVVAR, moves)
        _set_env(LAST_RUN_ENVVAR, now)
    except Exception as ex:
        _set_env(IDLING_ERROR_ENVVAR, str(ex).split("\n")[0])


def _subscribe_events():
    try:
        uiapp = HOST_APP.uiapp
        if uiapp is None:
            _set_env(SUBSCRIBED_ENVVAR, "uiapp-none")
            return

        if os.environ.get(SUBSCRIBED_ENVVAR, "") == "1":
            return

        uiapp.ViewActivated += _on_view_activated
        uiapp.Idling += _on_idling
        _set_env(SUBSCRIBED_ENVVAR, "1")
    except Exception as ex:
        _set_env(STARTUP_ERROR_ENVVAR, str(ex).split("\n")[0])
        _set_env(SUBSCRIBED_ENVVAR, "0")


_set_env(STARTUP_LOADED_ENVVAR, "1")
_set_env(STARTUP_ERROR_ENVVAR, "")
_subscribe_events()

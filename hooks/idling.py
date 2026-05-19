# -*- coding: utf-8 -*-

import os
import sys
import time

from pyrevit.revit import tabs
from pyrevit.userconfig import user_config

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
LAST_RUN_ENVVAR = "PYMENVIC_TABS_SORT_LAST_RUN"
MIN_INTERVAL_SECONDS = 0.50


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


def _run_sort_if_due():
    state = tab_state.read_state()
    hit_count = _safe_int(state.get("IDLING_HIT", "0"), 0) + 1
    should_sort = tab_state.is_enabled(user_config, tabs)
    tab_state.update_state(
        IDLING_HIT=hit_count,
        IDLING_SHOULD_SORT="1" if should_sort else "0",
        IDLING_ERROR="",
    )

    if not should_sort:
        os.environ[PENDING_ENVVAR] = "0"
        tab_state.update_state(IDLING_MOVES="skipped")
        return

    now = time.time()
    last_run = _safe_float(os.environ.get(LAST_RUN_ENVVAR, "0"), 0.0)
    if now - last_run < MIN_INTERVAL_SECONDS:
        return

    try:
        moves = sort_tabs_by_document()
        tab_state.update_state(IDLING_MOVES=moves, IDLING_LAST_RUN=now)
    except Exception as ex:
        tab_state.update_state(IDLING_ERROR=str(ex).split("\n")[0])

    os.environ[LAST_RUN_ENVVAR] = str(now)

    pending = _safe_int(os.environ.get(PENDING_ENVVAR, "0"), 0)
    if pending > 0:
        pending -= 1
        if pending < 0:
            pending = 0
        os.environ[PENDING_ENVVAR] = str(pending)


try:
    _run_sort_if_due()
except Exception as ex:
    tab_state.update_state(IDLING_ERROR=str(ex).split("\n")[0])

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


PYMENVIC_SORT_ENVVAR = "PYMENVIC_TABS_BY_DOCUMENT_ENABLED"
PYMENVIC_SORT_CONFIG = "pymenvic_sort_doc_tabs"
LAST_RUN_ENVVAR = "PYMENVIC_TABS_COMMAND_LAST_RUN"
MIN_INTERVAL_SECONDS = 0.25
STATE_FILE = os.path.join(os.environ.get("TEMP", os.getcwd()), "pyMenvic_tab_sort_state.txt")


def _safe_bool(value):
    try:
        return bool(value)
    except:
        return False


def _safe_float(value, default_value):
    try:
        return float(value)
    except:
        return default_value


def _read_state():
    data = {}
    try:
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, "r") as state_file:
                for line in state_file:
                    if "=" in line:
                        key, value = line.rstrip("\n").split("=", 1)
                        data[key] = value
    except:
        pass
    return data


def _write_state(data):
    try:
        folder = os.path.dirname(STATE_FILE)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(STATE_FILE, "w") as state_file:
            for key in sorted(data.keys()):
                state_file.write("{0}={1}\n".format(key, data[key]))
    except:
        pass


def _update_state(**kwargs):
    data = _read_state()
    for key, value in kwargs.items():
        data[key] = str(value)
    _write_state(data)


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


def _run_sort_if_due():
    should_sort = _should_sort_tabs()
    _update_state(
        STATE_FILE=STATE_FILE,
        COMMAND_HIT=time.time(),
        COMMAND_SHOULD_SORT="1" if should_sort else "0",
        COMMAND_ERROR="",
    )

    if not should_sort:
        _update_state(COMMAND_MOVES="skipped")
        return

    now = time.time()
    last_run = _safe_float(os.environ.get(LAST_RUN_ENVVAR, "0"), 0.0)
    if now - last_run < MIN_INTERVAL_SECONDS:
        return

    try:
        moves = sort_tabs_by_document()
        _update_state(COMMAND_MOVES=moves, COMMAND_LAST_RUN=now)
    except Exception as ex:
        _update_state(COMMAND_ERROR=str(ex).split("\n")[0])

    os.environ[LAST_RUN_ENVVAR] = str(now)


try:
    _run_sort_if_due()
except Exception as ex:
    _update_state(COMMAND_ERROR=str(ex).split("\n")[0])

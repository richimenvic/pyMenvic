# -*- coding: utf-8 -*-

import os
import sys
import time

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
MIN_INTERVAL_SECONDS = 0.50
STATE_FILE = os.path.join(os.environ.get("LOCALAPPDATA", os.environ.get("TEMP", os.getcwd())), "Temp", "pyMenvic", "tab_sort_state.txt")
if not os.environ.get("LOCALAPPDATA", ""):
    STATE_FILE = os.path.join(os.environ.get("TEMP", os.getcwd()), "pyMenvic", "tab_sort_state.txt")


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


def _read_state():
    data = {}
    try:
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, "r") as state_file:
                for line in state_file:
                    if "=" in line:
                        key, value = line.split("=", 1)
                        data[str(key).strip()] = str(value).strip()
    except:
        pass
    return data


def _write_state(data):
    try:
        folder = os.path.dirname(STATE_FILE)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        data["STATE_FILE"] = STATE_FILE
        with open(STATE_FILE, "w") as state_file:
            for key in sorted(data.keys()):
                state_file.write("{0}={1}\n".format(str(key).strip(), str(data[key]).strip()))
    except:
        pass


def _update_state(**kwargs):
    data = _read_state()
    for key, value in kwargs.items():
        data[str(key).strip()] = str(value).strip()
    _write_state(data)


def _is_enabled():
    state = _read_state()
    return state.get("ENABLED", "").strip() == "1"


def _run_sort_if_due():
    state = _read_state()
    hit_count = _safe_int(state.get("IDLING_HIT", "0"), 0) + 1
    should_sort = _is_enabled()
    _update_state(
        IDLING_HIT=hit_count,
        IDLING_SHOULD_SORT="1" if should_sort else "0",
        IDLING_ERROR="",
    )

    if not should_sort:
        os.environ[PENDING_ENVVAR] = "0"
        _update_state(IDLING_MOVES="skipped")
        return

    now = time.time()
    last_run = _safe_float(os.environ.get(LAST_RUN_ENVVAR, "0"), 0.0)
    if now - last_run < MIN_INTERVAL_SECONDS:
        return

    try:
        moves = sort_tabs_by_document()
        _update_state(IDLING_MOVES=moves, IDLING_LAST_RUN=now)
    except Exception as ex:
        _update_state(IDLING_ERROR=str(ex).split("\n")[0])

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
    _update_state(IDLING_ERROR=str(ex).split("\n")[0])

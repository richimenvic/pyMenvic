# -*- coding: utf-8 -*-

import os
import sys

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
STATE_FILE = os.path.join(os.environ.get("LOCALAPPDATA", os.environ.get("TEMP", os.getcwd())), "Temp", "pyMenvic", "tab_sort_state.txt")
if not os.environ.get("LOCALAPPDATA", ""):
    STATE_FILE = os.path.join(os.environ.get("TEMP", os.getcwd()), "pyMenvic", "tab_sort_state.txt")


def _safe_int(value, default_value):
    try:
        return int(value)
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


def _safe_sort_tabs():
    try:
        return sort_tabs_by_document()
    except Exception as ex:
        _update_state(HOOK_ERROR=str(ex).split("\n")[0])
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
        _update_state(HOOK_ERROR=str(ex).split("\n")[0])
        return False


try:
    state = _read_state()
    hit_count = _safe_int(state.get("HOOK_HIT", "0"), 0) + 1
    _update_state(HOOK_HIT=hit_count, HOOK_ERROR="")

    should_sort = _is_enabled()
    _update_state(HOOK_SHOULD_SORT="1" if should_sort else "0")

    if should_sort:
        os.environ[PENDING_ENVVAR] = PENDING_TICKS
        immediate_moves = _safe_sort_tabs()
        dispatcher_ok = _dispatcher_sort()
        _update_state(
            HOOK_IMMEDIATE_MOVES=immediate_moves,
            HOOK_DISPATCHER="1" if dispatcher_ok else "0",
        )
    else:
        _update_state(HOOK_IMMEDIATE_MOVES="skipped", HOOK_DISPATCHER="skipped")
except Exception as ex:
    _update_state(HOOK_ERROR=str(ex).split("\n")[0])

# -*- coding: utf-8 -*-

import os
import sys
import time

from pyrevit import HOST_APP

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


STATE_FILE = os.path.join(os.environ.get("LOCALAPPDATA", os.environ.get("TEMP", os.getcwd())), "Temp", "pyMenvic", "tab_sort_request.txt")
if not os.environ.get("LOCALAPPDATA", ""):
    STATE_FILE = os.path.join(os.environ.get("TEMP", os.getcwd()), "pyMenvic", "tab_sort_request.txt")

REQUEST_VALUE = "1"
CONSUMED_VALUE = "0"
SUBSCRIBED_ENVVAR = "PYMENVIC_TABS_STARTUP_SUBSCRIBED"
LAST_RUN_ENVVAR = "PYMENVIC_TABS_STARTUP_LAST_RUN"
MIN_INTERVAL_SECONDS = 0.30


def _read_request():
    try:
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, "r") as state_file:
                return state_file.read().strip()
    except:
        pass
    return ""


def _write_request(value):
    try:
        folder = os.path.dirname(STATE_FILE)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(STATE_FILE, "w") as state_file:
            state_file.write(str(value))
    except:
        pass


def _safe_float(value, default_value):
    try:
        return float(value)
    except:
        return default_value


def _on_idling(sender, args):
    try:
        if _read_request() != REQUEST_VALUE:
            return

        now = time.time()
        last_run = _safe_float(os.environ.get(LAST_RUN_ENVVAR, "0"), 0.0)
        if now - last_run < MIN_INTERVAL_SECONDS:
            return

        # Consume first so this never becomes a persistent auto-sorter.
        _write_request(CONSUMED_VALUE)
        os.environ[LAST_RUN_ENVVAR] = str(now)
        sort_tabs_by_document()
    except:
        try:
            _write_request(CONSUMED_VALUE)
        except:
            pass


def _subscribe_events():
    try:
        if os.environ.get(SUBSCRIBED_ENVVAR, "") == "1":
            return
        uiapp = HOST_APP.uiapp
        if uiapp is None:
            return
        uiapp.Idling += _on_idling
        os.environ[SUBSCRIBED_ENVVAR] = "1"
    except:
        pass


_subscribe_events()

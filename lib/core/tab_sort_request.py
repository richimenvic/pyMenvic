# -*- coding: utf-8 -*-

import os


REQUEST_VALUE = "1"
CONSUMED_VALUE = "0"


def _get_base_temp_dir():
    local_app_data = os.environ.get("LOCALAPPDATA", "")
    if local_app_data:
        return os.path.join(local_app_data, "Temp")
    temp_dir = os.environ.get("TEMP", "")
    if temp_dir:
        return temp_dir
    return os.getcwd()


STATE_DIR = os.path.join(_get_base_temp_dir(), "pyMenvic")
STATE_FILE = os.path.join(STATE_DIR, "tab_sort_request.txt")


def _ensure_state_dir():
    try:
        if not os.path.exists(STATE_DIR):
            os.makedirs(STATE_DIR)
    except:
        pass


def _read_value():
    try:
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, "r") as state_file:
                return state_file.read().strip()
    except:
        pass
    return ""


def _write_value(value):
    try:
        _ensure_state_dir()
        with open(STATE_FILE, "w") as state_file:
            state_file.write(str(value))
    except:
        pass


def request_sort():
    _write_value(REQUEST_VALUE)


def consume_sort_request():
    if _read_value() != REQUEST_VALUE:
        return False
    _write_value(CONSUMED_VALUE)
    return True

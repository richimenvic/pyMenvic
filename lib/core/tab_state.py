# -*- coding: utf-8 -*-

import os


PYMENVIC_SORT_ENVVAR = "PYMENVIC_TABS_BY_DOCUMENT_ENABLED"
PYMENVIC_SORT_CONFIG = "pymenvic_sort_doc_tabs"


def _get_base_temp_dir():
    local_app_data = os.environ.get("LOCALAPPDATA", "")
    if local_app_data:
        return os.path.join(local_app_data, "Temp")
    temp_dir = os.environ.get("TEMP", "")
    if temp_dir:
        return temp_dir
    return os.getcwd()


STATE_DIR = os.path.join(_get_base_temp_dir(), "pyMenvic")
STATE_FILE = os.path.join(STATE_DIR, "tab_sort_state.txt")


def safe_bool(value):
    try:
        if isinstance(value, basestring):
            return value.strip().lower() in ["1", "true", "yes", "on"]
    except:
        pass
    try:
        return bool(value)
    except:
        return False


def read_state():
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


def write_state(data):
    try:
        if not os.path.exists(STATE_DIR):
            os.makedirs(STATE_DIR)
        data["STATE_FILE"] = STATE_FILE
        with open(STATE_FILE, "w") as state_file:
            for key in sorted(data.keys()):
                state_file.write("{0}={1}\n".format(key, data[key]))
    except:
        pass


def update_state(**kwargs):
    data = read_state()
    for key, value in kwargs.items():
        data[key] = str(value)
    write_state(data)


def set_enabled(user_config, enabled):
    try:
        setattr(user_config, PYMENVIC_SORT_CONFIG, bool(enabled))
    except:
        pass
    try:
        os.environ[PYMENVIC_SORT_ENVVAR] = "1" if enabled else "0"
    except:
        pass
    update_state(ENABLED="1" if enabled else "0")


def is_enabled(user_config, tabs_module):
    state = read_state()
    if state.get("ENABLED", "") == "1":
        return True
    if state.get("ENABLED", "") == "0":
        return False

    if safe_bool(getattr(user_config, PYMENVIC_SORT_CONFIG, False)):
        return True

    try:
        if os.environ.get(PYMENVIC_SORT_ENVVAR, "") == "1":
            return True
    except:
        pass

    if not safe_bool(getattr(user_config, "colorize_docs", False)):
        return False

    try:
        theme = tabs_module.get_tabcoloring_theme(user_config)
        if hasattr(theme, "SortDocTabs"):
            return safe_bool(theme.SortDocTabs)
    except:
        pass

    return False

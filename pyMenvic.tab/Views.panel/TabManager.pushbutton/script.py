# -*- coding: utf-8 -*-

__title__ = "Tabs by Document"
__author__ = "Ricardo J. Mendieta"

import os
import sys

from pyrevit import EXEC_PARAMS, script
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
STATE_FILE = os.path.join(os.environ.get("TEMP", os.getcwd()), "pyMenvic_tab_sort_state.txt")


def _safe_bool(value):
    try:
        if isinstance(value, basestring):
            return value.strip().lower() in ["1", "true", "yes", "on"]
    except:
        pass
    try:
        return bool(value)
    except:
        return False


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
    data["STATE_FILE"] = STATE_FILE
    _write_state(data)


def _get_theme():
    try:
        return tabs.get_tabcoloring_theme(user_config)
    except:
        return None


def _save_config():
    try:
        user_config.save_changes()
    except:
        pass


def _set_icon(state):
    try:
        script.toggle_icon(state)
    except:
        pass


def _clear_runtime_flags():
    for key in [
        "PYMENVIC_TABS_SORT_PENDING",
        "PYMENVIC_TABS_HOOK_HIT",
        "PYMENVIC_TABS_HOOK_SHOULD_SORT",
        "PYMENVIC_TABS_HOOK_IMMEDIATE_MOVES",
        "PYMENVIC_TABS_HOOK_DISPATCHER",
        "PYMENVIC_TABS_HOOK_ERROR",
        "PYMENVIC_TABS_IDLING_HIT",
        "PYMENVIC_TABS_IDLING_SHOULD_SORT",
        "PYMENVIC_TABS_IDLING_MOVES",
        "PYMENVIC_TABS_IDLING_ERROR",
        "PYMENVIC_TABS_COMMAND_LAST_RUN",
    ]:
        try:
            if key in os.environ:
                del os.environ[key]
        except:
            pass


def _set_pymenvic_sort_flag(state):
    try:
        setattr(user_config, PYMENVIC_SORT_CONFIG, bool(state))
    except:
        pass
    try:
        os.environ[PYMENVIC_SORT_ENVVAR] = "1" if state else "0"
    except:
        pass
    _update_state(ENABLED="1" if state else "0")


def _is_sort_enabled():
    state = _read_state()
    if state.get("ENABLED", "") == "1":
        return True
    if state.get("ENABLED", "") == "0":
        return False

    if _safe_bool(getattr(user_config, PYMENVIC_SORT_CONFIG, False)):
        return True
    try:
        if os.environ.get(PYMENVIC_SORT_ENVVAR, "") == "1":
            return True
    except:
        pass
    return False


def _enable_and_sort_tabs():
    theme = _get_theme()
    if theme and hasattr(theme, "SortDocTabs"):
        theme.SortDocTabs = True
        tabs.set_tabcoloring_theme(user_config, theme)

    user_config.colorize_docs = True
    _set_pymenvic_sort_flag(True)
    _save_config()
    tabs.init_doc_colorizer(user_config)
    sort_tabs_by_document()
    _set_icon(True)


def _disable_tab_sorting():
    theme = _get_theme()
    if theme and hasattr(theme, "SortDocTabs"):
        theme.SortDocTabs = False
        tabs.set_tabcoloring_theme(user_config, theme)

    _set_pymenvic_sort_flag(False)
    _save_config()
    _clear_runtime_flags()

    try:
        tabs.init_doc_colorizer(user_config)
    except:
        pass
    _set_icon(False)


def _is_config_mode():
    try:
        return bool(EXEC_PARAMS.config_mode)
    except:
        return False


def _print_status(enabled):
    output = script.get_output()
    output.print_md("## MENVIC | TABS BY DOCUMENT")
    if enabled:
        output.print_md("- Auto tab sorting: `ON`")
    else:
        output.print_md("- Auto tab sorting: `OFF`")
    output.print_md("- State file: `{0}`".format(STATE_FILE))


def _print_error(ex):
    output = script.get_output()
    output.print_md("## MENVIC | TABS BY DOCUMENT")
    output.print_md("- Failed: {0}".format(str(ex).split("\n")[0]))


def main():
    try:
        if _is_config_mode():
            _disable_tab_sorting()
            _print_status(False)
            return

        if _is_sort_enabled():
            _disable_tab_sorting()
            _print_status(False)
        else:
            _enable_and_sort_tabs()
            _print_status(True)
    except Exception as ex:
        _print_error(ex)


if __name__ == "__main__":
    main()

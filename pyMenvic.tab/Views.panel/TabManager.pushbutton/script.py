# -*- coding: utf-8 -*-

__title__ = "Tab Manager"
__author__ = "Ricardo J. Mendieta"

from pyrevit import script
from pyrevit.revit import tabs
from pyrevit.userconfig import user_config


def _safe_bool(value):
    try:
        return bool(value)
    except:
        return False


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


def _enable_tabs():
    theme = _get_theme()
    if theme and hasattr(theme, "SortDocTabs"):
        theme.SortDocTabs = True
        tabs.set_tabcoloring_theme(user_config, theme)

    user_config.colorize_docs = True
    _save_config()
    tabs.init_doc_colorizer(user_config)
    _set_icon(True)


def _stop_tabs():
    user_config.colorize_docs = False
    _save_config()
    tabs.init_doc_colorizer(user_config)
    _set_icon(False)


def _print_error(ex):
    output = script.get_output()
    output.print_md("## MENVIC | TAB MANAGER")
    output.print_md("- Failed: {0}".format(str(ex).split("\n")[0]))


def main():
    try:
        if _safe_bool(tabs.get_doc_colorizer_state()):
            _stop_tabs()
        else:
            _enable_tabs()
    except Exception as ex:
        _print_error(ex)


if __name__ == "__main__":
    main()

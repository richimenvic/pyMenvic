# -*- coding: utf-8 -*-

__title__ = "Sort Tabs"
__author__ = "Ricardo J. Mendieta"

from pyrevit.revit import tabs
from pyrevit.userconfig import user_config


def _save_config():
    try:
        user_config.save_changes()
    except:
        pass


def _get_theme():
    try:
        return tabs.get_tabcoloring_theme(user_config)
    except:
        return None


def _set_sort_doc_tabs(theme, state):
    try:
        if theme and hasattr(theme, "SortDocTabs"):
            theme.SortDocTabs = bool(state)
            tabs.set_tabcoloring_theme(user_config, theme)
            return True
    except:
        pass
    return False


def _native_sort_tabs_once():
    theme = _get_theme()
    previous_colorize = False
    previous_sort = False

    try:
        previous_colorize = bool(getattr(user_config, "colorize_docs", False))
    except:
        previous_colorize = False

    try:
        if theme and hasattr(theme, "SortDocTabs"):
            previous_sort = bool(theme.SortDocTabs)
    except:
        previous_sort = False

    try:
        _set_sort_doc_tabs(theme, True)
        user_config.colorize_docs = True
        _save_config()
        tabs.init_doc_colorizer(user_config)
    except:
        pass

    try:
        if not previous_colorize:
            user_config.colorize_docs = False
            _save_config()
            tabs.init_doc_colorizer(user_config)
        else:
            if theme:
                _set_sort_doc_tabs(theme, previous_sort)
                _save_config()
                tabs.init_doc_colorizer(user_config)
    except:
        pass


def main():
    _native_sort_tabs_once()


if __name__ == "__main__":
    main()

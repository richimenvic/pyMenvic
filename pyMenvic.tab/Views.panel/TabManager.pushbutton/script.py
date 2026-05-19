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


def _sort_with_existing_tab_coloring():
    theme = _get_theme()

    # Do not disable pyRevit Tab Coloring. The user may already be using it.
    # This button only asks pyRevit to sort/group document tabs using its native system.
    try:
        _set_sort_doc_tabs(theme, True)
        user_config.colorize_docs = True
        _save_config()
        tabs.init_doc_colorizer(user_config)
    except:
        pass


def main():
    _sort_with_existing_tab_coloring()


if __name__ == "__main__":
    main()

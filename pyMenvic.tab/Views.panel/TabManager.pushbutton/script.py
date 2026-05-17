# -*- coding: utf-8 -*-

__title__ = "Tabs by Document"
__author__ = "Ricardo J. Mendieta"

from pyrevit import HOST_APP, script
from pyrevit.revit import tabs, ui
from pyrevit.runtime import types
from pyrevit.userconfig import user_config
from pyrevit.framework import Media


# --------------------------------------------------
# pyRevit tab coloring
# --------------------------------------------------

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


# --------------------------------------------------
# UI tab sorting helpers
# --------------------------------------------------

def _type_name(value):
    if value is None:
        return ""
    try:
        return value.GetType().FullName
    except:
        return str(type(value))


def _get(obj, prop):
    try:
        return getattr(obj, prop)
    except:
        return None


def _list_items(value):
    items = []
    if value is None:
        return items
    try:
        for item in value:
            items.append(item)
    except:
        pass
    return items


def _index_of(collection, item):
    try:
        return collection.IndexOf(item)
    except:
        pass

    try:
        count = collection.Count
    except:
        return -1

    index = 0
    while index < count:
        try:
            if collection[index] == item:
                return index
        except:
            pass
        index += 1

    return -1


def _doc_key(layout_doc):
    tooltip = _get(layout_doc, "ToolTip")
    if tooltip is None:
        tooltip = ""
    tooltip = str(tooltip)

    if " - " in tooltip:
        return tooltip.split(" - ", 1)[0]

    title = _get(layout_doc, "Title")
    if title is None:
        return ""
    return str(title)


def _find_document_pane_groups(root, limit):
    found = []
    queue = [root]
    visited = 0

    while queue and visited < limit:
        item = queue.pop(0)
        visited += 1

        if "LayoutDocumentPaneGroupControl" in _type_name(item):
            found.append(item)

        try:
            count = Media.VisualTreeHelper.GetChildrenCount(item)
            index = 0
            while index < count:
                child = Media.VisualTreeHelper.GetChild(item, index)
                if child is not None:
                    queue.append(child)
                index += 1
        except:
            pass

    return found


def _get_layout_children():
    api = types.DocumentTabEventUtils
    docking_manager = api.GetDockingManager(HOST_APP.uiapp)

    pane_groups = []
    if docking_manager is not None:
        pane_groups = _find_document_pane_groups(docking_manager, 2500)

    if not pane_groups:
        main_window = ui.get_mainwindow()
        pane_groups = _find_document_pane_groups(main_window, 2500)

    if not pane_groups:
        return None

    panes = api.GetDocumentPanes(pane_groups[0])
    pane_items = _list_items(panes)
    if not pane_items:
        return None

    tab_items = _list_items(api.GetDocumentTabs(pane_items[0]))
    if not tab_items:
        return None

    first_layout = _get(tab_items[0], "Header")
    parent = _get(first_layout, "Parent")
    return _get(parent, "Children")


def _sort_tabs_by_document():
    children = _get_layout_children()
    if children is None:
        return 0

    original = _list_items(children)
    if len(original) < 2:
        return 0

    doc_order = {}
    for item in original:
        key = _doc_key(item)
        if key not in doc_order:
            doc_order[key] = len(doc_order)

    desired = sorted(original, key=lambda x: (doc_order.get(_doc_key(x), 999), original.index(x)))

    moved = 0
    for target_index, item in enumerate(desired):
        current_index = _index_of(children, item)
        if current_index >= 0 and current_index != target_index:
            children.Move(current_index, target_index)
            moved += 1

    return moved


# --------------------------------------------------
# Commands
# --------------------------------------------------

def _enable_tabs():
    theme = _get_theme()
    if theme and hasattr(theme, "SortDocTabs"):
        theme.SortDocTabs = True
        tabs.set_tabcoloring_theme(user_config, theme)

    user_config.colorize_docs = True
    _save_config()
    tabs.init_doc_colorizer(user_config)
    _sort_tabs_by_document()
    _set_icon(True)


def _stop_tabs():
    user_config.colorize_docs = False
    _save_config()
    tabs.init_doc_colorizer(user_config)
    _set_icon(False)


def _print_error(ex):
    output = script.get_output()
    output.print_md("## MENVIC | TABS BY DOCUMENT")
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

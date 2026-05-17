# -*- coding: utf-8 -*-

from pyrevit import HOST_APP
from pyrevit.revit import ui
from pyrevit.runtime import types
from pyrevit.framework import Media


VISUAL_TREE_LIMIT = 2500


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
        pane_groups = _find_document_pane_groups(docking_manager, VISUAL_TREE_LIMIT)
    if not pane_groups:
        main_window = ui.get_mainwindow()
        pane_groups = _find_document_pane_groups(main_window, VISUAL_TREE_LIMIT)
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


def sort_tabs_by_document():
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

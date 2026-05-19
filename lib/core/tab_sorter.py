# -*- coding: utf-8 -*-

from pyrevit import HOST_APP
from pyrevit.revit import ui
from pyrevit.runtime import types
from pyrevit.framework import Media


VISUAL_TREE_LIMIT = 1200
MAX_MOVES_PER_CLICK = 12


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


def _is_true(value):
    try:
        return bool(value)
    except:
        return False


def _is_active_tab(item):
    return _is_true(_get(item, "IsSelected")) or _is_true(_get(item, "IsActive"))


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
    tooltip = str(tooltip).strip()
    if " - " in tooltip:
        return tooltip.split(" - ", 1)[0].strip()
    title = _get(layout_doc, "Title")
    if title is None:
        return ""
    return str(title).strip()


def _find_first_document_pane_group(root, limit):
    if root is None:
        return None

    queue = [root]
    cursor = 0
    visited = 0
    while cursor < len(queue) and visited < limit:
        item = queue[cursor]
        cursor += 1
        visited += 1

        if "LayoutDocumentPaneGroupControl" in _type_name(item):
            return item

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
    return None


def _get_layout_children():
    api = types.DocumentTabEventUtils

    try:
        docking_manager = api.GetDockingManager(HOST_APP.uiapp)
    except:
        docking_manager = None

    pane_group = _find_first_document_pane_group(docking_manager, VISUAL_TREE_LIMIT)
    if pane_group is None:
        pane_group = _find_first_document_pane_group(ui.get_mainwindow(), VISUAL_TREE_LIMIT)
    if pane_group is None:
        return None

    try:
        panes = api.GetDocumentPanes(pane_group)
    except:
        return None

    pane_items = _list_items(panes)
    if not pane_items:
        return None

    try:
        tab_items = _list_items(api.GetDocumentTabs(pane_items[0]))
    except:
        return None

    if not tab_items:
        return None

    first_layout = _get(tab_items[0], "Header")
    parent = _get(first_layout, "Parent")
    return _get(parent, "Children")


def _desired_order(items):
    doc_order = {}
    original_index = {}
    index = 0
    for item in items:
        original_index[item] = index
        key = _doc_key(item)
        if key not in doc_order:
            doc_order[key] = len(doc_order)
        index += 1

    return sorted(
        items,
        key=lambda x: (doc_order.get(_doc_key(x), 999), original_index.get(x, 999999))
    )


def _current_items(children):
    return _list_items(children)


def sort_tabs_by_document():
    children = _get_layout_children()
    if children is None:
        return 0

    original = _current_items(children)
    if len(original) < 2:
        return 0

    desired = _desired_order(original)
    moved = 0

    while moved < MAX_MOVES_PER_CLICK:
        current = _current_items(children)
        if current == desired:
            break

        made_move = False
        target_index = 0
        while target_index < len(desired):
            wanted = desired[target_index]
            if target_index < len(current) and current[target_index] == wanted:
                target_index += 1
                continue

            if _is_active_tab(wanted):
                target_index += 1
                continue

            current_index = _index_of(children, wanted)
            if current_index >= 0 and current_index != target_index:
                try:
                    children.Move(current_index, target_index)
                    moved += 1
                    made_move = True
                except:
                    pass
            break

        if not made_move:
            break

    return moved

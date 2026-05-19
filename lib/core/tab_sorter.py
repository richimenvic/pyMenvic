# -*- coding: utf-8 -*-

from pyrevit import HOST_APP
from pyrevit.revit import ui
from pyrevit.runtime import types
from pyrevit.framework import Media


VISUAL_TREE_LIMIT = 1200
MAX_SORT_PASSES = 8


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


def _find_first_single_move(items):
    keys = []
    for item in items:
        keys.append(_doc_key(item))

    seen_closed = {}
    open_key = None
    index = 0
    for key in keys:
        if key != open_key:
            if open_key is not None:
                seen_closed[open_key] = True
            open_key = key
        if key in seen_closed:
            target = 0
            while target < index and keys[target] == key:
                target += 1
            return index, target
        index += 1
    return None, None


def _move_tabs_around_active(children, original, source_index, target_index):
    active_item = original[source_index]
    moved = 0

    for item in original[target_index:source_index]:
        item_index = _index_of(children, item)
        active_index = _index_of(children, active_item)
        if item_index < 0 or active_index < 0:
            continue
        if item_index > active_index:
            continue
        try:
            children.Move(item_index, active_index + 1)
            moved += 1
        except:
            pass
    return moved


def _sort_one_pass(children):
    original = _list_items(children)
    if len(original) < 2:
        return 0

    source_index, target_index = _find_first_single_move(original)
    if source_index is None or target_index is None:
        return 0
    if source_index == target_index:
        return 0

    item = original[source_index]
    if _is_active_tab(item) and target_index < source_index:
        return _move_tabs_around_active(children, original, source_index, target_index)

    current_index = _index_of(children, item)
    if current_index < 0:
        return 0

    try:
        children.Move(current_index, target_index)
        return 1
    except:
        return 0


def sort_tabs_by_document():
    children = _get_layout_children()
    if children is None:
        return 0

    total_moved = 0
    pass_index = 0
    while pass_index < MAX_SORT_PASSES:
        moved = _sort_one_pass(children)
        if moved <= 0:
            break
        total_moved += moved
        pass_index += 1
    return total_moved

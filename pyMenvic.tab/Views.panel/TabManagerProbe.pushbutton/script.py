# -*- coding: utf-8 -*-

__title__ = "Tab Probe"
__author__ = "Ricardo J. Mendieta"

import os

from pyrevit import HOST_APP, script
from pyrevit.runtime import types
from pyrevit.revit import ui
from pyrevit.framework import Media

try:
    from System import Action
    from System.Windows.Threading import DispatcherPriority
except:
    Action = None
    DispatcherPriority = None


VISUAL_TREE_LIMIT = 2500
HOOK_KEYS = [
    "PYMENVIC_TABS_BY_DOCUMENT_ENABLED",
    "PYMENVIC_TABS_SORT_PENDING",
    "PYMENVIC_TABS_HOOK_HIT",
    "PYMENVIC_TABS_HOOK_SHOULD_SORT",
    "PYMENVIC_TABS_HOOK_IMMEDIATE_MOVES",
    "PYMENVIC_TABS_HOOK_DISPATCHER",
    "PYMENVIC_TABS_HOOK_ERROR",
]


def _err(ex):
    return str(ex).split("\n")[0]


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
    result = []
    if value is None:
        return result
    try:
        for item in value:
            result.append(item)
    except:
        pass
    return result


def _find_pane_groups(root):
    found = []
    queue = [root]
    visited = 0
    while queue and visited < VISUAL_TREE_LIMIT:
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


def _doc_key(layout_doc):
    tooltip = _get(layout_doc, "ToolTip")
    if tooltip is None:
        tooltip = ""
    tooltip = str(tooltip)
    if " - " in tooltip:
        return tooltip.split(" - ", 1)[0]
    title = _get(layout_doc, "Title")
    return str(title) if title is not None else ""


def _title(layout_doc):
    title = _get(layout_doc, "Title")
    return str(title) if title is not None else ""


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


def _get_tab_children(output):
    api = types.DocumentTabEventUtils
    docking_manager = None
    try:
        docking_manager = api.GetDockingManager(HOST_APP.uiapp)
        output.print_md("- `GetDockingManager(uiapp)` OK")
    except Exception as ex:
        output.print_md("- `GetDockingManager(uiapp)` FAILED: `{0}`".format(_err(ex)))

    pane_groups = []
    if docking_manager is not None:
        pane_groups = _find_pane_groups(docking_manager)
    if not pane_groups:
        main_window = ui.get_mainwindow()
        if main_window is not None:
            pane_groups = _find_pane_groups(main_window)
    if not pane_groups:
        output.print_md("- No pane group found.")
        return None

    try:
        panes = api.GetDocumentPanes(pane_groups[0])
        pane_items = _list_items(panes)
    except Exception as ex:
        output.print_md("- `GetDocumentPanes` FAILED: `{0}`".format(_err(ex)))
        return None
    if not pane_items:
        output.print_md("- No document panes found.")
        return None

    try:
        tabs = api.GetDocumentTabs(pane_items[0])
        tab_items = _list_items(tabs)
    except Exception as ex:
        output.print_md("- `GetDocumentTabs` FAILED: `{0}`".format(_err(ex)))
        return None
    if not tab_items:
        output.print_md("- No document tabs found.")
        return None

    first_layout = _get(tab_items[0], "Header")
    parent = _get(first_layout, "Parent")
    children = _get(parent, "Children")
    if children is None:
        output.print_md("- Parent.Children not found.")
        return None
    return children


def _print_order(output, label, children):
    output.print_md("")
    output.print_md("### {0}".format(label))
    items = _list_items(children)
    output.print_md("- Count: `{0}`".format(len(items)))
    for index, item in enumerate(items):
        output.print_md("- `{0}` | key: `{1}` | title: `{2}`".format(index, _doc_key(item), _title(item)))


def _sort_children(output, children, label):
    output.print_md("")
    output.print_md("### {0}".format(label))
    original = _list_items(children)
    if len(original) < 2:
        output.print_md("- Reorder moves: `0`")
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
    output.print_md("- Reorder moves: `{0}`".format(moved))
    return moved


def _dispatcher_sort(output, priority_name):
    output.print_md("")
    output.print_md("### Dispatcher Invoke | {0}".format(priority_name))
    if Action is None or DispatcherPriority is None:
        output.print_md("- Dispatcher types not available.")
        return
    try:
        main_window = ui.get_mainwindow()
        dispatcher = main_window.Dispatcher if main_window is not None else None
        priority = getattr(DispatcherPriority, priority_name)
        if dispatcher is None:
            output.print_md("- Dispatcher not available.")
            return

        def _run():
            children = _get_tab_children(output)
            if children is not None:
                _sort_children(output, children, "Dispatcher pass executed")

        dispatcher.Invoke(priority, Action(_run))
        output.print_md("- Dispatcher Invoke completed.")
    except Exception as ex:
        output.print_md("- Dispatcher Invoke failed: `{0}`".format(_err(ex)))


def _print_hook_state(output):
    output.print_md("")
    output.print_md("### Auto hook state")
    for key in HOOK_KEYS:
        try:
            value = os.environ.get(key, "")
        except:
            value = ""
        output.print_md("- `{0}`: `{1}`".format(key, value))


def _print_active_context(output):
    output.print_md("")
    output.print_md("### Active context")
    output.print_md("- Active document: `{0}`".format(HOST_APP.doc.Title if HOST_APP.doc else "None"))
    try:
        output.print_md("- Active view: `{0}`".format(HOST_APP.active_view.Name if HOST_APP.active_view else "None"))
    except:
        output.print_md("- Active view: `None`")

    try:
        docs = []
        for doc in HOST_APP.app.Documents:
            docs.append(doc.Title)
        output.print_md("- Revit API documents: `{0}`".format(len(docs)))
        for index, title in enumerate(docs):
            output.print_md("  - `{0}` | `{1}`".format(index, title))
    except:
        pass


def main():
    output = script.get_output()
    output.print_md("## MENVIC | TAB MANAGER PROBE")
    output.print_md("Reads tab order, auto-hook state, and tests the same reorder used by the automatic organizer.")
    output.print_md("It does not modify model data. It only reorders Revit UI tabs for testing.")

    _print_hook_state(output)
    _print_active_context(output)

    children = _get_tab_children(output)
    if children is None:
        return

    _print_order(output, "Before manual reorder", children)
    _sort_children(output, children, "Manual reorder")
    _print_order(output, "After manual reorder", children)

    _dispatcher_sort(output, "ApplicationIdle")
    _dispatcher_sort(output, "ContextIdle")
    _dispatcher_sort(output, "Background")

    children = _get_tab_children(output)
    if children is not None:
        _print_order(output, "Final order", children)

    _print_hook_state(output)

    output.print_md("")
    output.print_md("### How to read this")
    output.print_md("- If `PYMENVIC_TABS_HOOK_HIT` is empty, `view-activated.py` is not firing.")
    output.print_md("- If `PYMENVIC_TABS_HOOK_SHOULD_SORT` is `0`, the auto organizer is blocked by settings.")
    output.print_md("- If manual reorder moves tabs but hook moves are `0`, the hook ran before the new tab existed visually.")


if __name__ == "__main__":
    main()

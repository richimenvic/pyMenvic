# -*- coding: utf-8 -*-

__title__ = "Tab Probe"
__author__ = "Ricardo J. Mendieta"

from pyrevit import HOST_APP, script
from pyrevit.runtime import types
from pyrevit.revit import ui
from pyrevit.framework import Media


def _err(ex):
    return str(ex).split("\n")[0]


def _tn(value):
    if value is None:
        return "None"
    try:
        return value.GetType().FullName
    except:
        return str(type(value))


def _call(output, label, func, args):
    try:
        result = func(*args)
        output.print_md("- `{0}` OK".format(label))
        return result
    except Exception as ex:
        output.print_md("- `{0}` FAILED: `{1}`".format(label, _err(ex)))
        return None


def _children(root, limit):
    found = []
    queue = [root]
    visited = 0
    while queue and visited < limit:
        item = queue.pop(0)
        visited += 1
        if "LayoutDocumentPaneGroupControl" in _tn(item):
            found.append(item)
        try:
            count = Media.VisualTreeHelper.GetChildrenCount(item)
            for i in range(count):
                child = Media.VisualTreeHelper.GetChild(item, i)
                if child is not None:
                    queue.append(child)
        except:
            pass
    return found


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


def _get(obj, prop):
    try:
        return getattr(obj, prop)
    except:
        return None


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


def _title(layout_doc):
    title = _get(layout_doc, "Title")
    if title is None:
        return ""
    return str(title)


def _index_of(collection, item):
    try:
        return collection.IndexOf(item)
    except:
        pass
    try:
        count = collection.Count
    except:
        return -1
    i = 0
    while i < count:
        try:
            if collection[i] == item:
                return i
        except:
            pass
        i += 1
    return -1


def _print_order(output, label, items):
    output.print_md("")
    output.print_md("### {0}".format(label))
    for i, item in enumerate(items):
        output.print_md("- `{0}` | `{1}` | `{2}`".format(i, _doc_key(item), _title(item)))


def _sort_children_by_document(output, children):
    original = _list_items(children)
    _print_order(output, "Before", original)

    doc_order = {}
    for item in original:
        key = _doc_key(item)
        if key not in doc_order:
            doc_order[key] = len(doc_order)

    desired = sorted(original, key=lambda x: (doc_order.get(_doc_key(x), 999), original.index(x)))
    _print_order(output, "Target", desired)

    moved = 0
    for target_index, item in enumerate(desired):
        current_index = _index_of(children, item)
        if current_index >= 0 and current_index != target_index:
            children.Move(current_index, target_index)
            moved += 1

    final_items = _list_items(children)
    _print_order(output, "After", final_items)
    output.print_md("")
    output.print_md("- `moves`: `{0}`".format(moved))


def main():
    output = script.get_output()
    api = types.DocumentTabEventUtils

    output.print_md("## MENVIC | TAB WRITE PROBE")
    output.print_md("This test only reorders Revit UI tabs. It does not modify model data.")

    main_window = ui.get_mainwindow()
    docking_manager = _call(output, "GetDockingManager(uiapp)", api.GetDockingManager, [HOST_APP.uiapp])

    pane_groups = []
    if docking_manager is not None:
        pane_groups = _children(docking_manager, 2500)
    if not pane_groups and main_window is not None:
        pane_groups = _children(main_window, 2500)

    if not pane_groups:
        output.print_md("- No pane group found.")
        return

    panes = _call(output, "GetDocumentPanes(pane_group)", api.GetDocumentPanes, [pane_groups[0]])
    pane_items = _list_items(panes)
    if not pane_items:
        output.print_md("- No document pane found.")
        return

    tabs = _call(output, "GetDocumentTabs(pane 0)", api.GetDocumentTabs, [pane_items[0]])
    tab_items = _list_items(tabs)
    if not tab_items:
        output.print_md("- No document tabs found.")
        return

    first_layout = _get(tab_items[0], "Header")
    parent = _get(first_layout, "Parent")
    children = _get(parent, "Children")
    if children is None:
        output.print_md("- Parent.Children not found.")
        return

    try:
        _sort_children_by_document(output, children)
    except Exception as ex:
        output.print_md("- Reorder failed: `{0}`".format(_err(ex)))

    output.print_md("- Active document: `{0}`".format(HOST_APP.doc.Title if HOST_APP.doc else "None"))


if __name__ == "__main__":
    main()

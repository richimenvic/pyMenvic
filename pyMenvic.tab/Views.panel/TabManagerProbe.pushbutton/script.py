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


def _desc(value):
    if value is None:
        return "None"
    text = _tn(value)
    for prop in ["Count", "Name", "Title"]:
        try:
            text = text + " | {0}: {1}".format(prop, getattr(value, prop))
        except:
            pass
    return text


def _call(output, label, func, args):
    try:
        result = func(*args)
        output.print_md("- `{0}` OK: `{1}`".format(label, _desc(result)))
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
    return found, visited


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
    except Exception as ex:
        return "<error: {0}>".format(_err(ex))


def _try_index(collection, index):
    try:
        return collection[index]
    except:
        try:
            return collection.Item[index]
        except:
            try:
                return collection.Item(index)
            except:
                return None


def _print_collection(output, label, collection):
    output.print_md("- `{0}`: `{1}`".format(label, _desc(collection)))
    if collection is None:
        return
    count = 0
    try:
        count = collection.Count
    except:
        try:
            count = len(collection)
        except:
            count = 0
    output.print_md("  - count: `{0}`".format(count))
    i = 0
    while i < count and i < 12:
        item = _try_index(collection, i)
        output.print_md("  - item `{0}`: `{1}`".format(i, _desc(item)))
        i += 1


def _print_layout_brief(output, label, obj):
    output.print_md("- `{0}`: `{1}`".format(label, _desc(obj)))
    for prop in ["Title", "ToolTip", "IsSelected", "IsActive"]:
        output.print_md("  - `{0}`: `{1}`".format(prop, _get(obj, prop)))


def _print_parent_info(output, layout_doc):
    parent = _get(layout_doc, "Parent")
    root = _get(layout_doc, "Root")
    output.print_md("- `Parent`: `{0}`".format(_desc(parent)))
    output.print_md("- `Root`: `{0}`".format(_desc(root)))

    for prop in ["Children", "Items", "Descendents", "FloatingWindows"]:
        value = _get(parent, prop)
        if not str(value).startswith("<error"):
            _print_collection(output, "Parent.{0}".format(prop), value)

    for prop in ["Children", "Items", "Descendents", "RootPanel"]:
        value = _get(root, prop)
        if not str(value).startswith("<error"):
            _print_collection(output, "Root.{0}".format(prop), value)


def main():
    output = script.get_output()
    api = types.DocumentTabEventUtils

    output.print_md("## MENVIC | TAB MANAGER PROBE")
    output.print_md("Read-only runtime probe. No model changes.")

    main_window = ui.get_mainwindow()
    docking_manager = _call(output, "GetDockingManager(uiapp)", api.GetDockingManager, [HOST_APP.uiapp])

    pane_groups = []
    if docking_manager is not None:
        found, visited = _children(docking_manager, 2500)
        output.print_md("- `docking_manager` visited: `{0}` | pane groups: `{1}`".format(visited, len(found)))
        pane_groups.extend(found)
    if not pane_groups and main_window is not None:
        found, visited = _children(main_window, 2500)
        output.print_md("- `main_window` visited: `{0}` | pane groups: `{1}`".format(visited, len(found)))
        pane_groups.extend(found)

    if not pane_groups:
        output.print_md("- No pane group found.")
        return

    pane_group = pane_groups[0]
    panes = _call(output, "GetDocumentPanes(pane_group)", api.GetDocumentPanes, [pane_group])
    pane_items = _list_items(panes)
    output.print_md("- `pane count`: `{0}`".format(len(pane_items)))

    tabs = None
    if pane_items:
        tabs = _call(output, "GetDocumentTabs(pane 0)", api.GetDocumentTabs, [pane_items[0]])

    tab_items = _list_items(tabs)
    output.print_md("- `tab count`: `{0}`".format(len(tab_items)))

    output.print_md("")
    output.print_md("### Tab order and parent collection")
    first_layout = None
    for index, tab in enumerate(tab_items):
        if index >= 12:
            break
        header = _get(tab, "Header")
        title = _get(header, "Title")
        tooltip = _get(header, "ToolTip")
        output.print_md("- `{0}` | Title: `{1}` | ToolTip: `{2}`".format(index, title, tooltip))
        if first_layout is None:
            first_layout = header

    if first_layout is not None:
        output.print_md("")
        output.print_md("### First tab LayoutDocument parent details")
        _print_layout_brief(output, "First LayoutDocument", first_layout)
        _print_parent_info(output, first_layout)

    group = _call(output, "GetDocumentTabGroup(uiapp)", api.GetDocumentTabGroup, [HOST_APP.uiapp])
    output.print_md("- `group`: `{0}`".format(_desc(group)))
    output.print_md("- Active document: `{0}`".format(HOST_APP.doc.Title if HOST_APP.doc else "None"))


if __name__ == "__main__":
    main()

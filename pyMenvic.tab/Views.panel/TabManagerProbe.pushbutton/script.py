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
    try:
        text = text + " | Count: {0}".format(value.Count)
    except:
        pass
    try:
        text = text + " | Name: {0}".format(value.Name)
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
        name = _tn(item)
        if "LayoutDocumentPaneGroupControl" in name:
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


def _print_layout(output, prefix, obj):
    output.print_md("- `{0}` type: `{1}`".format(prefix, _desc(obj)))
    for prop in ["Title", "ContentId", "ToolTip", "Description", "IsSelected", "IsActive", "CanClose", "CanFloat"]:
        output.print_md("  - `{0}`: `{1}`".format(prop, _get(obj, prop)))


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
    output.print_md("- `pane_group`: `{0}`".format(_desc(pane_group)))

    panes = _call(output, "GetDocumentPanes(pane_group)", api.GetDocumentPanes, [pane_group])
    pane_items = _list_items(panes)
    output.print_md("- `pane count`: `{0}`".format(len(pane_items)))

    if pane_items:
        tabs = _call(output, "GetDocumentTabs(pane 0)", api.GetDocumentTabs, [pane_items[0]])
    else:
        tabs = None

    tab_items = _list_items(tabs)
    output.print_md("- `tab count`: `{0}`".format(len(tab_items)))

    index = 0
    for tab in tab_items:
        if index >= 12:
            break
        output.print_md("### Tab {0}".format(index))
        output.print_md("- `tab`: `{0}`".format(_desc(tab)))
        header = _get(tab, "Header")
        content = _get(tab, "Content")
        _print_layout(output, "Header", header)
        _print_layout(output, "Content", content)
        index += 1

    group = _call(output, "GetDocumentTabGroup(uiapp)", api.GetDocumentTabGroup, [HOST_APP.uiapp])
    output.print_md("- `group`: `{0}`".format(_desc(group)))
    output.print_md("- Active document: `{0}`".format(HOST_APP.doc.Title if HOST_APP.doc else "None"))


if __name__ == "__main__":
    main()

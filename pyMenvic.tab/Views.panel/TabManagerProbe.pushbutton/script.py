# -*- coding: utf-8 -*-

__title__ = "Tab Probe"
__author__ = "Ricardo J. Mendieta"

from pyrevit import HOST_APP, script
from pyrevit.runtime import types
from pyrevit.revit import ui
from pyrevit.framework import Media


TARGET_NAMES = [
    "LayoutDocumentPaneGroupControl",
    "LayoutDocumentPaneControl",
    "LayoutDocumentPane",
    "DocumentPane",
    "DocumentPaneTabPanel",
    "LayoutDocumentTabItem",
    "TabPanel",
    "TabItem"
]


def _short_error(ex):
    return str(ex).split("\n")[0]


def _type_name(value):
    if value is None:
        return "None"
    try:
        return value.GetType().FullName
    except:
        return str(type(value))


def _describe(value):
    if value is None:
        return "None"
    text = _type_name(value)
    try:
        text = text + " | Count: {0}".format(value.Count)
    except:
        try:
            text = text + " | len: {0}".format(len(value))
        except:
            pass
    try:
        text = text + " | Name: {0}".format(value.Name)
    except:
        pass
    return text


def _try(output, label, func, args):
    try:
        result = func(*args)
        output.print_md("- `{0}` OK: `{1}`".format(label, _describe(result)))
        return result
    except Exception as ex:
        output.print_md("- `{0}` FAILED: `{1}`".format(label, _short_error(ex)))
        return None


def _items(value):
    result = []
    if value is None:
        return result
    try:
        for item in value:
            result.append(item)
    except:
        pass
    return result


def _child_count(value):
    try:
        return Media.VisualTreeHelper.GetChildrenCount(value)
    except:
        return 0


def _child_at(value, index):
    try:
        return Media.VisualTreeHelper.GetChild(value, index)
    except:
        return None


def _matches_target(value):
    name = _type_name(value)
    for target in TARGET_NAMES:
        if target in name:
            return True
    return False


def _walk_visual_tree(root, max_nodes):
    found = []
    queue = [(root, 0)]
    visited = 0

    while queue and visited < max_nodes:
        item, depth = queue.pop(0)
        visited += 1

        if _matches_target(item):
            found.append((item, depth))

        count = _child_count(item)
        index = 0
        while index < count:
            child = _child_at(item, index)
            if child is not None:
                queue.append((child, depth + 1))
            index += 1

    return found, visited


def _print_found(output, title, found):
    output.print_md("")
    output.print_md("### {0}".format(title))
    if not found:
        output.print_md("- None")
        return

    index = 0
    for item, depth in found:
        if index >= 25:
            break
        output.print_md("- `{0}` depth `{1}` | `{2}`".format(index, depth, _describe(item)))
        index += 1


def _sample(output, title, values):
    output.print_md("")
    output.print_md("### {0}".format(title))
    if not values:
        output.print_md("- None")
        return
    index = 0
    for value in values:
        if index >= 10:
            break
        output.print_md("- `{0}` | `{1}`".format(index, _describe(value)))
        index += 1


def _sample_tabs(output, values):
    output.print_md("")
    output.print_md("### Document tabs")
    if not values:
        output.print_md("- None")
        return
    index = 0
    for tab in values:
        if index >= 20:
            break
        header = None
        content = None
        try:
            header = tab.Header
        except:
            pass
        try:
            content = tab.Content
        except:
            pass
        output.print_md("- `{0}` | `{1}` | Header: `{2}` | Content: `{3}`".format(index, _describe(tab), header, _describe(content)))
        index += 1


def main():
    output = script.get_output()
    obj = types.DocumentTabEventUtils

    output.print_md("## MENVIC | TAB MANAGER PROBE")
    output.print_md("Read-only runtime probe. No model changes.")

    output.print_md("")
    output.print_md("### Base objects")

    main_window = None
    try:
        main_window = ui.get_mainwindow()
        output.print_md("- `ui.get_mainwindow()` OK: `{0}`".format(_describe(main_window)))
    except Exception as ex:
        output.print_md("- `ui.get_mainwindow()` FAILED: `{0}`".format(_short_error(ex)))

    docking_manager = None
    if hasattr(obj, "GetDockingManager"):
        docking_manager = _try(output, "GetDockingManager(uiapp)", obj.GetDockingManager, [HOST_APP.uiapp])

    output.print_md("")
    output.print_md("### Visual tree search")

    roots = []
    if docking_manager is not None:
        roots.append(("docking_manager", docking_manager))
    if main_window is not None:
        roots.append(("main_window", main_window))

    all_found = []
    for label, root in roots:
        found, visited = _walk_visual_tree(root, 2500)
        output.print_md("- `{0}` visited nodes: `{1}` | matches: `{2}`".format(label, visited, len(found)))
        for entry in found:
            all_found.append(entry)
        _print_found(output, "Matches under {0}".format(label), found)

    output.print_md("")
    output.print_md("### Method chain")

    pane_group = None
    for candidate, depth in all_found:
        if "LayoutDocumentPaneGroupControl" in _type_name(candidate):
            pane_group = candidate
            break

    if pane_group is None:
        output.print_md("- `pane_group` not found.")
    else:
        output.print_md("- `pane_group` found: `{0}`".format(_describe(pane_group)))

    panes = None
    if pane_group is not None and hasattr(obj, "GetDocumentPanes"):
        panes = _try(output, "GetDocumentPanes(pane_group)", obj.GetDocumentPanes, [pane_group])

    pane_items = _items(panes)
    _sample(output, "Document panes", pane_items)

    tabs_pane = None
    if pane_group is not None and hasattr(obj, "GetDocumentTabsPane"):
        tabs_pane = _try(output, "GetDocumentTabsPane(pane_group)", obj.GetDocumentTabsPane, [pane_group])

    doc_tabs = None
    if pane_items and hasattr(obj, "GetDocumentTabs"):
        doc_tabs = _try(output, "GetDocumentTabs(pane 0)", obj.GetDocumentTabs, [pane_items[0]])

    tab_items = _items(doc_tabs)
    _sample_tabs(output, tab_items)

    if hasattr(obj, "GetDocumentTabGroup"):
        _try(output, "GetDocumentTabGroup(uiapp)", obj.GetDocumentTabGroup, [HOST_APP.uiapp])
        if tabs_pane is not None:
            _try(output, "GetDocumentTabGroup(tabs_pane)", obj.GetDocumentTabGroup, [tabs_pane])
        for index, tab in enumerate(tab_items):
            if index >= 5:
                break
            _try(output, "GetDocumentTabGroup(tab {0})".format(index), obj.GetDocumentTabGroup, [tab])

    output.print_md("")
    output.print_md("### Context")
    output.print_md("- Active document: `{0}`".format(HOST_APP.doc.Title if HOST_APP.doc else "None"))
    output.print_md("- Active colorizer: `{0}`".format(getattr(obj, "IsUpdatingDocumentTabs", None)))


if __name__ == "__main__":
    main()

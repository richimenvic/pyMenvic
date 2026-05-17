# -*- coding: utf-8 -*-

__title__ = "Tab Probe"
__author__ = "Ricardo J. Mendieta"

from pyrevit import HOST_APP, script
from pyrevit.runtime import types
from pyrevit.revit import ui


def _short_error(ex):
    return str(ex).split("\n")[0]


def _describe(value):
    if value is None:
        return "None"
    text = str(type(value))
    try:
        text = text + " | Count: {0}".format(value.Count)
    except:
        try:
            text = text + " | len: {0}".format(len(value))
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


def main():
    output = script.get_output()
    obj = types.DocumentTabEventUtils

    output.print_md("## MENVIC | TAB MANAGER PROBE")
    output.print_md("Read-only runtime probe. No model changes.")

    output.print_md("")
    output.print_md("### Visible members")
    for name in sorted(dir(obj)):
        if "Tab" in name or "Doc" in name or "Pane" in name or "Group" in name or "Dock" in name:
            output.print_md("- `{0}` | `{1}`".format(name, type(getattr(obj, name))))

    output.print_md("")
    output.print_md("### Argument chain tests")

    main_window = None
    try:
        main_window = ui.get_mainwindow()
        output.print_md("- `ui.get_mainwindow()` OK: `{0}`".format(_describe(main_window)))
    except Exception as ex:
        output.print_md("- `ui.get_mainwindow()` FAILED: `{0}`".format(_short_error(ex)))

    docking_manager = None
    if hasattr(obj, "GetDockingManager"):
        if main_window is not None:
            docking_manager = _try(output, "GetDockingManager(main_window)", obj.GetDockingManager, [main_window])
        if docking_manager is None:
            docking_manager = _try(output, "GetDockingManager(uiapp)", obj.GetDockingManager, [HOST_APP.uiapp])

    panes = None
    if docking_manager is not None and hasattr(obj, "GetDocumentPanes"):
        panes = _try(output, "GetDocumentPanes(docking_manager)", obj.GetDocumentPanes, [docking_manager])

    pane_items = _items(panes)
    _sample(output, "Document panes", pane_items)

    tabs_pane = None
    if hasattr(obj, "GetDocumentTabsPane"):
        for index, pane in enumerate(pane_items):
            result = _try(output, "GetDocumentTabsPane(pane {0})".format(index), obj.GetDocumentTabsPane, [pane])
            if result is not None and tabs_pane is None:
                tabs_pane = result

    doc_tabs = None
    if tabs_pane is not None and hasattr(obj, "GetDocumentTabs"):
        doc_tabs = _try(output, "GetDocumentTabs(tabs_pane)", obj.GetDocumentTabs, [tabs_pane])

    tab_items = _items(doc_tabs)
    _sample(output, "Document tabs", tab_items)

    if hasattr(obj, "GetDocumentTabGroup"):
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

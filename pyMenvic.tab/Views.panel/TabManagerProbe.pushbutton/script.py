# -*- coding: utf-8 -*-

__title__ = "Tab Probe"
__author__ = "Ricardo J. Mendieta"

from pyrevit import HOST_APP, script
from pyrevit.runtime import types


KEYWORDS = [
    "Tab",
    "Group",
    "Doc",
    "Pane",
    "Sort"
]


def _matches(name):
    for keyword in KEYWORDS:
        if keyword in name:
            return True
    return False


def _short_error(ex):
    return str(ex).split("\n")[0]


def _describe_value(value):
    if value is None:
        return "None"

    value_type = str(type(value))

    try:
        count = len(value)
        return "{0} | len: {1}".format(value_type, count)
    except:
        pass

    try:
        count_prop = getattr(value, "Count")
        return "{0} | Count: {1}".format(value_type, count_prop)
    except:
        pass

    return value_type


def _print_member(output, obj, name):
    try:
        value = getattr(obj, name)
        output.print_md("- `{0}` | `{1}`".format(name, type(value)))
    except Exception as ex:
        output.print_md("- `{0}` | error: `{1}`".format(name, _short_error(ex)))


def _try_call(output, label, func, args):
    try:
        result = func(*args)
        output.print_md("- `{0}` OK: `{1}`".format(label, _describe_value(result)))
        return result
    except Exception as ex:
        output.print_md("- `{0}` FAILED: `{1}`".format(label, _short_error(ex)))
        return None


def _print_tab_sample(output, tabs_list):
    if tabs_list is None:
        return

    output.print_md("")
    output.print_md("### Tab sample")

    index = 0
    try:
        for tab in tabs_list:
            if index >= 10:
                break
            title = getattr(tab, "Title", None)
            if title is None:
                title = getattr(tab, "Header", None)
            doc_id = getattr(tab, "DocumentId", None)
            view_id = getattr(tab, "ViewId", None)
            output.print_md("- `{0}` | Title: `{1}` | DocumentId: `{2}` | ViewId: `{3}`".format(index, title, doc_id, view_id))
            index += 1
    except Exception as ex:
        output.print_md("- Could not enumerate tabs: `{0}`".format(_short_error(ex)))


def main():
    output = script.get_output()
    output.print_md("## MENVIC | TAB MANAGER PROBE")
    output.print_md("Read-only runtime probe. No model changes.")
    output.print_md("")

    obj = types.DocumentTabEventUtils

    output.print_md("### Visible members")
    names = []
    try:
        for name in dir(obj):
            if _matches(name):
                names.append(name)
    except Exception as ex:
        output.print_md("- Failed reading members: `{0}`".format(_short_error(ex)))
        return

    if not names:
        output.print_md("- No matching members found.")
        return

    for name in sorted(names):
        _print_member(output, obj, name)

    output.print_md("")
    output.print_md("### Direct checks")
    for name in [
        "UpdateDocumentTabGroups",
        "ClearDocumentTabGroups",
        "GetDocumentTabs",
        "GetDocumentTabsPane",
        "GetDocumentTabGroup",
        "StartGroupingDocumentTabs",
        "StopGroupingDocumentTabs",
        "ResetGroupingDocumentTabs",
        "IsUpdatingDocumentTabs",
        "TabColoringTheme"
    ]:
        output.print_md("- `{0}` exists: `{1}`".format(name, hasattr(obj, name)))

    output.print_md("")
    output.print_md("### Safe call tests")

    tabs_pane = None
    doc_tabs = None

    if hasattr(obj, "GetDocumentTabsPane"):
        tabs_pane = _try_call(output, "GetDocumentTabsPane()", obj.GetDocumentTabsPane, [])

    if hasattr(obj, "GetDocumentTabs"):
        doc_tabs = _try_call(output, "GetDocumentTabs()", obj.GetDocumentTabs, [])
        if doc_tabs is None and tabs_pane is not None:
            doc_tabs = _try_call(output, "GetDocumentTabs(tabs_pane)", obj.GetDocumentTabs, [tabs_pane])

    if hasattr(obj, "GetDocumentTabGroup"):
        _try_call(output, "GetDocumentTabGroup()", obj.GetDocumentTabGroup, [])
        if tabs_pane is not None:
            _try_call(output, "GetDocumentTabGroup(tabs_pane)", obj.GetDocumentTabGroup, [tabs_pane])
        if doc_tabs is not None:
            try:
                first_tab = None
                for tab in doc_tabs:
                    first_tab = tab
                    break
                if first_tab is not None:
                    _try_call(output, "GetDocumentTabGroup(first_tab)", obj.GetDocumentTabGroup, [first_tab])
            except Exception as ex:
                output.print_md("- `GetDocumentTabGroup(first_tab)` prep FAILED: `{0}`".format(_short_error(ex)))

    _print_tab_sample(output, doc_tabs)

    output.print_md("")
    output.print_md("### Context")
    output.print_md("- Active document: `{0}`".format(HOST_APP.doc.Title if HOST_APP.doc else "None"))
    output.print_md("- Active colorizer: `{0}`".format(getattr(obj, "IsUpdatingDocumentTabs", None)))


if __name__ == "__main__":
    main()

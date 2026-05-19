# -*- coding: utf-8 -*-

__title__ = "Tab Probe"
__author__ = "Ricardo J. Mendieta"

from pyrevit import HOST_APP, script
from pyrevit.runtime import types
from pyrevit.revit import ui
from pyrevit.framework import Media

try:
    from System.Threading import Thread
except:
    Thread = None


PROBE_DELAYS_MS = [250, 500, 1000]


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


def _sleep(ms):
    if Thread is None:
        return
    try:
        Thread.Sleep(ms)
    except:
        pass


def _get_open_documents():
    titles = []
    try:
        for doc in HOST_APP.app.Documents:
            try:
                titles.append(doc.Title)
            except:
                titles.append("<unknown>")
    except:
        pass
    return titles


def _print_open_documents(output):
    docs = _get_open_documents()
    output.print_md("")
    output.print_md("### Revit API documents")
    output.print_md("- Count: `{0}`".format(len(docs)))
    for i, title in enumerate(docs):
        output.print_md("- `{0}` | `{1}`".format(i, title))


def _get_tab_context(output, api, verbose):
    main_window = ui.get_mainwindow()
    docking_manager = None
    if verbose:
        docking_manager = _call(output, "GetDockingManager(uiapp)", api.GetDockingManager, [HOST_APP.uiapp])
    else:
        try:
            docking_manager = api.GetDockingManager(HOST_APP.uiapp)
        except:
            docking_manager = None

    pane_groups = []
    if docking_manager is not None:
        pane_groups = _children(docking_manager, 2500)
    if not pane_groups and main_window is not None:
        pane_groups = _children(main_window, 2500)

    if not pane_groups:
        return None

    if verbose:
        panes = _call(output, "GetDocumentPanes(pane_group)", api.GetDocumentPanes, [pane_groups[0]])
    else:
        try:
            panes = api.GetDocumentPanes(pane_groups[0])
        except:
            panes = None
    pane_items = _list_items(panes)
    if not pane_items:
        return None

    if verbose:
        tabs = _call(output, "GetDocumentTabs(pane 0)", api.GetDocumentTabs, [pane_items[0]])
    else:
        try:
            tabs = api.GetDocumentTabs(pane_items[0])
        except:
            tabs = None
    tab_items = _list_items(tabs)
    if not tab_items:
        return None

    first_layout = _get(tab_items[0], "Header")
    parent = _get(first_layout, "Parent")
    children = _get(parent, "Children")
    if children is None:
        return None

    return {
        "pane_groups": pane_groups,
        "panes": pane_items,
        "tabs": tab_items,
        "children": children,
    }


def _print_items(output, title, items):
    output.print_md("#### {0}".format(title))
    output.print_md("- Count: `{0}`".format(len(items)))
    for i, item in enumerate(items):
        output.print_md("- `{0}` | key: `{1}` | title: `{2}`".format(i, _doc_key(item), _title(item)))


def _snapshot(output, api, label):
    output.print_md("")
    output.print_md("### {0}".format(label))

    context = _get_tab_context(output, api, False)
    if context is None:
        output.print_md("- Tab context not available.")
        return None

    output.print_md("- Pane groups: `{0}`".format(len(context["pane_groups"])))
    output.print_md("- Panes: `{0}`".format(len(context["panes"])))
    _print_items(output, "DocumentTabEventUtils.GetDocumentTabs", context["tabs"])
    visual_items = _list_items(context["children"])
    _print_items(output, "Visual parent.Children order", visual_items)
    return context


def _sort_children_by_document(output, children):
    original = _list_items(children)
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


def _reorder_current_context(output, api, label):
    output.print_md("")
    output.print_md("### {0}".format(label))
    context = _get_tab_context(output, api, False)
    if context is None:
        output.print_md("- Tab context not available. Reorder skipped.")
        return 0
    try:
        return _sort_children_by_document(output, context["children"])
    except Exception as ex:
        output.print_md("- Reorder failed: `{0}`".format(_err(ex)))
        return 0


def main():
    output = script.get_output()
    api = types.DocumentTabEventUtils

    output.print_md("## MENVIC | TAB MANAGER TIMING PROBE")
    output.print_md("This diagnostic reads Revit tab state, performs two controlled reorder passes, and compares immediate/delayed tab lists.")
    output.print_md("It does not modify model data. It only tests Revit UI tab ordering.")

    output.print_md("")
    output.print_md("### Active context")
    output.print_md("- Active document: `{0}`".format(HOST_APP.doc.Title if HOST_APP.doc else "None"))
    try:
        output.print_md("- Active view: `{0}`".format(HOST_APP.active_view.Name if HOST_APP.active_view else "None"))
    except:
        output.print_md("- Active view: `None`")

    _print_open_documents(output)

    initial_context = _get_tab_context(output, api, True)
    if initial_context is None:
        output.print_md("")
        output.print_md("- No document tabs found.")
        return

    _snapshot(output, api, "T+0 ms | Before first reorder")
    _reorder_current_context(output, api, "Pass 1 | Immediate reorder")
    _snapshot(output, api, "T+0 ms | After first reorder")

    elapsed = 0
    for delay in PROBE_DELAYS_MS:
        _sleep(delay)
        elapsed += delay
        _snapshot(output, api, "T+{0} ms | Delayed snapshot".format(elapsed))

    _reorder_current_context(output, api, "Pass 2 | Delayed reorder")
    _snapshot(output, api, "After delayed reorder")

    output.print_md("")
    output.print_md("### How to read this")
    output.print_md("- If the newest tab is missing at T+0 but appears in delayed snapshots, the issue is timing.")
    output.print_md("- If GetDocumentTabs and parent.Children disagree, the issue is between Revit API state and visual WPF state.")
    output.print_md("- If Pass 2 fixes the last tab, the final organizer should use a delayed second pass, not a heavy watcher.")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-

__title__ = "Tab Probe"
__author__ = "Ricardo J. Mendieta"

from pyrevit import script
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


def _print_member(output, obj, name):
    try:
        value = getattr(obj, name)
        output.print_md("- `{0}` | `{1}`".format(name, type(value)))
    except Exception as ex:
        output.print_md("- `{0}` | error: `{1}`".format(name, str(ex).split("\n")[0]))


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
        output.print_md("- Failed reading members: `{0}`".format(str(ex).split("\n")[0]))
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


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-

from pyrevit import revit, DB, forms, script

__title__ = "SYNC LINK WORKSETS"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
SYNC LINK WORKSETS
_____________________________________________________

Synchronizes host worksets with worksets detected
in loaded Revit links.

The tool can:

- List worksets detected in links
- Create missing worksets in the host
- Detect unused host worksets
- Delete unused host worksets
"""

doc = revit.doc
output = script.get_output()

MODE_RUN_CREATE = "RUN: CREATE MISSING WORKSETS"
MODE_PROBE_LINKS = "PROBE: LIST LINK WORKSETS"
MODE_PROBE_UNUSED = "PROBE: LIST UNUSED HOST WORKSETS"
MODE_RUN_DELETE = "RUN: DELETE UNUSED HOST WORKSETS"
MODE_CANCEL = "CANCEL"

modes = [
    MODE_RUN_CREATE,
    MODE_PROBE_LINKS,
    MODE_PROBE_UNUSED,
    MODE_RUN_DELETE,
    MODE_CANCEL
]

switch = forms.CommandSwitchWindow.show(
    modes,
    message="Selecciona modo"
)

if not switch or switch == MODE_CANCEL:
    script.exit()

if not doc.IsWorkshared:
    forms.alert("El modelo no es workshared.", exitscript=True)


def safe_name(name):
    if not name:
        return ""
    return str(name).strip()


def normalize(name):
    return safe_name(name).lower()


def get_host_worksets():
    result = {}
    collector = DB.FilteredWorksetCollector(doc)\
        .OfKind(DB.WorksetKind.UserWorkset)

    for ws in collector:
        result[normalize(ws.Name)] = ws

    return result


def get_link_instances():
    return DB.FilteredElementCollector(doc)\
        .OfClass(DB.RevitLinkInstance)


def get_link_doc(link):
    try:
        return link.GetLinkDocument()
    except:
        return None


def get_link_worksets(link_doc):
    result = []
    collector = DB.FilteredWorksetCollector(link_doc)\
        .OfKind(DB.WorksetKind.UserWorkset)

    for ws in collector:
        result.append(ws)

    return result


host_worksets = get_host_worksets()
candidate_link_worksets = {}

skipped_links = []
scanned_links = 0

for link in get_link_instances():

    link_name = safe_name(link.Name)
    link_doc = get_link_doc(link)

    if not link_doc:
        skipped_links.append(link_name)
        continue

    scanned_links += 1

    for ws in get_link_worksets(link_doc):

        ws_name = safe_name(ws.Name)
        ws_norm = normalize(ws_name)

        if not ws_name:
            continue

        if ws_norm not in candidate_link_worksets:
            candidate_link_worksets[ws_norm] = {
                "name": ws_name,
                "link": link_name
            }


names_to_create = []

for norm in sorted(candidate_link_worksets):

    if norm not in host_worksets:
        names_to_create.append(candidate_link_worksets[norm])


if switch == MODE_PROBE_LINKS:

    output.print_md("# Worksets detectados en links")
    output.print_md("| Link | Workset | Estado |")
    output.print_md("|---|---|---|")

    for ws in candidate_link_worksets.values():

        name = ws["name"]

        if normalize(name) in host_worksets:
            state = "Ya existe"
        else:
            state = "Falta crear"

        output.print_md(
            "| {} | {} | {} |".format(
                ws["link"],
                name,
                state
            )
        )

    script.exit()


if switch == MODE_RUN_CREATE:

    if not names_to_create:

        if scanned_links == 0:
            forms.alert(
                "No se pudieron leer los links cargados.",
                exitscript=True
            )

        forms.alert(
            "No hay worksets faltantes por crear.",
            exitscript=True
        )

    selection_items = [x["name"] for x in names_to_create]

    selected = forms.SelectFromList.show(
        selection_items,
        title="Selecciona los worksets a crear",
        multiselect=True,
        button_name="Crear seleccionados"
    )

    if not selected:
        script.exit()

    tx = DB.Transaction(doc, "Create Worksets")
    tx.Start()

    created = 0

    for item in names_to_create:

        if item["name"] not in selected:
            continue

        try:
            DB.Workset.Create(doc, item["name"])
            created += 1
        except:
            pass

    tx.Commit()

    forms.alert("{} worksets creados.".format(created))


if switch == MODE_PROBE_UNUSED:

    output.print_md("# Worksets host")

    collector = DB.FilteredWorksetCollector(doc)\
        .OfKind(DB.WorksetKind.UserWorkset)

    for ws in collector:

        output.print_md(
            "- {}".format(ws.Name)
        )


if switch == MODE_RUN_DELETE:

    forms.alert(
        "Delete worksets no implementado en esta version."
    )


output.print_md("")
output.print_md("Links escaneados: {}".format(scanned_links))
output.print_md("Links omitidos: {}".format(len(skipped_links)))
# -*- coding: utf-8 -*-

__title__ = "Audit Link Worksets"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
AUDIT LINK WORKSETS
_____________________________________________________

Description:

Audits the worksets present in the host model and
compares them with the worksets detected in all
loaded Revit links.

This tool does not modify the model. It only creates
a detailed report.

_____________________________________________________
What the tool reports:

• Worksets present in the host model
• Worksets detected in linked models
• Worksets that exist only in the host
• Worksets that exist only in links
• Number of links using each workset

_____________________________________________________
Output:

The report includes:

• Global summary
• Table of all detected worksets
• Worksets present only in host
• Worksets present only in links
• Errors detected during scanning

_____________________________________________________
Usage:

1. Click the button.
2. Select "REPORT: AUDIT LINK WORKSETS".
3. The tool scans all loaded links.
4. The report is printed in the pyRevit output window.

_____________________________________________________
Notes:

• The tool is read-only.
• No worksets are created or deleted.
• Only User Worksets are analyzed.

_____________________________________________________
Author: Ricardo J. Mendieta
"""

from pyrevit import revit, DB, forms, script


# ==================================================
# CONFIG
# ==================================================

TOOL_NAME = "AUDIT LINK WORKSETS"

MODE_REPORT = "REPORT: AUDIT LINK WORKSETS"
MODE_CANCEL = "CANCEL"

PROTECTED_PREFIXES = ["<"]
PROTECTED_CONTAINS = ["<", ">"]


# ==================================================
# HELPERS
# ==================================================

def element_id_value(element_id, default=-1):
    if element_id is None:
        return default
    try:
        return element_id.Value
    except Exception:
        pass
    try:
        return element_id.IntegerValue
    except Exception:
        pass
    return default


output = script.get_output()
doc = revit.doc


def first_line(ex):
    try:
        return str(ex).splitlines()[0]
    except Exception:
        try:
            return str(ex)
        except Exception:
            return "Unknown error"


def safe_str(value):
    if value is None:
        return ""
    try:
        return str(value)
    except Exception:
        return ""


def safe_name(name):
    if name is None:
        return ""
    try:
        return name.strip()
    except Exception:
        try:
            return str(name).strip()
        except Exception:
            return ""


def normalize_name(name):
    text = safe_name(name)
    if not text:
        return ""
    return text.lower()


def is_protected_name(name):
    text = safe_name(name)
    if not text:
        return True

    for prefix in PROTECTED_PREFIXES:
        if text.startswith(prefix):
            return True

    for token in PROTECTED_CONTAINS:
        if token in text:
            return True

    return False


def get_host_user_worksets(host_doc):
    results = []
    collector = DB.FilteredWorksetCollector(host_doc).OfKind(DB.WorksetKind.UserWorkset)
    for ws in collector:
        results.append(ws)
    return results


def get_host_workset_map(host_doc):
    data = {}

    for ws in get_host_user_worksets(host_doc):
        try:
            ws_name = safe_name(ws.Name)
            ws_norm = normalize_name(ws_name)
            if not ws_name or not ws_norm:
                continue
            if is_protected_name(ws_name):
                continue

            if ws_norm not in data:
                data[ws_norm] = {
                    "name": ws_name,
                    "id": element_id_value(ws.Id)
                }
        except Exception:
            pass

    return data


def get_loaded_link_instances(host_doc):
    results = []
    collector = DB.FilteredElementCollector(host_doc).OfClass(DB.RevitLinkInstance)
    for inst in collector:
        results.append(inst)
    return results


def get_link_document(link_instance):
    try:
        return link_instance.GetLinkDocument()
    except Exception:
        return None


def get_link_display_name(link_instance):
    try:
        if link_instance.Name:
            return safe_name(link_instance.Name)
    except Exception:
        pass

    try:
        link_type = doc.GetElement(link_instance.GetTypeId())
        if link_type and link_type.Name:
            return safe_name(link_type.Name)
    except Exception:
        pass

    try:
        return "Link Id {}".format(element_id_value(link_instance.Id))
    except Exception:
        return "Unknown Link"


def get_user_worksets_from_link(link_doc):
    results = []
    collector = DB.FilteredWorksetCollector(link_doc).OfKind(DB.WorksetKind.UserWorkset)
    for ws in collector:
        results.append(ws)
    return results


def join_links(link_names):
    if not link_names:
        return "-"

    sorted_names = sorted(link_names)
    return ", ".join(sorted_names)


def print_fail_table(failures):
    if not failures:
        return

    output.print_md("### Fallos relevantes")
    output.print_md("")
    output.print_md("| Elemento | Acción | Motivo |")
    output.print_md("| --- | --- | --- |")

    for item in failures:
        output.print_md("| {} | {} | {} |".format(
            item.get("elemento", ""),
            item.get("accion", ""),
            item.get("motivo", "")
        ))


def print_main_table(rows):
    output.print_md("### Informe por workset")
    output.print_md("")

    if not rows:
        output.print_md("✅ No se detectaron worksets de usuario en links ni en host.")
        output.print_md("")
        return

    output.print_md("| Workset | Existe en host | Nº links | Links |")
    output.print_md("| --- | --- | --- | --- |")

    for row in rows:
        output.print_md("| {} | {} | {} | {} |".format(
            row.get("workset", ""),
            row.get("host", ""),
            row.get("num_links", 0),
            row.get("links", "")
        ))

    output.print_md("")


def print_host_only_table(rows):
    output.print_md("### Solo en host")
    output.print_md("")

    if not rows:
        output.print_md("✅ Ninguno.")
        output.print_md("")
        return

    output.print_md("| Workset |")
    output.print_md("| --- |")

    for row in rows:
        output.print_md("| {} |".format(row.get("workset", "")))

    output.print_md("")


def print_links_only_table(rows):
    output.print_md("### Solo en links y no en host")
    output.print_md("")

    if not rows:
        output.print_md("✅ Ninguno.")
        output.print_md("")
        return

    output.print_md("| Workset | Nº links | Links |")
    output.print_md("| --- | --- | --- |")

    for row in rows:
        output.print_md("| {} | {} | {} |".format(
            row.get("workset", ""),
            row.get("num_links", 0),
            row.get("links", "")
        ))

    output.print_md("")


# ==================================================
# COLLECT / SCAN
# ==================================================

switch = forms.CommandSwitchWindow.show(
    [MODE_REPORT, MODE_CANCEL],
    message="Selecciona modo"
)

if not switch or switch == MODE_CANCEL:
    script.exit()

failures = []
scanned_links = 0
skipped_links = []

host_worksets = get_host_workset_map(doc)
host_norm_names = set(host_worksets.keys())

# normalized workset name -> {name: display_name, links: set([...])}
link_workset_map = {}

link_instances = get_loaded_link_instances(doc)

for link_inst in link_instances:
    link_name = get_link_display_name(link_inst)
    link_doc = get_link_document(link_inst)

    if link_doc is None:
        skipped_links.append(link_name)
        continue

    scanned_links += 1

    try:
        link_worksets = get_user_worksets_from_link(link_doc)
    except Exception as ex:
        failures.append({
            "elemento": link_name,
            "accion": "Scan Link Worksets",
            "motivo": first_line(ex)
        })
        continue

    for ws in link_worksets:
        try:
            ws_name = safe_name(ws.Name)
            ws_norm = normalize_name(ws_name)
        except Exception as ex:
            failures.append({
                "elemento": link_name,
                "accion": "Read Workset Name",
                "motivo": first_line(ex)
            })
            continue

        if not ws_name or not ws_norm:
            continue

        if is_protected_name(ws_name):
            continue

        if ws_norm not in link_workset_map:
            link_workset_map[ws_norm] = {
                "name": ws_name,
                "links": set()
            }

        link_workset_map[ws_norm]["links"].add(link_name)


# ==================================================
# PROCESS
# ==================================================

all_norm_names = set()
for key in host_worksets.keys():
    all_norm_names.add(key)
for key in link_workset_map.keys():
    all_norm_names.add(key)

main_rows = []
host_only_rows = []
links_only_rows = []

for ws_norm in sorted(all_norm_names):
    in_host = ws_norm in host_norm_names
    in_links = ws_norm in link_workset_map

    if in_links:
        ws_name = link_workset_map[ws_norm]["name"]
        link_names = link_workset_map[ws_norm]["links"]
    else:
        ws_name = host_worksets[ws_norm]["name"]
        link_names = set()

    row = {
        "workset": ws_name,
        "host": "Sí" if in_host else "No",
        "num_links": len(link_names),
        "links": join_links(link_names)
    }

    main_rows.append(row)

    if in_host and not in_links:
        host_only_rows.append({
            "workset": ws_name
        })

    if in_links and not in_host:
        links_only_rows.append({
            "workset": ws_name,
            "num_links": len(link_names),
            "links": join_links(link_names)
        })


# ==================================================
# CLEANUP / PURGE
# ==================================================

# No cleanup needed for this read-only tool.


# ==================================================
# REPORT
# ==================================================

output.print_md("# MENVIC | {} — REPORT".format(TOOL_NAME))
output.print_md("")
output.print_md("## Resumen")
output.print_md("")
output.print_md("* Worksets únicos detectados: {}".format(len(all_norm_names)))
output.print_md("* Links cargados escaneados: {}".format(scanned_links))
output.print_md("* Links omitidos (sin documento cargado): {}".format(len(skipped_links)))
output.print_md("* Worksets presentes en host: {}".format(len(host_worksets)))
output.print_md("* Worksets detectados en links: {}".format(len(link_workset_map)))
output.print_md("* Solo en host: {}".format(len(host_only_rows)))
output.print_md("* Solo en links y no en host: {}".format(len(links_only_rows)))
output.print_md("* Fallos: {}".format(len(failures)))
output.print_md("")

if skipped_links:
    output.print_md("### Links omitidos")
    output.print_md("")
    for name in skipped_links:
        output.print_md("* {}".format(name))
    output.print_md("")

print_main_table(main_rows)
print_host_only_table(host_only_rows)
print_links_only_table(links_only_rows)
print_fail_table(failures)

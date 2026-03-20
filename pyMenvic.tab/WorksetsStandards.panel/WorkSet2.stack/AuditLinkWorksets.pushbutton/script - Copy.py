# -*- coding: utf-8 -*-

from pyrevit import revit, DB, forms, script


# ==================================================
# CONFIG
# ==================================================

TOOL_NAME = "SYNC LINK WORKSETS"

MODE_RUN_CREATE = "RUN: CREATE MISSING WORKSETS"
MODE_PROBE_LINKS = "PROBE: LIST LINK WORKSETS"
MODE_PROBE_UNUSED = "PROBE: LIST UNUSED HOST WORKSETS"
MODE_RUN_DELETE = "RUN: DELETE UNUSED HOST WORKSETS"
MODE_CANCEL = "CANCEL"

PROTECTED_PREFIXES = ["<"]
PROTECTED_CONTAINS = ["<", ">"]


# ==================================================
# HELPERS
# ==================================================

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


def has_delete_workset_api():
    try:
        if not hasattr(DB, "DeleteWorksetSettings"):
            return False
        if not hasattr(DB, "DeleteWorksetOption"):
            return False
        if not hasattr(DB.WorksetTable, "CanDeleteWorkset"):
            return False
        if not hasattr(DB.WorksetTable, "DeleteWorkset"):
            return False
        return True
    except Exception:
        return False


def get_host_user_worksets(host_doc):
    results = []
    collector = DB.FilteredWorksetCollector(host_doc).OfKind(DB.WorksetKind.UserWorkset)
    for ws in collector:
        results.append(ws)
    return results


def get_host_user_workset_maps(host_doc):
    by_norm = {}
    by_id_int = {}

    worksets = get_host_user_worksets(host_doc)
    for ws in worksets:
        try:
            ws_name = safe_name(ws.Name)
            norm = normalize_name(ws_name)
            if norm and norm not in by_norm:
                by_norm[norm] = ws
            by_id_int[ws.Id.IntegerValue] = ws
        except Exception:
            pass

    return by_norm, by_id_int


def get_loaded_link_instances(host_doc):
    links = []
    collector = DB.FilteredElementCollector(host_doc).OfClass(DB.RevitLinkInstance)
    for inst in collector:
        links.append(inst)
    return links


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
        return "Link Id {}".format(link_instance.Id.IntegerValue)
    except Exception:
        return "Unknown Link"


def get_user_worksets_from_link(link_doc):
    results = []
    collector = DB.FilteredWorksetCollector(link_doc).OfKind(DB.WorksetKind.UserWorkset)
    for ws in collector:
        results.append(ws)
    return results


def get_active_workset_id(host_doc):
    try:
        table = host_doc.GetWorksetTable()
        return table.GetActiveWorksetId()
    except Exception:
        return None


def get_default_user_workset(host_doc):
    worksets = get_host_user_worksets(host_doc)

    for ws in worksets:
        try:
            if ws.IsDefaultWorkset:
                return ws
        except Exception:
            pass

    if worksets:
        return worksets[0]

    return None


def workset_has_any_instance_elements(host_doc, workset_id):
    try:
        ws_filter = DB.ElementWorksetFilter(workset_id)
        collector = DB.FilteredElementCollector(host_doc).WherePasses(ws_filter).WhereElementIsNotElementType()

        for _el in collector:
            return True

        return False
    except Exception:
        return True


def get_workset_owner_text(ws):
    try:
        return safe_str(ws.Owner)
    except Exception:
        return ""


def is_workset_deletable(host_doc, ws, delete_settings):
    try:
        return DB.WorksetTable.CanDeleteWorkset(host_doc, ws.Id, delete_settings)
    except Exception:
        return False


def can_create_workset_name(host_doc, ws_name):
    try:
        return DB.WorksetTable.IsWorksetNameUnique(host_doc, ws_name)
    except Exception:
        return True


def build_delete_settings(default_target_workset):
    if default_target_workset is None:
        return None

    try:
        return DB.DeleteWorksetSettings(
            DB.DeleteWorksetOption.MoveElementsToWorkset,
            default_target_workset.Id
        )
    except Exception:
        return None


def get_unused_workset_status(host_doc, ws, candidate_link_worksets, active_workset_id, delete_settings):
    ws_name = safe_name(ws.Name)
    ws_norm = normalize_name(ws_name)

    if not ws_name or not ws_norm:
        return "SKIP", "Nombre inválido"

    if is_protected_name(ws_name):
        return "SKIP", "Protegido"

    try:
        if ws.IsDefaultWorkset:
            return "NO", "Es workset por defecto"
    except Exception:
        pass

    try:
        if active_workset_id and ws.Id.IntegerValue == active_workset_id.IntegerValue:
            return "NO", "Es workset activo"
    except Exception:
        pass

    if ws_norm in candidate_link_worksets:
        return "NO", "Presente en links"

    has_elements = workset_has_any_instance_elements(host_doc, ws.Id)
    if has_elements:
        return "NO", "Tiene elementos"

    deletable = is_workset_deletable(host_doc, ws, delete_settings)
    if deletable:
        return "YES", "Vacío y borrable"

    owner_text = get_workset_owner_text(ws)
    if owner_text:
        return "NO", "Vacío pero no borrable por Revit (owner: {})".format(owner_text)

    return "NO", "Vacío pero no borrable por Revit"


def print_fail_table(failures):
    if not failures:
        return

    output.print_md("### Fallos relevantes")
    output.print_md("")
    output.print_md("| Elemento | Acción | Motivo |")
    output.print_md("| --- | --- | --- |")

    for item in failures:
        elem = item.get("elemento", "")
        action = item.get("accion", "")
        reason = item.get("motivo", "")
        output.print_md("| {} | {} | {} |".format(elem, action, reason))


def print_probe_links_table(rows):
    if not rows:
        output.print_md("✅ No se encontraron worksets de usuario en los vínculos cargados.")
        return

    output.print_md("### Worksets detectados en vínculos cargados")
    output.print_md("")
    output.print_md("| Link | Workset | Estado en host |")
    output.print_md("| --- | --- | --- |")

    for row in rows:
        output.print_md("| {} | {} | {} |".format(
            row.get("link", ""),
            row.get("workset", ""),
            row.get("estado", "")
        ))


def print_probe_unused_table(rows):
    if not rows:
        output.print_md("✅ No se detectaron worksets host candidatos.")
        return

    output.print_md("### Diagnóstico de worksets host")
    output.print_md("")
    output.print_md("| Workset | Borrable | Estado |")
    output.print_md("| --- | --- | --- |")

    for row in rows:
        output.print_md("| {} | {} | {} |".format(
            row.get("workset", ""),
            row.get("borrable", ""),
            row.get("estado", "")
        ))


# ==================================================
# COLLECT / SCAN
# ==================================================

modes = [MODE_RUN_CREATE, MODE_PROBE_LINKS]

if has_delete_workset_api():
    modes.append(MODE_PROBE_UNUSED)
    modes.append(MODE_RUN_DELETE)

modes.append(MODE_CANCEL)

switch = forms.CommandSwitchWindow.show(
    modes,
    message="Selecciona modo"
)

if not switch or switch == MODE_CANCEL:
    script.exit()

if not doc.IsWorkshared:
    forms.alert(
        "El modelo host no es workshared. No se pueden crear ni borrar worksets.",
        exitscript=True
    )

host_by_norm, host_by_id = get_host_user_workset_maps(doc)
host_norm_names = set(host_by_norm.keys())

default_host_workset = get_default_user_workset(doc)
active_workset_id = get_active_workset_id(doc)

created_count = 0
updated_count = 0
modified_count = 0
deleted_count = 0
failure_count = 0

failures = []
probe_link_rows = []
probe_unused_rows = []

skipped_links = []
scanned_links = 0

candidate_link_worksets = {}

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
        failure_count += 1
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
            failure_count += 1
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

        if switch == MODE_PROBE_LINKS:
            if ws_norm in host_norm_names:
                state = "Ya existe"
            else:
                state = "Falta crear"

            probe_link_rows.append({
                "link": link_name,
                "workset": ws_name,
                "estado": state
            })

        if ws_norm not in candidate_link_worksets:
            candidate_link_worksets[ws_norm] = {
                "name": ws_name,
                "link": link_name
            }


# ==================================================
# PROCESS
# ==================================================

if switch == MODE_RUN_CREATE:
    names_to_create = []

    sorted_norms = sorted(candidate_link_worksets.keys())
    for ws_norm in sorted_norms:
        if ws_norm not in host_norm_names:
            names_to_create.append(candidate_link_worksets[ws_norm])

    if names_to_create:
        tx = DB.Transaction(doc, "MENVIC - Create Missing Worksets")
        tx.Start()

        try:
            for item in names_to_create:
                ws_name = item.get("name", "")
                ws_norm = normalize_name(ws_name)

                try:
                    if not ws_name or not ws_norm:
                        continue

                    if ws_norm in host_norm_names:
                        continue

                    if not can_create_workset_name(doc, ws_name):
                        host_norm_names.add(ws_norm)
                        continue

                    DB.Workset.Create(doc, ws_name)
                    host_norm_names.add(ws_norm)
                    created_count += 1

                except Exception as ex:
                    failure_count += 1
                    failures.append({
                        "elemento": ws_name,
                        "accion": "Create Workset",
                        "motivo": first_line(ex)
                    })

            tx.Commit()

        except Exception as ex:
            try:
                tx.RollBack()
            except Exception:
                pass

            failure_count += 1
            failures.append({
                "elemento": "Transaction",
                "accion": "Commit Create",
                "motivo": first_line(ex)
            })

elif switch == MODE_PROBE_UNUSED or switch == MODE_RUN_DELETE:
    if not has_delete_workset_api():
        forms.alert(
            "La API de borrado de worksets no está disponible en esta versión de Revit.",
            exitscript=True
        )

    if default_host_workset is None:
        forms.alert(
            "No se encontró un workset host por defecto para construir la configuración de borrado.",
            exitscript=True
        )

    delete_settings = build_delete_settings(default_host_workset)
    if delete_settings is None:
        forms.alert(
            "No se pudo construir DeleteWorksetSettings.",
            exitscript=True
        )

    delete_candidates = []

    for ws in get_host_user_worksets(doc):
        try:
            status_code, status_text = get_unused_workset_status(
                doc,
                ws,
                candidate_link_worksets,
                active_workset_id,
                delete_settings
            )
        except Exception as ex:
            failure_count += 1
            failures.append({
                "elemento": safe_name(ws.Name),
                "accion": "Evaluate Host Workset",
                "motivo": first_line(ex)
            })
            continue

        ws_name = safe_name(ws.Name)

        if switch == MODE_PROBE_UNUSED:
            if status_code != "SKIP":
                probe_unused_rows.append({
                    "workset": ws_name,
                    "borrable": "Sí" if status_code == "YES" else "No",
                    "estado": status_text
                })

        if switch == MODE_RUN_DELETE:
            if status_code == "YES":
                delete_candidates.append(ws)

    if switch == MODE_RUN_DELETE and delete_candidates:
        tx = DB.Transaction(doc, "MENVIC - Delete Unused Host Worksets")
        tx.Start()

        try:
            for ws in delete_candidates:
                ws_name = safe_name(ws.Name)

                try:
                    status_code, status_text = get_unused_workset_status(
                        doc,
                        ws,
                        candidate_link_worksets,
                        active_workset_id,
                        delete_settings
                    )

                    if status_code != "YES":
                        continue

                    DB.WorksetTable.DeleteWorkset(doc, ws.Id, delete_settings)
                    deleted_count += 1

                except Exception as ex:
                    failure_count += 1
                    failures.append({
                        "elemento": ws_name,
                        "accion": "Delete Workset",
                        "motivo": first_line(ex)
                    })

            tx.Commit()

        except Exception as ex:
            try:
                tx.RollBack()
            except Exception:
                pass

            failure_count += 1
            failures.append({
                "elemento": "Transaction",
                "accion": "Commit Delete",
                "motivo": first_line(ex)
            })


# ==================================================
# CLEANUP / PURGE
# ==================================================

# No extra cleanup needed.


# ==================================================
# REPORT
# ==================================================

mode_label = "RUN"
if switch == MODE_PROBE_LINKS:
    mode_label = "PROBE LINKS"
elif switch == MODE_PROBE_UNUSED:
    mode_label = "PROBE UNUSED"
elif switch == MODE_RUN_DELETE:
    mode_label = "RUN DELETE"

output.print_md("# MENVIC | {} — {}".format(TOOL_NAME, mode_label))
output.print_md("")
output.print_md("## Resumen")
output.print_md("")
output.print_md("* Creados: {}".format(created_count))
output.print_md("* Actualizados: {}".format(updated_count))
output.print_md("* Modificados/Reasignados: {}".format(modified_count))
output.print_md("* Eliminados OK: {}".format(deleted_count))
output.print_md("* Fallos: {}".format(failure_count))
output.print_md("")
output.print_md("**Links cargados escaneados:** {}".format(scanned_links))
output.print_md("**Links omitidos (sin documento cargado):** {}".format(len(skipped_links)))
output.print_md("**Worksets únicos detectados en links:** {}".format(len(candidate_link_worksets)))
output.print_md("")

if switch == MODE_RUN_CREATE and created_count == 0 and failure_count == 0:
    output.print_md("✅ Todo ya estaba estandarizado.")

if switch == MODE_RUN_DELETE and deleted_count == 0 and failure_count == 0:
    output.print_md("✅ No se encontraron worksets host vacíos y borrables.")

if skipped_links:
    output.print_md("### Links omitidos")
    output.print_md("")
    for name in skipped_links:
        output.print_md("* {}".format(name))
    output.print_md("")

if switch == MODE_PROBE_LINKS:
    print_probe_links_table(probe_link_rows)

if switch == MODE_PROBE_UNUSED:
    print_probe_unused_table(probe_unused_rows)

print_fail_table(failures)
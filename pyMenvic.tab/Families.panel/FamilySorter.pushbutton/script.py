# -*- coding: utf-8 -*-
from pyrevit import forms, script
from Autodesk.Revit.DB import BasicFileInfo, OpenOptions, ModelPathUtils
import os
import re
import shutil
import time
import codecs
from collections import defaultdict

app = __revit__.Application

# ------------------------
# CONFIG
# ------------------------
DRY_RUN = False           # True = no mueve/renombra, solo reporta
SOBREESCRIBIR = False     # True = pisa si existe destino (ojo)
PAUSA_SEG = 0.02

EXT_RELACIONADAS = ['.txt', '.xml', '.jpg', '.png', '.pdf', '.cvr', '.otd']
REPORTE_CSV = "Recategorizar_y_Renombrar_Report.csv"
DUPLICADOS_CSV = "Recategorizar_y_Renombrar_Duplicados.csv"
REPORTE_TXT = "Recategorizar_y_Renombrar_Summary.txt"

INVALID_WIN = r'[<>:"/\\|?*]'

# ------------------------
# Seleccionar carpeta base
# ------------------------
carpeta_base = forms.pick_folder(title='Selecciona la carpeta BASE (biblioteca de familias)')
if not carpeta_base:
    script.exit()

# ------------------------
# Modo de organización
# ------------------------
opciones_modo = {
    "Organize ALL files": "ALL",
    "Organize ONLY files not in category": "ONLY_NOT_IN_CATEGORY",
    "Cancel": "CANCEL"
}

modo_sel = forms.CommandSwitchWindow.show(
    opciones_modo.keys(),
    message="Choose how to organize the library:",
    title="pyRevit"
)

if not modo_sel or opciones_modo.get(modo_sel) == "CANCEL":
    script.exit()

MODO_ORGANIZAR = opciones_modo[modo_sel]

# Carpetas “especiales” (no son categorías)
SPECIAL_FOLDERS = ["Projects", "ProjectTemplates", "FamilyTemplates"]

# ------------------------
# Helpers
# ------------------------
def limpiar_win(s):
    return re.sub(INVALID_WIN, '_', s)

def carpeta_categoria(cat):
    """Ej: 'Generic Models' -> 'GenericModels'"""
    if not cat:
        cat = "Uncategorized"
    cat = re.sub(r'\s+', '', cat)
    cat = limpiar_win(cat)
    return cat or "Uncategorized"

def version_real_year(ruta_rfa):
    """Devuelve (year_int o None, 'Vxx')."""
    try:
        bfi = BasicFileInfo.Extract(ruta_rfa)
        fmt = str(bfi.Format) if bfi else ""
        m = re.search(r'(19|20)\d{2}', fmt)
        if m:
            year = int(m.group(0))
            return year, "V" + str(year)[-2:]
    except Exception:
        pass
    return None, "V00"

def leer_categoria(ruta_rfa):
    """Devuelve (categoria, error). Abre el RFA."""
    doc = None
    try:
        opts = OpenOptions()
        mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(ruta_rfa)
        doc = app.OpenDocumentFile(mp, opts)

        fam = getattr(doc, "OwnerFamily", None)
        if fam and fam.FamilyCategory:
            return fam.FamilyCategory.Name, ""
        return "Uncategorized", "Sin FamilyCategory"
    except Exception as e:
        return "Uncategorized", str(e)
    finally:
        try:
            if doc:
                doc.Close(False)
        except:
            pass

def asegurar_carpeta(path):
    if os.path.exists(path):
        return True, ""
    if DRY_RUN:
        return True, "DRY_RUN"
    try:
        os.makedirs(path)
        return True, ""
    except Exception as e:
        return False, str(e)

def mover(origen, destino):
    if origen == destino:
        return "NO_CHANGE", ""
    if os.path.exists(destino) and not SOBREESCRIBIR:
        return "SKIP_EXISTS", "Destino ya existe"
    if DRY_RUN:
        return "DRY_RUN", ""
    try:
        if os.path.exists(destino) and SOBREESCRIBIR:
            os.remove(destino)
        shutil.move(origen, destino)
        return "MOVED", ""
    except Exception as e:
        return "ERROR", str(e)

def renombrar(origen, destino):
    if origen == destino:
        return "NO_CHANGE", ""
    if os.path.exists(destino) and not SOBREESCRIBIR:
        return "SKIP_EXISTS", "Destino ya existe"
    if DRY_RUN:
        return "DRY_RUN", ""
    try:
        if os.path.exists(destino) and SOBREESCRIBIR:
            os.remove(destino)
        os.rename(origen, destino)
        return "RENAMED", ""
    except Exception as e:
        return "ERROR", str(e)

def normalizar_base_con_version(base, vxx):
    """
    - Elimina años sueltos tipo .2019, _2019, -2025, (espacio)2020, (2025)
    - Elimina tokens V## existentes en cualquier parte + basura V__ V??
    - Compacta separadores
    - Deja SIEMPRE un solo sufijo final _VXX (versión real)
    """
    s = base

    # 0) eliminar años 19xx/20xx "sueltos" con separador delante (incluye '(' )
    s = re.sub(r'(?<!\d)([._\-\s\(])((?:19|20)\d{2})(?!\d)', r'\1', s)
    s = re.sub(r'\(\s*\)', '', s)  # paréntesis vacíos tras borrar (2025)

    # 1) quitar basura tipo V__ V?? V___
    s = re.sub(r'(?i)V[_?]{1,}', '', s)

    # 2) quitar tokens V\d{2} (en cualquier parte razonable)
    s = re.sub(r'(?i)(?:^|[_\-\s\.])V\d{2}(?=$|[_\-\s\.])', '', s)
    s = re.sub(r'(?i)_?V\d{2}_?', '_', s)

    # 3) compactar separadores
    s = s.replace(' ', '_')
    s = re.sub(r'_{2,}', '_', s).strip('_')
    s = re.sub(r'[-]{2,}', '-', s).strip('-')
    s = re.sub(r'[\.]{2,}', '.', s).strip('.')

    # 4) limpiar caracteres inválidos
    s = limpiar_win(s)
    if not s:
        s = "Family"

    # 5) agregar sufijo final exacto (reemplaza si ya había)
    if re.search(r'(?i)_V\d{2}$', s):
        s = re.sub(r'(?i)_V\d{2}$', "_" + vxx, s)
    else:
        s = s + "_" + vxx

    return s

def base_para_duplicados(nombre_base):
    b = re.sub(r'(?i)_V\d{2}$', '', nombre_base)
    b = re.sub(r'_{2,}', '_', b).strip('_')
    return b.lower()

def accion_final(estado_move, estado_rename, open_status):
    if open_status.startswith("SKIPPED"):
        return "SKIPPED"
    moved = estado_move == "MOVED"
    renamed = estado_rename == "RENAMED"
    if moved and renamed:
        return "MOVED+RENAMED"
    if moved:
        return "MOVED"
    if renamed:
        return "RENAMED"
    if estado_move in ("NO_CHANGE", "DRY_RUN", "SKIP_EXISTS") and estado_rename in ("NO_CHANGE", "DRY_RUN", "SKIP_EXISTS"):
        return "UNCHANGED"
    return "PARTIAL"

# ------------------------
# Mover proyectos y templates (sin tocar las familias .rfa)
# ------------------------
for root, dirs, files in os.walk(carpeta_base):
    # evitar entrar en las carpetas destino
    dirs[:] = [d for d in dirs if d not in SPECIAL_FOLDERS]

    for f in files:
        ext = os.path.splitext(f)[1].lower()
        folder = None
        if ext == ".rvt":
            folder = "Projects"
        elif ext == ".rte":
            folder = "ProjectTemplates"
        elif ext == ".rft":
            folder = "FamilyTemplates"

        if not folder:
            continue

        destino_dir = os.path.join(carpeta_base, folder)
        ok_mk, err_mk = asegurar_carpeta(destino_dir)
        if not ok_mk:
            continue

        origen = os.path.join(root, f)
        destino = os.path.join(destino_dir, f)
        mover(origen, destino)


# ------------------------
# Recolectar RFAs (recursivo)
# ------------------------
rfas = []
for root, dirs, files in os.walk(carpeta_base):
    # No entrar en carpetas especiales (proyectos / templates)
    dirs[:] = [d for d in dirs if d not in SPECIAL_FOLDERS]

    # Si el usuario eligió “ONLY_NOT_IN_CATEGORY”, ignorar cualquier carpeta de 1er nivel
    # (asumimos que las carpetas de categorías están en el 1er nivel bajo carpeta_base)
    if MODO_ORGANIZAR == "ONLY_NOT_IN_CATEGORY":
        rel = os.path.relpath(root, carpeta_base)
        if rel != ".":
            top = rel.split(os.sep)[0]
            if top not in SPECIAL_FOLDERS:
                # No procesar familias dentro de carpetas de categoría
                dirs[:] = []
                continue

    for f in files:
        if f.lower().endswith('.rfa'):
            rfas.append(os.path.join(root, f))

if not rfas:
    forms.alert("No se encontraron archivos .rfa en la carpeta seleccionada.", title="pyRevit")
    script.exit()

# ------------------------
# Proceso + Reporte
# ------------------------
t0 = time.time()

revit_actual = int(app.VersionNumber)  # ej 2024
CARPETA_NEWER = "__NEWER_THAN_{}".format(revit_actual)

reporte_path = os.path.join(carpeta_base, REPORTE_CSV)
dup_path = os.path.join(carpeta_base, DUPLICADOS_CSV)
txt_path = os.path.join(carpeta_base, REPORTE_TXT)

filas = [
    "accion,open_status,version_real,version_year,categoria,carpeta_categoria,"
    "origen_path,origen_archivo,destino_path,destino_archivo,"
    "estado_move,estado_rename,error"
]

cnt = {"TOTAL": 0, "UNCHANGED": 0, "RENAMED": 0, "MOVED": 0, "MOVED+RENAMED": 0, "SKIPPED": 0, "ERROR": 0}
dup_map = {}
dup_count = 0

# BIM Manager audit counters
by_folder = defaultdict(int)
by_action = defaultdict(int)
by_open_status = defaultdict(int)
rel_files_moved = 0
rel_files_renamed = 0
folders_created = set()
skipped_files = []
error_items = []

# TXT agrupado tipo “reporte”
grouped_txt = {}  # grouped_txt[folder][action] = [ (file, vxx, open_status, note) ]

with forms.ProgressBar(title='Recategorizando + renombrando ({value}/{max})', step=1, cancellable=True) as pb:
    total = len(rfas)
    for i, ruta_rfa in enumerate(rfas, start=1):
        if pb.cancelled:
            break
        pb.update_progress(i, total)

        cnt["TOTAL"] += 1

        origen_dir = os.path.dirname(ruta_rfa)
        origen_archivo = os.path.basename(ruta_rfa)
        base_original, ext = os.path.splitext(origen_archivo)

        year, vxx = version_real_year(ruta_rfa)

        open_status = "OK"
        cat = "Uncategorized"
        err_cat = ""

        # Si es más nuevo, no abrir; mandar a carpeta especial
        if year is not None and year > revit_actual:
            open_status = "SKIPPED_NEWER_VERSION"
            err_cat = "Archivo {} > Revit {}".format(year, revit_actual)
            carpeta_cat = CARPETA_NEWER
        else:
            cat, err_cat = leer_categoria(ruta_rfa)
            if err_cat:
                open_status = "OPEN_ERROR"
            carpeta_cat = carpeta_categoria(cat)

        # BIM counters (pre)
        by_folder[carpeta_cat] += 1
        by_open_status[open_status] += 1

        destino_dir = os.path.join(carpeta_base, carpeta_cat)
        ok_mk, err_mk = asegurar_carpeta(destino_dir)
        if ok_mk:
            folders_created.add(destino_dir)

        if not ok_mk:
            acc = "ERROR"
            cnt["ERROR"] += 1
            by_action[acc] += 1

            note = "MKDIR_ERROR: {}".format(err_mk)
            error_items.append((origen_archivo, note))

            filas.append("{},{},{},{},{},{},{},{},{},{},{},{},{}".format(
                acc, "MKDIR_ERROR", vxx, year if year else "", cat, carpeta_cat,
                origen_dir.replace(",", ";"), origen_archivo.replace(",", ";"),
                "", "",
                "", "",
                note.replace(",", ";")
            ))

            grouped_txt.setdefault(carpeta_cat, {}).setdefault(acc, []).append(
                (origen_archivo, vxx, "MKDIR_ERROR", note)
            )
            continue

        # mover rfa
        ruta_en_destino = os.path.join(destino_dir, origen_archivo)
        estado_move, err_move = mover(ruta_rfa, ruta_en_destino)

        # mover relacionados
        for ext_rel in EXT_RELACIONADAS:
            rel = base_original + ext_rel
            ruta_rel = os.path.join(origen_dir, rel)
            if os.path.exists(ruta_rel):
                st_rel, _ = mover(ruta_rel, os.path.join(destino_dir, rel))
                if st_rel == "MOVED":
                    rel_files_moved += 1

        # renombrar rfa
        ruta_actual = ruta_en_destino if estado_move in ("MOVED", "NO_CHANGE", "DRY_RUN", "SKIP_EXISTS") else ruta_rfa
        base_limpia = normalizar_base_con_version(base_original, vxx)
        destino_archivo = base_limpia + ext
        ruta_final = os.path.join(destino_dir, destino_archivo)

        estado_rename, err_rename = renombrar(ruta_actual, ruta_final)

        # renombrar relacionados en destino
        for ext_rel in EXT_RELACIONADAS:
            rel_old = os.path.join(destino_dir, base_original + ext_rel)
            if os.path.exists(rel_old):
                st_rr, _ = renombrar(rel_old, os.path.join(destino_dir, base_limpia + ext_rel))
                if st_rr == "RENAMED":
                    rel_files_renamed += 1

        acc = accion_final(estado_move, estado_rename, open_status)
        by_action[acc] += 1

        if acc in cnt:
            cnt[acc] += 1
        else:
            cnt["ERROR"] += 1

        # skipped list
        if open_status.startswith("SKIPPED"):
            skipped_files.append(origen_archivo)

        # duplicados map
        key = (carpeta_cat, base_para_duplicados(base_limpia))
        dup_map.setdefault(key, []).append(destino_archivo)

        err_total = " | ".join([x for x in [err_cat, err_mk, err_move, err_rename] if x])
        if (open_status == "OPEN_ERROR") or (estado_move == "ERROR") or (estado_rename == "ERROR"):
            cnt["ERROR"] += 1
            if err_total:
                error_items.append((destino_archivo, err_total))

        # CSV row
        filas.append("{},{},{},{},{},{},{},{},{},{},{},{},{}".format(
            acc,
            open_status,
            vxx,
            year if year else "",
            (cat or "").replace(",", ";"),
            carpeta_cat.replace(",", ";"),
            origen_dir.replace(",", ";"),
            origen_archivo.replace(",", ";"),
            destino_dir.replace(",", ";"),
            destino_archivo.replace(",", ";"),
            estado_move,
            estado_rename,
            (err_total or "").replace(",", ";")
        ))

        # TXT grouped
        grouped_txt.setdefault(carpeta_cat, {}).setdefault(acc, []).append(
            (destino_archivo, vxx, open_status, err_total if err_total else "")
        )

        time.sleep(PAUSA_SEG)

# ------------------------
# Guardar CSV reportes (UTF-8)
# ------------------------
with codecs.open(reporte_path, "w", encoding="utf-8") as out:
    out.write(u"\n".join([u"{}".format(x) for x in filas]))

dup_lines = ["carpeta_categoria,base_sin_version,archivos"]
dup_count = 0
for (catfolder, basekey), files in sorted(dup_map.items()):
    if len(files) > 1:
        dup_count += 1
        dup_lines.append("{},{},{}".format(
            catfolder.replace(",", ";"),
            basekey.replace(",", ";"),
            (" | ".join(sorted(files))).replace(",", ";")
        ))

with codecs.open(dup_path, "w", encoding="utf-8") as out:
    out.write(u"\n".join([u"{}".format(x) for x in dup_lines]))

# ------------------------
# TXT Summary “tipo ejemplo” (agrupado y legible)
# ------------------------
ACTION_ORDER = ["MOVED+RENAMED", "MOVED", "RENAMED", "UNCHANGED", "SKIPPED", "PARTIAL", "ERROR"]

def folder_sort_key(k):
    if k.startswith("__NEWER_THAN_"):
        return "ZZZ_" + k
    return k

txt_lines = []
txt_lines.append("FAMILY LIBRARY AUDIT REPORT")
txt_lines.append("")
txt_lines.append("Root Folder: {}".format(carpeta_base))
txt_lines.append("Revit current: {}".format(revit_actual))
txt_lines.append("Mode: {}".format("DRY_RUN" if DRY_RUN else "EXECUTED"))
txt_lines.append("")

txt_lines.append("SUMMARY")
txt_lines.append("")
txt_lines.append("TOTAL: {}".format(cnt["TOTAL"]))
txt_lines.append("UNCHANGED: {}".format(cnt["UNCHANGED"]))
txt_lines.append("RENAMED: {}".format(cnt["RENAMED"]))
txt_lines.append("MOVED: {}".format(cnt["MOVED"]))
txt_lines.append("MOVED+RENAMED: {}".format(cnt["MOVED+RENAMED"]))
txt_lines.append("SKIPPED (newer Revit): {}".format(cnt["SKIPPED"]))
txt_lines.append("ERROR (review): {}".format(cnt["ERROR"]))
txt_lines.append("DUPLICATE GROUPS: {}".format(dup_count))
txt_lines.append("")
txt_lines.append("")

for folder in sorted(grouped_txt.keys(), key=folder_sort_key):
    txt_lines.append("Folder: {}".format(folder))
    txt_lines.append("")
    actions = grouped_txt[folder]
    for act in ACTION_ORDER:
        if act not in actions:
            continue
        txt_lines.append("Action: {}".format(act))
        txt_lines.append("")
        txt_lines.append("FileName\tVersion\tOpenStatus\tNotes")
        txt_lines.append("")
        for (fname, vxx, open_status, note) in sorted(actions[act], key=lambda x: x[0].lower()):
            note_clean = (note or "").replace("\t", " ").strip()
            txt_lines.append(u"{}\t{}\t{}\t{}".format(fname, vxx, open_status, note_clean))
        txt_lines.append("")
        txt_lines.append("")
    txt_lines.append("")

txt_lines.append("POSSIBLE DUPLICATES (groups)")
txt_lines.append("")
shown = 0
for (catfolder, basekey), files in sorted(dup_map.items()):
    if len(files) > 1:
        shown += 1
        txt_lines.append("Group {}: Folder={} Base={}".format(shown, catfolder, basekey))
        for f in sorted(files):
            txt_lines.append("  - {}".format(f))
        txt_lines.append("")
        if shown >= 50:
            txt_lines.append("... (truncated to 50 groups)")
            txt_lines.append("")
            break

txt_lines.append("SEARCH COMPLETED - {} families processed.".format(cnt["TOTAL"]))

with codecs.open(txt_path, "w", encoding="utf-8") as out:
    out.write(u"\n".join([u"{}".format(x) for x in txt_lines]))

# ------------------------
# BIM Manager popup summary
# ------------------------
t1 = time.time()
elapsed = max(0.001, (t1 - t0))
rate = float(cnt["TOTAL"]) / elapsed

top_folders = sorted(by_folder.items(), key=lambda x: x[1], reverse=True)[:8]
top_folders_txt = "\n".join(["- {}: {}".format(k, v) for k, v in top_folders]) if top_folders else "- (none)"

action_order = ["MOVED+RENAMED", "MOVED", "RENAMED", "UNCHANGED", "SKIPPED", "PARTIAL", "ERROR"]
actions_txt = "\n".join(["- {}: {}".format(a, by_action.get(a, 0)) for a in action_order])

status_txt = "\n".join(["- {}: {}".format(k, by_open_status[k]) for k in sorted(by_open_status.keys())])

err_preview = "- (none)"
if error_items:
    err_preview = "\n".join(["- {} :: {}".format(a, b) for a, b in error_items[:5]])
    if len(error_items) > 5:
        err_preview += "\n- ... ({} more)".format(len(error_items) - 5)

sk_preview = "- (none)"
if skipped_files:
    sk_preview = "\n".join(["- {}".format(x) for x in skipped_files[:10]])
    if len(skipped_files) > 10:
        sk_preview += "\n- ... ({} more)".format(len(skipped_files) - 10)

resumen_bim = "\n".join([
    "BIM MANAGER SUMMARY",
    "",
    "Root: {}".format(carpeta_base),
    "Revit: {}".format(revit_actual),
    "Mode: {}".format("DRY_RUN" if DRY_RUN else "EXECUTED"),
    "",
    "Throughput:",
    "- Processed: {}".format(cnt["TOTAL"]),
    "- Time: {:.1f} s".format(elapsed),
    "- Speed: {:.2f} families/s".format(rate),
    "",
    "Actions:",
    actions_txt,
    "",
    "Open status:",
    status_txt,
    "",
    "Category folders (top):",
    top_folders_txt,
    "",
    "Related files:",
    "- Moved: {}".format(rel_files_moved),
    "- Renamed: {}".format(rel_files_renamed),
    "",
    "Duplicates:",
    "- Groups: {}".format(dup_count),
    "- File: {}".format(dup_path),
    "",
    "Skipped (newer than {}):".format(revit_actual),
    sk_preview,
    "",
    "Errors (preview):",
    err_preview,
    "",
    "Outputs:",
    "- CSV Report: {}".format(reporte_path),
    "- CSV Duplicates: {}".format(dup_path),
    "- TXT Summary: {}".format(txt_path),
])

forms.alert(resumen_bim, title="pyRevit - Family Library Audit")
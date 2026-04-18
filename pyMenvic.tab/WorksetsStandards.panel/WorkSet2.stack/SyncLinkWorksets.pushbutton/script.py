# -*- coding: utf-8 -*-

from pyrevit import revit, DB, forms, script
import os
import clr

clr.AddReference("PresentationCore")
clr.AddReference("System")
clr.AddReference("System.Core")
clr.AddReference("System.Windows")

from System.Threading import Thread

__title__ = "SYNC LINK WORKSETS"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
==========================================================
pyMENVIC | SYNC LINK WORKSETS
Revit + pyRevit

Descripción
-----------
Detecta worksets de usuario en los modelos linkeados cargados,
los consolida por nombre y permite crear en el modelo host solo
los worksets seleccionados desde una ventana XAML.

Capacidades
-----------
- Escanea worksets de todos los Revit links cargados
- Consolida nombres repetidos entre links
- Filtra por link, texto y estado faltante
- Crea solo worksets faltantes seleccionados
- Ejecuta directo a la ventana XAML

Funciones principales
---------------------
safe_name
    Limpia nombres vacíos o nulos

normalize
    Normaliza el nombre para comparar sin duplicados

scan_link_worksets
    Construye la lista consolidada de worksets detectados

load_ui_strings
    Carga textos UI desde strings.py

Reglas importantes
------------------
- Solo crea worksets UserWorkset en el modelo host
- No crea duplicados si el nombre ya existe
- Si un link no está cargado, se omite sin detener el script
- La creación ocurre solo al confirmar desde la UI

Autor
-----
Ricardo J. Mendieta
pyMENVIC – Ayudas para MENVIC ARQ
==========================================================
"""

# ==================================================
# CONFIG
# ==================================================

doc = revit.doc
output = script.get_output()
SCRIPT_DIR = os.path.dirname(__file__)
XAML_PATH = os.path.join(SCRIPT_DIR, "sync_link_worksets.xaml")
STRINGS_PATH = os.path.join(SCRIPT_DIR, "strings.py")

DEFAULT_STRINGS = {
    "Sync Link Worksets - pyMENVIC": "Sync Link Worksets - pyMENVIC",
    "SYNC LINK WORKSETS": "SYNC LINK WORKSETS",
    "Detect worksets in loaded links, filter the list, and create only the selected ones in the host model.": "Detect worksets in loaded links, filter the list, and create only the selected ones in the host model.",
    "WORKSETS: {0}/{1}": "WORKSETS: {0}/{1}",
    "LINKS: {0}": "LINKS: {0}",
    "LINK FILTER": "LINK FILTER",
    "SEARCH WORKSET": "SEARCH WORKSET",
    "ONLY MISSING": "ONLY MISSING",
    "REFRESH": "REFRESH",
    "USE": "USE",
    "WORKSET": "WORKSET",
    "STATUS": "STATUS",
    "SOURCE LINKS": "SOURCE LINKS",
    "LINKS HEADER": "LINKS",
    "Only missing checked rows will be created.": "Only missing checked rows will be created.",
    "CHECK VISIBLE": "CHECK VISIBLE",
    "UNCHECK VISIBLE": "UNCHECK VISIBLE",
    "CHECK MISSING": "CHECK MISSING",
    "CREATE SELECTED": "CREATE SELECTED",
    "CLOSE": "CLOSE",
    "ALL LINKS": "ALL LINKS",
    "Missing": "Missing",
    "Exists": "Exists",
    "Model is not workshared.": "Model is not workshared.",
    "Could not read loaded links.": "Could not read loaded links.",
    "No missing worksets were found.": "No missing worksets were found.",
    "No rows are selected.": "No rows are selected.",
    "No checked rows are missing. Nothing was created.": "No checked rows are missing. Nothing was created.",
    "Create Worksets": "Create Worksets",
    "{0} worksets created.": "{0} worksets created.",
    "Scanned links: {0}": "Scanned links: {0}",
    "Skipped links: {0}": "Skipped links: {0}",
    "XAML file not found: {0}": "XAML file not found: {0}",
    "LINKED: {0}": "LINKED: {0}",
    "CHECKED: {0}": "CHECKED: {0}",
}

# ==================================================
# HELPERS
# ==================================================


def safe_name(name):
    if not name:
        return ""
    return str(name).strip()


def normalize(name):
    return safe_name(name).lower()

def get_ui_language():
    try:
        lang = Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName

        if lang == "es":
            return "es"
    except:
        pass

    return "en"

def load_ui_strings():
    ui_strings = dict(DEFAULT_STRINGS)

    if not os.path.exists(STRINGS_PATH):
        return ui_strings

    namespace = {}

    try:
        execfile(STRINGS_PATH, namespace)
        detected_language = get_ui_language()
        all_strings = namespace.get("STRINGS", {})

        if detected_language in all_strings:
            ui_strings.update(all_strings.get(detected_language, {}))
        elif "en" in all_strings:
            ui_strings.update(all_strings.get("en", {}))
    except:
        pass

    return ui_strings


def tr(ui_strings, key):
    return ui_strings.get(key, key)


def get_host_worksets():
    result = {}
    collector = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)

    for ws in collector:
        result[normalize(ws.Name)] = ws

    return result


def get_link_instances():
    return DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)


def get_link_doc(link):
    try:
        return link.GetLinkDocument()
    except:
        return None


def get_link_worksets(link_doc):
    result = []
    collector = DB.FilteredWorksetCollector(link_doc).OfKind(DB.WorksetKind.UserWorkset)

    for ws in collector:
        result.append(ws)

    return result


def scan_link_worksets(host_worksets, ui_strings):
    candidate_link_worksets = {}
    skipped_links = []
    scanned_links = 0
    link_names = []

    for link in get_link_instances():
        link_name = safe_name(link.Name)
        link_doc = get_link_doc(link)

        if not link_doc:
            skipped_links.append(link_name)
            continue

        scanned_links += 1
        link_names.append(link_name)

        for ws in get_link_worksets(link_doc):
            ws_name = safe_name(ws.Name)
            ws_norm = normalize(ws_name)

            if not ws_name:
                continue

            if ws_norm not in candidate_link_worksets:
                candidate_link_worksets[ws_norm] = {
                    "name": ws_name,
                    "links": []
                }

            if link_name not in candidate_link_worksets[ws_norm]["links"]:
                candidate_link_worksets[ws_norm]["links"].append(link_name)

    rows = []

    for norm in sorted(candidate_link_worksets.keys()):
        item = candidate_link_worksets[norm]
        is_missing = norm not in host_worksets

        if is_missing:
            status = tr(ui_strings, "Missing")
        else:
            status = tr(ui_strings, "Exists")

        rows.append(
            WorksetRow(
                item["name"],
                norm,
                status,
                item["links"],
                is_missing
            )
        )

    link_names = sorted(list(set(link_names)))
    return rows, scanned_links, skipped_links, link_names


# ==================================================
# DATA MODEL
# ==================================================

class WorksetRow(object):
    def __init__(self, workset_name, normalized_name, status, source_links, is_missing):
        self.IsChecked = is_missing
        self.WorksetName = workset_name
        self.NormalizedName = normalized_name
        self.Status = status
        self.SourceLinks = source_links
        self.SourceLinksText = ", ".join(source_links)
        self.SourceCountText = str(len(source_links))
        self.IsMissing = is_missing


# ==================================================
# UI
# ==================================================

class SyncLinkWorksetsWindow(forms.WPFWindow):
    def __init__(self, xaml_path, rows, scanned_links, skipped_links, link_names, ui_strings):
        self.ui_strings = ui_strings
        self.all_rows = rows
        self.filtered_rows = []
        self.scanned_links = scanned_links
        self.skipped_links = skipped_links
        self.link_names = link_names

        forms.WPFWindow.__init__(self, xaml_path)

        self._load_logo()
        self._apply_translations()
        self._setup_filters()
        self._apply_current_filter()
        self._refresh_summary()
        self.ShowDialog()

    def _apply_translations(self):
        self.Title = tr(self.ui_strings, "Sync Link Worksets - pyMENVIC")
        self.title_text.Text = tr(self.ui_strings, "SYNC LINK WORKSETS")
        self.subtitle_text.Text = tr(
            self.ui_strings,
            "Detect worksets in loaded links, filter the list, and create only the selected ones in the host model."
        )
        self.link_filter_label.Text = tr(self.ui_strings, "LINK FILTER")
        self.search_label.Text = tr(self.ui_strings, "SEARCH WORKSET")
        self.only_missing_check.Content = tr(self.ui_strings, "ONLY MISSING")
        self.refresh_button.Content = tr(self.ui_strings, "REFRESH")
        self.col_use_text.Text = tr(self.ui_strings, "USE")
        self.col_workset_text.Text = tr(self.ui_strings, "WORKSET")
        self.col_status_text.Text = tr(self.ui_strings, "STATUS")
        self.col_source_links_text.Text = tr(self.ui_strings, "SOURCE LINKS")
        self.col_link_count_text.Text = tr(self.ui_strings, "LINKS")
        self.footer_info_text.Text = tr(
            self.ui_strings,
            "Only missing checked rows will be created."
        )
        self.check_visible_button.Content = tr(self.ui_strings, "CHECK VISIBLE")
        self.uncheck_visible_button.Content = tr(self.ui_strings, "UNCHECK VISIBLE")
        self.check_missing_button.Content = tr(self.ui_strings, "CHECK MISSING")
        self.create_button.Content = tr(self.ui_strings, "CREATE SELECTED")
        self.close_button.Content = tr(self.ui_strings, "CLOSE")

    def _load_logo(self):
        logo_path = os.path.join(SCRIPT_DIR, "logo.png")

        try:
            if os.path.exists(logo_path):
                from System import Uri, UriKind
                from System.Windows.Media.Imaging import BitmapImage

                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()

                self.logo_image.Source = bitmap
        except:
            pass

    def _setup_filters(self):
        items = [tr(self.ui_strings, "ALL LINKS")]
        items.extend(self.link_names)

        self.link_filter_combo.ItemsSource = items
        self.link_filter_combo.SelectedIndex = 0

        self.search_box.TextChanged += self.on_filter_changed
        self.link_filter_combo.SelectionChanged += self.on_filter_changed
        self.only_missing_check.Checked += self.on_filter_changed
        self.only_missing_check.Unchecked += self.on_filter_changed
        self.refresh_button.Click += self.on_refresh_clicked
        self.check_visible_button.Click += self.on_check_visible_clicked
        self.uncheck_visible_button.Click += self.on_uncheck_visible_clicked
        self.check_missing_button.Click += self.on_check_missing_clicked
        self.create_button.Click += self.on_create_clicked
        self.close_button.Click += self.on_close_clicked

        self.worksets_grid.CurrentCellChanged += self.on_grid_selection_changed
    def on_grid_selection_changed(self, sender, args):
        try:
            self._refresh_summary()
        except:
            pass

    def _apply_current_filter(self):
        selected_link = safe_name(self.link_filter_combo.SelectedItem)
        all_links_label = tr(self.ui_strings, "ALL LINKS")
        search_text = normalize(self.search_box.Text)
        only_missing = False

        try:
            only_missing = bool(self.only_missing_check.IsChecked)
        except:
            only_missing = False

        filtered = []

        for item in self.all_rows:
            if selected_link and selected_link != all_links_label:
                if selected_link not in item.SourceLinks:
                    continue

            if search_text:
                if search_text not in normalize(item.WorksetName):
                    continue

            if only_missing and not item.IsMissing:
                continue

            filtered.append(item)

        self.filtered_rows = filtered
        self.worksets_grid.ItemsSource = None
        self.worksets_grid.ItemsSource = self.filtered_rows
        self.worksets_grid.Items.Refresh()

    def _iter_visible_rows(self):
        for row in self.filtered_rows:
            yield row

    def _refresh_summary(self):
        visible_count = len(self.filtered_rows)
        total_count = len(self.all_rows)

        checked_count = 0
        for row in self.all_rows:
            if row.IsChecked:
                checked_count += 1

        self.summary_text.Text = "{0}/{1}".format(visible_count, total_count)
        self.linked_text.Text = str(self.scanned_links)
        self.checked_text.Text = str(checked_count)

    def on_filter_changed(self, sender, args):
        self._apply_current_filter()
        self._refresh_summary()

    def on_refresh_clicked(self, sender, args):
        self._apply_current_filter()
        self._refresh_summary()

    def on_check_visible_clicked(self, sender, args):
        for row in self._iter_visible_rows():
            row.IsChecked = True
        self.worksets_grid.Items.Refresh()
        self._refresh_summary()

    def on_uncheck_visible_clicked(self, sender, args):
        for row in self._iter_visible_rows():
            row.IsChecked = False
        self.worksets_grid.Items.Refresh()
        self._refresh_summary()

    def on_check_missing_clicked(self, sender, args):
        for row in self._iter_visible_rows():
            row.IsChecked = row.IsMissing
        self.worksets_grid.Items.Refresh()
        self._refresh_summary()

    def on_create_clicked(self, sender, args):
        self.DialogResult = True
        self.Close()

    def on_close_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()


# ==================================================
# SCAN
# ==================================================

UI_STRINGS = load_ui_strings()

if not doc.IsWorkshared:
    forms.alert(tr(UI_STRINGS, "Model is not workshared."), exitscript=True)

host_worksets = get_host_worksets()
rows, scanned_links, skipped_links, link_names = scan_link_worksets(host_worksets, UI_STRINGS)

# ==================================================
# PROCESS
# ==================================================

missing_count = 0
for row in rows:
    if row.IsMissing:
        missing_count += 1

if missing_count == 0:
    if scanned_links == 0:
        forms.alert(tr(UI_STRINGS, "Could not read loaded links."), exitscript=True)

    forms.alert(tr(UI_STRINGS, "No missing worksets were found."), exitscript=True)

if not os.path.exists(XAML_PATH):
    forms.alert(tr(UI_STRINGS, "XAML file not found: {0}").format(XAML_PATH), exitscript=True)

window = SyncLinkWorksetsWindow(
    XAML_PATH,
    rows,
    scanned_links,
    skipped_links,
    link_names,
    UI_STRINGS
)

if not window.DialogResult:
    script.exit()

selected_rows = []
for row in rows:
    if row.IsChecked:
        selected_rows.append(row)

if not selected_rows:
    forms.alert(tr(UI_STRINGS, "No rows are selected."), exitscript=True)

current_host_worksets = get_host_worksets()
names_to_create = []

for row in selected_rows:
    if row.NormalizedName not in current_host_worksets:
        names_to_create.append(row)

if not names_to_create:
    forms.alert(tr(UI_STRINGS, "No checked rows are missing. Nothing was created."), exitscript=True)

tx = DB.Transaction(doc, tr(UI_STRINGS, "Create Worksets"))
tx.Start()

created = 0

for row in names_to_create:
    try:
        if DB.WorksetTable.IsWorksetNameUnique(doc, row.WorksetName):
            DB.Workset.Create(doc, row.WorksetName)
            created += 1
    except:
        pass

tx.Commit()

# ==================================================
# REPORT
# ==================================================

forms.alert(tr(UI_STRINGS, "{0} worksets created.").format(created))

output.print_md("")
output.print_md(tr(UI_STRINGS, "Scanned links: {0}").format(scanned_links))
output.print_md(tr(UI_STRINGS, "Skipped links: {0}").format(len(skipped_links)))
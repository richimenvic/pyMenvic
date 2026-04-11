# -*- coding: utf-8 -*-
__title__ = "Revision Sheet Set"

# pylint: disable=E0401,C0103,C0111

import os
import xml.etree.ElementTree as ET

import Autodesk.Revit.DB as DB
from pyrevit import revit, forms, script
from System.Windows import Thickness, Visibility, GridLength, GridUnitType, VerticalAlignment, HorizontalAlignment, TextAlignment, TextTrimming
from System.Windows.Controls import CheckBox, Grid, ColumnDefinition, TextBlock
from System.Windows.Controls import Border
from System.Windows.Media import ColorConverter


doc = revit.doc
output = script.get_output()
logger = script.get_logger()
SCRIPT_DIR = os.path.dirname(__file__)
CONFIG_PATH = os.path.join(SCRIPT_DIR, 'config.xml')
CONFIG_ES_PATH = os.path.join(SCRIPT_DIR, 'config.es.xml')
XAML_PATH = os.path.join(SCRIPT_DIR, 'RevisionSelector.xaml')

DEFAULT_CONFIG = {
    "settings": {
        "language": "auto",
    },
    "ui": {
        "button_name": "Create Sheet Set",
        "selector_title": "Select Revisions",
        "search_label": "Search revisions",
        "search_placeholder": "Type to filter revisions...",
        "mode_label": "Sheets to include",
        "mode_help": "All selected revisions = only sheets that contain every selected revision. Any selected revision = sheets that contain at least one selected revision.",
        "revisions_label": "Available revisions",
        "selection_status": "Selected: {} | Visible: {} | Total: {}",
        "pick_option": "Pick an option:",
        "empty_sheets_title": "**Empty sheets detected:**",
        "report_title": "### pyMenvic | Revision Sheet Set - RUN",
        "report_sheetset": "# Sheet Set: {}",
        "report_revision": "## {}",
        "report_summary": "Status: {} | Mode: {} | Sheets: {} | Empty: {}",
        "report_note": "**Note:** {}",
        "report_cancelled": "**Note:** Operation cancelled by user.",
        "report_revisions_included": "**Revisions included:**",
        "report_graphic_check": "#### Graphic Check",
    },
    "options": {
        "match_any": "Any selected revision",
        "match_all": "All selected revisions",
    },
    "messages": {
        "not_created": "Not created",
        "created_name_unknown": "Created (name unknown)",
        "no_revision": "None",
        "revision_fallback": "Revision",
        "revision_id_prefix": "Revision Id {}",
        "multiple_revisions": "{} revisions",
        "mode_none": "None",
        "no_matches": "No sheets match the selected revision(s).",
        "create_failed": "Failed to create sheet set: {}",
        "xaml_missing": "Missing XAML file:\n{}",
        "select_revisions_required": "Select at least one revision.",
        "check_yes": "Yes",
        "check_no": "No",
    },
    "checks": {
        "revisions_selected": "Revisions selected",
        "match_option_picked": "Match option picked",
        "has_matches": "Has matches (pre-check)",
        "sheetset_created": "Sheet set created",
        "sheets_in_set": "Sheets in set > 0",
        "no_empty_sheets": "No empty sheets",
        "status_ok": "[OK]",
        "status_fail": "[FAIL]",
    },
    "buttons": {
        "select_visible": "Select Visible",
        "clear_visible": "Clear Visible",
        "invert_visible": "Invert Visible",
        "cancel": "Cancel",
        "create": "Create",
    },
    "theme": {
        "dark_primary": "#2C3440",
        "dark_secondary": "#3B4352",
        "dark_window_bg": "#252B34",
        "dark_surface": "#313947",
        "dark_list_surface": "#3A4353",
        "dark_input_bg": "#232933",
        "dark_text": "#F4F7FA",
        "dark_text_muted": "#BFC7D1",
        "dark_border": "#4B5565",
        "dark_primary_button": "#D89C34",
        "dark_primary_button_border": "#B57E22",
        "light_primary": "#F4F6F8",
        "light_secondary": "#E7EBF0",
        "light_window_bg": "#F7F7F8",
        "light_surface": "#FFFFFF",
        "light_list_surface": "#FFFFFF",
        "light_input_bg": "#FFFFFF",
        "light_text": "#1F2933",
        "light_text_muted": "#5B6672",
        "light_border": "#C7CDD4",
        "light_primary_button": "#D89C34",
        "light_primary_button_border": "#B57E22",
    },
}

# ----------------------------
# INPUTS
# ----------------------------
revisions = None

# ----------------------------
# PROCESS VARS
# ----------------------------
selected_switch = None
match_any = False

rev_sheetset = None
rev_sheetset_count = 0
empty_sheets = []

match_count = 0
error_reason = None
sheetset_name = DEFAULT_CONFIG["messages"]["not_created"]
run_status = "CANCELLED"
CONFIG = None


def _log_debug(message, ex=None):
    try:
        if ex is None:
            logger.debug(message)
        else:
            logger.debug("{} | {}".format(message, ex))
    except Exception:
        pass


def _new_result():
    return {
        "selected_switch": None,
        "match_any": False,
        "rev_sheetset": None,
        "rev_sheetset_count": 0,
        "empty_sheets": [],
        "match_count": 0,
        "error_reason": None,
        "sheetset_name": _cfg("messages.not_created"),
        "run_status": "CANCELLED",
    }


def _deep_copy_dict(data):
    copied = {}
    for key, value in data.items():
        if isinstance(value, dict):
            copied[key] = _deep_copy_dict(value)
        else:
            copied[key] = value
    return copied


def _merge_dicts(target, updates):
    for key, value in updates.items():
        if isinstance(value, dict) and isinstance(target.get(key), dict):
            _merge_dicts(target[key], value)
        else:
            target[key] = value


def _xml_node_text(parent, tag_name):
    if parent is None:
        return None

    child = parent.find(tag_name)
    if child is None or child.text is None:
        return None

    value = child.text.strip()
    return value if value else None


def _load_config():
    config = _deep_copy_dict(DEFAULT_CONFIG)
    _load_config_file(CONFIG_PATH, config)

    language_code = _resolve_language_code(config.get("settings", {}).get("language", "auto"))
    if language_code == "es":
        _load_config_file(CONFIG_ES_PATH, config)

    return config


def _load_config_file(config_path, config):
    if not os.path.exists(config_path):
        return

    try:
        root = ET.parse(config_path).getroot()
    except Exception as ex:
        _log_debug("Failed to parse {}".format(os.path.basename(config_path)), ex)
        return

    loaded = {}
    for section_name in config.keys():
        loaded[section_name] = {}
        section_node = root.find(section_name)
        if section_node is None:
            continue

        for key in config[section_name].keys():
            value = _xml_node_text(section_node, key)
            if value is not None:
                loaded[section_name][key] = value

    _merge_dicts(config, loaded)


def _resolve_language_code(language_setting):
    language_text = _safe_str(language_setting).strip().lower()
    if language_text in ("en", "es"):
        return language_text

    try:
        app_language = _safe_str(__revit__.Application.Language).lower()
        if "spanish" in app_language or "espan" in app_language:
            return "es"
    except Exception as ex:
        _log_debug("Failed to detect Revit language", ex)

    return "en"


def _cfg(path):
    current = CONFIG if CONFIG else DEFAULT_CONFIG
    for part in path.split('.'):
        current = current.get(part)
        if current is None:
            return ""
    return current


def _is_revit_dark_theme():
    try:
        import Autodesk.Revit.UI as UI
        theme_manager = getattr(UI, "UIThemeManager", None)
        if theme_manager is not None:
            current_theme = getattr(theme_manager, "CurrentTheme", None)
            if current_theme is not None:
                return "dark" in _safe_str(current_theme).lower()
    except Exception as ex:
        _log_debug("UIThemeManager theme detection unavailable", ex)

    try:
        app_theme = getattr(__revit__.Application, "Theme", None)
        if app_theme is not None:
            return "dark" in _safe_str(app_theme).lower()
    except Exception as ex:
        _log_debug("Application theme detection unavailable", ex)

    return True


def _get_theme_palette(is_dark):
    if is_dark:
        return {
            "WindowBgBrush": _cfg("theme.dark_window_bg"),
            "PanelBgBrush": _cfg("theme.dark_primary"),
            "PanelAltBrush": _cfg("theme.dark_secondary"),
            "SurfaceBgBrush": _cfg("theme.dark_surface"),
            "ListSurfaceBgBrush": _cfg("theme.dark_list_surface"),
            "InputBgBrush": _cfg("theme.dark_input_bg"),
            "BorderBrush": _cfg("theme.dark_border"),
            "TextPrimaryBrush": _cfg("theme.dark_text"),
            "TextSecondaryBrush": _cfg("theme.dark_text_muted"),
            "ButtonBgBrush": _cfg("theme.dark_secondary"),
            "ButtonHoverBrush": _cfg("theme.dark_primary"),
            "PrimaryButtonBgBrush": _cfg("theme.dark_primary_button"),
            "PrimaryButtonBorderBrush": _cfg("theme.dark_primary_button_border"),
        }

    return {
        "WindowBgBrush": _cfg("theme.light_window_bg"),
        "PanelBgBrush": _cfg("theme.light_primary"),
        "PanelAltBrush": _cfg("theme.light_secondary"),
        "SurfaceBgBrush": _cfg("theme.light_surface"),
        "ListSurfaceBgBrush": _cfg("theme.light_list_surface"),
        "InputBgBrush": _cfg("theme.light_input_bg"),
        "BorderBrush": _cfg("theme.light_border"),
        "TextPrimaryBrush": _cfg("theme.light_text"),
        "TextSecondaryBrush": _cfg("theme.light_text_muted"),
        "ButtonBgBrush": _cfg("theme.light_secondary"),
        "ButtonHoverBrush": _cfg("theme.light_primary"),
        "PrimaryButtonBgBrush": _cfg("theme.light_primary_button"),
        "PrimaryButtonBorderBrush": _cfg("theme.light_primary_button_border"),
    }


class RevisionOption(object):
    def __init__(self, revision, seq, desc, rev_date, label):
        self.revision = revision
        self.seq = seq
        self.desc = desc
        self.rev_date = rev_date
        self.label = label
        self.id_value = _safe_revision_id_value(revision)
        self.sort_seq = _sort_key_for_seq(seq)


class RevisionSelectorWindow(forms.WPFWindow):
    def __init__(self, xaml_path, revisions):
        forms.WPFWindow.__init__(self, xaml_path)
        self.all_options = revisions
        self.filtered_options = list(revisions)
        self.selected_ids = set()
        self.selected_revisions = []
        self.match_any = False
        self.is_confirmed = False

        self._apply_theme()
        self._configure_ui()
        self._bind_events()
        self._update_search_placeholder()
        self._refresh_list()
        self._update_status()

    def _apply_theme(self):
        palette = _get_theme_palette(_is_revit_dark_theme())
        for key, value in palette.items():
            try:
                self.Resources[key].Color = ColorConverter.ConvertFromString(value)
            except Exception as ex:
                _log_debug("Failed to apply theme resource {}".format(key), ex)

    def _configure_ui(self):
        self.Title = _cfg("ui.selector_title")
        self.lblSearch.Text = _cfg("ui.search_label")
        self.lblMode.Text = _cfg("ui.mode_label")
        self.lblModeHelp.Text = _cfg("ui.mode_help")
        self.lblRevisions.Text = _cfg("ui.revisions_label")
        self.txtSearch.ToolTip = _cfg("ui.search_placeholder")
        self.searchPlaceholder.Text = _cfg("ui.search_placeholder")
        self.btnSelectVisible.Content = _cfg("buttons.select_visible")
        self.btnClearVisible.Content = _cfg("buttons.clear_visible")
        self.btnInvertVisible.Content = _cfg("buttons.invert_visible")
        self.btnCancel.Content = _cfg("buttons.cancel")
        self.btnCreate.Content = _cfg("buttons.create")
        self.rbMatchAll.Content = _cfg("options.match_all")
        self.rbMatchAny.Content = _cfg("options.match_any")
        self.rbMatchAll.IsChecked = True

    def _bind_events(self):
        self.txtSearch.TextChanged += self.on_search_changed
        self.txtSearch.GotFocus += self.on_search_focus_changed
        self.txtSearch.LostFocus += self.on_search_focus_changed
        self.btnSelectVisible.Click += self.on_select_visible
        self.btnClearVisible.Click += self.on_clear_visible
        self.btnInvertVisible.Click += self.on_invert_visible
        self.btnCancel.Click += self.on_cancel
        self.btnCreate.Click += self.on_create

    def _refresh_list(self):
        self.lstRevisions.Items.Clear()
        for option in self.filtered_options:
            row_grid = Grid()
            row_grid.HorizontalAlignment = HorizontalAlignment.Stretch
            row_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(34)))
            row_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(44)))
            row_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
            row_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(110)))
            row_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(110)))

            checkbox = CheckBox()
            checkbox.Tag = option.id_value
            checkbox.Margin = Thickness(0)
            checkbox.VerticalAlignment = VerticalAlignment.Center
            checkbox.IsChecked = option.id_value in self.selected_ids
            checkbox.Checked += self.on_revision_checked
            checkbox.Unchecked += self.on_revision_unchecked

            seq_text = TextBlock()
            seq_text.Text = option.seq
            seq_text.VerticalAlignment = VerticalAlignment.Center
            seq_text.HorizontalAlignment = HorizontalAlignment.Center

            date_text = TextBlock()
            date_text.Text = option.rev_date
            date_text.VerticalAlignment = VerticalAlignment.Center
            date_text.HorizontalAlignment = HorizontalAlignment.Center
            date_text.TextAlignment = TextAlignment.Center

            desc_text = TextBlock()
            desc_text.Text = option.desc
            desc_text.VerticalAlignment = VerticalAlignment.Center
            desc_text.TextTrimming = TextTrimming.CharacterEllipsis

            row_grid.Children.Add(_make_table_cell(checkbox, 0, Thickness(0), HorizontalAlignment.Center))
            row_grid.Children.Add(_make_table_cell(seq_text, 1, Thickness(8, 0, 8, 0)))
            row_grid.Children.Add(_make_table_cell(desc_text, 2, Thickness(2, 0, 0, 0)))
            row_grid.Children.Add(_make_table_cell(date_text, 3, Thickness(8, 0, 8, 0)))
            self.lstRevisions.Items.Add(row_grid)

    def _update_status(self):
        self.lblStatus.Text = _cfg("ui.selection_status").format(
            len(self.selected_ids),
            len(self.filtered_options),
            len(self.all_options)
        )

    def _update_search_placeholder(self):
        try:
            has_text = bool(_safe_str(self.txtSearch.Text).strip())
            self.searchPlaceholder.Visibility = Visibility.Collapsed if has_text or self.txtSearch.IsKeyboardFocused else Visibility.Visible
        except Exception:
            pass

    def _apply_filter(self):
        search_text = _safe_str(self.txtSearch.Text).strip().lower()
        if not search_text:
            self.filtered_options = list(self.all_options)
        else:
            self.filtered_options = [
                option for option in self.all_options
                if search_text in option.label.lower()
                or search_text in option.seq.lower()
                or search_text in option.desc.lower()
                or search_text in option.rev_date.lower()
            ]
        self._refresh_list()
        self._update_status()

    def _set_filtered_selection(self, is_checked):
        for option in self.filtered_options:
            if is_checked:
                self.selected_ids.add(option.id_value)
            else:
                self.selected_ids.discard(option.id_value)
        self._refresh_list()
        self._update_status()

    def on_search_changed(self, sender, args):
        self._update_search_placeholder()
        self._apply_filter()

    def on_search_focus_changed(self, sender, args):
        self._update_search_placeholder()

    def on_select_visible(self, sender, args):
        self._set_filtered_selection(True)

    def on_clear_visible(self, sender, args):
        self._set_filtered_selection(False)

    def on_invert_visible(self, sender, args):
        for option in self.filtered_options:
            if option.id_value in self.selected_ids:
                self.selected_ids.discard(option.id_value)
            else:
                self.selected_ids.add(option.id_value)
        self._refresh_list()
        self._update_status()

    def on_revision_checked(self, sender, args):
        self.selected_ids.add(sender.Tag)
        self._update_status()

    def on_revision_unchecked(self, sender, args):
        self.selected_ids.discard(sender.Tag)
        self._update_status()

    def on_cancel(self, sender, args):
        self.Close()

    def on_create(self, sender, args):
        selected = [opt.revision for opt in self.all_options if opt.id_value in self.selected_ids]
        if not selected:
            forms.alert(_cfg("messages.select_revisions_required"), title=_cfg("ui.selector_title"))
            return

        self.selected_revisions = selected
        self.match_any = bool(self.rbMatchAny.IsChecked)
        self.is_confirmed = True
        self.Close()

# ----------------------------
# HELPERS
# ----------------------------
def _safe_str(x):
    try:
        return "{}".format(x)
    except Exception as ex:
        _log_debug("Failed to stringify value", ex)
        return "<unprintable>"


def _safe_revision_id_value(revision):
    try:
        return revision.Id.IntegerValue
    except Exception as ex:
        _log_debug("Failed to read revision integer id", ex)
        return None

def _get_revision_label(rev):
    seq, desc, rev_date = _get_revision_parts(rev)

    parts = []
    if seq:
        parts.append(seq)
    if desc:
        parts.append(desc)
    if rev_date:
        parts.append(rev_date)

    if parts:
        return " | ".join(parts)

    try:
        return _cfg("messages.revision_id_prefix").format(rev.Id.IntegerValue)
    except Exception as ex:
        _log_debug("Failed to read revision id", ex)
        return _cfg("messages.revision_fallback")


def _get_revision_parts(rev):
    seq = ""
    desc = ""
    rev_date = ""

    try:
        seq = _safe_str(rev.SequenceNumber)
    except Exception as ex:
        _log_debug("Failed to read revision sequence number", ex)

    try:
        desc = _safe_str(rev.Description)
    except Exception as ex:
        _log_debug("Failed to read revision description", ex)

    try:
        rev_date = _safe_str(rev.RevisionDate)
    except Exception as ex:
        _log_debug("Failed to read revision date", ex)

    return seq, desc, rev_date


def _sort_key_for_seq(seq):
    text = _safe_str(seq).strip()
    try:
        return (0, int(text))
    except Exception:
        return (1, text.lower())


def _make_table_cell(content, column_index, margin, horizontal_alignment=None):
    cell = Border()
    cell.BorderThickness = Thickness(0)
    cell.Padding = Thickness(6, 4, 6, 4)
    cell.Margin = margin

    if horizontal_alignment is not None:
        try:
            content.HorizontalAlignment = horizontal_alignment
        except Exception:
            pass

    cell.Child = content
    Grid.SetColumn(cell, column_index)
    return cell


def _collect_revisions():
    revisions = []

    collector = DB.FilteredElementCollector(doc) \
                  .OfClass(DB.Revision) \
                  .WhereElementIsNotElementType()

    for revision in collector:
        try:
            seq, desc, rev_date = _get_revision_parts(revision)
            revisions.append(
                RevisionOption(
                    revision,
                    seq or "-",
                    desc or _cfg("messages.revision_fallback"),
                    rev_date or "-",
                    _get_revision_label(revision)
                )
            )
        except Exception as ex:
            _log_debug("Failed to collect revision option", ex)

    return sorted(revisions, key=lambda option: (option.sort_seq, option.desc.lower(), option.rev_date.lower()))


def _select_revisions_with_xaml():
    if not os.path.exists(XAML_PATH):
        forms.alert(_cfg("messages.xaml_missing").format("RevisionSelector.xaml"), title=__title__)
        return None, False

    options = _collect_revisions()
    window = RevisionSelectorWindow(XAML_PATH, options)
    window.ShowDialog()

    if not window.is_confirmed:
        return None, False

    return window.selected_revisions, window.match_any

def _count_matching_sheets(selected_revision_ids, match_any_flag):
    """Counts ViewSheets that match selected revisions (ANY or ALL)."""
    count = 0

    sheets = DB.FilteredElementCollector(doc) \
               .OfClass(DB.ViewSheet) \
               .WhereElementIsNotElementType()

    for sh in sheets:
        try:
            sh_rev_ids = sh.GetAllRevisionIds()   # ICollection[ElementId]
        except Exception as ex:
            _log_debug("Failed to get revisions for sheet {}".format(_safe_str(sh.Id)), ex)
            continue

        if not sh_rev_ids:
            continue

        if match_any_flag:
            found = False
            for rid in selected_revision_ids:
                if sh_rev_ids.Contains(rid):
                    found = True
                    break
            if found:
                count += 1
        else:
            all_found = True
            for rid in selected_revision_ids:
                if not sh_rev_ids.Contains(rid):
                    all_found = False
                    break
            if all_found:
                count += 1

    return count

def _try_get_current_sheetset_name():
    """Best-effort: reads current ViewSheetSet name from PrintManager."""
    try:
        pm = doc.PrintManager
        vss = pm.ViewSheetSetting
        # CurrentViewSheetSet is usually what create_revision_sheetset sets
        try:
            css = vss.CurrentViewSheetSet
            if css and css.Name:
                return css.Name
        except Exception as ex:
            _log_debug("Failed to read CurrentViewSheetSet", ex)
            pass

        # Fallback: sometimes InSession has the set info
        try:
            ins = vss.InSession
            if ins and ins.Name:
                return ins.Name
        except Exception as ex:
            _log_debug("Failed to read InSession sheet set", ex)
            pass
    except Exception as ex:
        _log_debug("Failed to access PrintManager ViewSheetSetting", ex)
        pass

    return None


def _collect_revision_ids(revisions):
    revision_ids = []

    for rev in revisions:
        try:
            revision_ids.append(rev.Id)
        except Exception as ex:
            _log_debug("Failed to collect selected revision id", ex)

    return revision_ids


def _collect_empty_sheets(rev_sheetset):
    empty = []

    if not rev_sheetset:
        return empty

    for sheet in rev_sheetset:
        try:
            if revit.query.is_sheet_empty(sheet):
                empty.append(sheet)
        except Exception as ex:
            _log_debug("Failed to evaluate whether sheet is empty", ex)

    return empty


def _count_sheetset_items(rev_sheetset):
    count = 0

    if not rev_sheetset:
        return count

    for _sheet in rev_sheetset:
        count += 1

    return count


def _report_empty_sheets(empty_sheets):
    if not empty_sheets:
        return

    output.print_md(_cfg("ui.empty_sheets_title"))
    for esheet in empty_sheets:
        try:
            revit.report.print_sheet(esheet)
        except Exception as ex:
            _log_debug("Failed to print empty sheet report", ex)


def _create_revision_sheetset(revisions, match_any):
    try:
        with revit.Transaction('Create Revision Sheet Set'):
            return revit.create.create_revision_sheetset(
                revisions,
                match_any=match_any
            )
    except Exception as ex:
        _log_debug("Sheet set creation failed", ex)
        raise


def _process_revision_selection(revisions, match_any):
    result = _new_result()

    if not revisions:
        return result

    result["run_status"] = "INPUT_SELECTED"
    result["match_any"] = match_any if len(revisions) > 1 else False
    result["selected_switch"] = _cfg("options.match_any") if result["match_any"] else _cfg("options.match_all")
    result["run_status"] = "MODE_SELECTED"

    sel_rev_ids = _collect_revision_ids(revisions)
    result["match_count"] = _count_matching_sheets(sel_rev_ids, result["match_any"])

    if result["match_count"] == 0:
        result["error_reason"] = _cfg("messages.no_matches")
        result["run_status"] = "NO_MATCHES"
        return result

    try:
        result["rev_sheetset"] = _create_revision_sheetset(revisions, result["match_any"])
    except Exception as ex:
        result["error_reason"] = _cfg("messages.create_failed").format(ex)
        result["run_status"] = "ERROR"
        return result

    result["run_status"] = "OK"

    name_now = _try_get_current_sheetset_name()
    if name_now:
        result["sheetset_name"] = name_now
    else:
        result["sheetset_name"] = _cfg("messages.created_name_unknown")

    result["empty_sheets"] = _collect_empty_sheets(result["rev_sheetset"])
    result["rev_sheetset_count"] = _count_sheetset_items(result["rev_sheetset"])
    _report_empty_sheets(result["empty_sheets"])

    return result


# ----------------------------
# MAIN
# ----------------------------
CONFIG = _load_config()
revisions, match_any = _select_revisions_with_xaml()
result = _process_revision_selection(revisions, match_any)

selected_switch = result["selected_switch"]
match_any = result["match_any"]
rev_sheetset = result["rev_sheetset"]
rev_sheetset_count = result["rev_sheetset_count"]
empty_sheets = result["empty_sheets"]
match_count = result["match_count"]
error_reason = result["error_reason"]
sheetset_name = result["sheetset_name"]
run_status = result["run_status"]

# ----------------------------
# FINAL REPORT
# ----------------------------
# Revision label
rev_labels = []
if revisions:
    for r in revisions:
        rev_labels.append(_get_revision_label(r))

if len(rev_labels) == 1:
    rev_one = rev_labels[0]
elif len(rev_labels) > 1:
    rev_one = _cfg("messages.multiple_revisions").format(len(rev_labels))
else:
    rev_one = _cfg("messages.no_revision")

mode_txt = selected_switch if selected_switch else _cfg("messages.mode_none")

ok_selected = bool(revisions)
ok_mode = bool(selected_switch)
ok_created = (rev_sheetset is not None)
ok_sheets = (rev_sheetset_count > 0)
ok_empty = (len(empty_sheets) == 0)

status_txt = run_status

# PRINT
output.print_md(_cfg("ui.report_title"))

# Sheet Set grande
output.print_md(_cfg("ui.report_sheetset").format(sheetset_name))

# Revision grande (subtitulo)
output.print_md(_cfg("ui.report_revision").format(rev_one))

# Resumen
output.print_md(_cfg("ui.report_summary").format(
    status_txt,
    mode_txt,
    rev_sheetset_count,
    len(empty_sheets)
))

if error_reason:
    output.print_md(_cfg("ui.report_note").format(error_reason))
elif run_status == "CANCELLED":
    output.print_md(_cfg("ui.report_cancelled"))

# If multiple revisions, list them
if len(rev_labels) > 1:
    output.print_md(_cfg("ui.report_revisions_included"))
    for rl in rev_labels:
        output.print_md("- {}".format(rl))

# GRAPHIC CHECK (single)
checks = []

def add_check(label, ok, detail):
    checks.append([label, _cfg("checks.status_ok") if ok else _cfg("checks.status_fail"), detail])

add_check(_cfg("checks.revisions_selected"), ok_selected, "{}".format(len(revisions) if revisions else 0))
add_check(_cfg("checks.match_option_picked"), ok_mode, "{}".format(mode_txt))
add_check(_cfg("checks.has_matches"), (match_count > 0), "{}".format(match_count))
add_check(_cfg("checks.sheetset_created"), ok_created, _cfg("messages.check_yes") if ok_created else _cfg("messages.check_no"))
add_check(_cfg("checks.sheets_in_set"), ok_sheets, "{}".format(rev_sheetset_count))
add_check(_cfg("checks.no_empty_sheets"), ok_empty, "{}".format(len(empty_sheets)))

output.print_md(_cfg("ui.report_graphic_check"))
output.print_table(
    table_data=checks,
    columns=["Check", "Status", "Detail"]
)

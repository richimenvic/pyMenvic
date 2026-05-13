# -*- coding: utf-8 -*-

__title__ = "Filter Standardizer"
__author__ = "Ricardo J. Mendieta"

import clr
import re

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Collections.ObjectModel import ObservableCollection
from System.IO import FileStream, FileMode, FileAccess
from System.Windows import LogicalTreeHelper
from System.Windows.Controls import Button, CheckBox, DataGrid, DataGridEditingUnit, TextBlock
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption

from Autodesk.Revit.DB import Element, FilteredElementCollector, ParameterFilterElement, ViewFamily, ViewFamilyType
from pyrevit import forms, revit, script
from lib.core.branding import get_logo_path


doc = revit.doc
XAML_FILE = script.get_bundle_file("filters_renamer.xaml")
EDIT_VIEW_TYPES_XAML = script.get_bundle_file("edit_view_types.xaml")
LOGO_FILE = get_logo_path()
STRINGS_FILE = script.get_bundle_file("strings.py")


def _load_ui_strings():
    data = {}
    try:
        execfile(STRINGS_FILE, data)
        language = data.get("LANGUAGE", "en")
        return data.get("STRINGS", {}).get(language, {})
    except Exception:
        return {}


UI_STRINGS = _load_ui_strings()

def _element_id_value(element_id):
    if element_id is None:
        return None
    for attr_name in ("Value", "IntegerValue"):
        try:
            return int(getattr(element_id, attr_name))
        except Exception:
            pass
    try:
        return int(element_id)
    except Exception:
        return str(element_id)


def _load_bitmap_from_file(file_path):
    if not file_path:
        return None
    stream = None
    try:
        stream = FileStream(file_path, FileMode.Open, FileAccess.Read)
        bitmap = BitmapImage()
        bitmap.BeginInit()
        bitmap.CacheOption = BitmapCacheOption.OnLoad
        bitmap.StreamSource = stream
        bitmap.EndInit()
        bitmap.Freeze()
        return bitmap
    except Exception:
        return None
    finally:
        try:
            if stream:
                stream.Close()
        except Exception:
            pass


def _ui_text(text):
    return UI_STRINGS.get(text, text)


def _set_localized_property(obj, property_name):
    try:
        value = getattr(obj, property_name)
        if isinstance(value, basestring):
            setattr(obj, property_name, _ui_text(value))
    except Exception:
        pass


def _localize_datagrid_columns(grid):
    try:
        for column in grid.Columns:
            _set_localized_property(column, "Header")
    except Exception:
        pass


def _localize_window_text(root):
    _set_localized_property(root, "Title")

    def walk(element):
        if isinstance(element, TextBlock):
            _set_localized_property(element, "Text")
        if isinstance(element, Button) or isinstance(element, CheckBox):
            _set_localized_property(element, "Content")
            _set_localized_property(element, "ToolTip")
        if isinstance(element, DataGrid):
            _localize_datagrid_columns(element)

        try:
            for child in LogicalTreeHelper.GetChildren(element):
                walk(child)
        except Exception:
            pass

    walk(root)


class FilterRenameRow(object):
    def __init__(self, element_id, current_name, proposed_name, section_suggestion="", include=True):
        self.ElementId = element_id
        self.CurrentName = current_name
        self.ProposedName = proposed_name
        self.SectionSuggestion = section_suggestion
        self.Include = include
        self.Status = "Ready"


class ViewTypeDebugRow(object):
    def __init__(self, type_name, view_family_name):
        self.TypeName = type_name
        self.ViewFamilyName = view_family_name


class ViewTypeEditRow(object):
    def __init__(self, element_id, current_name, new_name, view_family_name="", include=False):
        self.ElementId = element_id
        self.CurrentName = current_name
        self.NewName = new_name
        self.ViewFamilyName = view_family_name
        self.Include = include
        self.Status = "Unchecked"


class RenameFiltersWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        _localize_window_text(self)
        self._load_header_logo()
        self.filters = self._collect_filters()
        self.section_type_names = self._collect_section_type_names()
        self.all_rows = []
        self.preview_rows = ObservableCollection[object]()
        self.all_view_type_debug_rows = []
        self.PreviewGrid.ItemsSource = self.preview_rows
        self.TotalFiltersText.Text = str(len(self.filters))
        self.MatchesText.Text = "0"
        self.ReadyText.Text = "0"
        self.NoChangeText.Text = "0"
        self.IssuesText.Text = "0"
        self.ApplyButton.IsEnabled = False
        self.SearchTextBox.Text = ""
        self.ReplaceTextBox.Text = ""
        self.EnableReplaceCheckBox.IsChecked = False
        self.UppercaseCheckBox.IsChecked = False
        self.RowFilterComboBox.ItemsSource = [
            "All Rows",
            "Only Checked",
            "Only Ready",
            "Only No Change",
            "Only Issues",
            "Only Changed"
        ]
        self.RowFilterComboBox.SelectedIndex = 0
        self._sync_check_all_checkbox([])
        self._load_view_type_debug_rows()
        self._load_rows(self._build_generated_rows())

    def _load_header_logo(self):
        try:
            bitmap = _load_bitmap_from_file(LOGO_FILE)
            if bitmap is not None:
                self.HeaderLogoImage.Source = bitmap
        except Exception:
            pass

    def _collect_filters(self):
        return sorted(
            list(FilteredElementCollector(doc).OfClass(ParameterFilterElement)),
            key=lambda item: item.Name.lower()
        )

    def _element_name(self, element):
        try:
            return Element.Name.GetValue(element)
        except Exception:
            try:
                return element.Name
            except Exception:
                return ""

    def _normalize_name(self, name_text):
        value = (name_text or "").strip()
        if self.UppercaseCheckBox.IsChecked:
            value = value.upper()
        return value

    def _match_key(self, name_text):
        cleaned = re.sub(r"[^A-Z0-9]+", " ", (name_text or "").upper())
        return " ".join(cleaned.split())

    def _name_tokens(self, name_text):
        raw_tokens = [token for token in self._match_key(name_text).split() if len(token) > 1]
        synonyms = {
            "SECTION": "SECCION",
            "SECTIONS": "SECCION",
            "SECCION": "SECCION",
            "SECCIONES": "SECCION",
            "CORTE": "CORTE",
            "CORTES": "CORTE",
            "STRUCTURAL": "ESTRUCTURAL",
            "COLUMNAS": "COLUMNA",
            "COLUMN": "COLUMNA",
            "COLUMNS": "COLUMNA",
            "VIGAS": "VIGA",
            "BEAMS": "VIGA",
            "BEAM": "VIGA",
            "MUROS": "MURO",
            "WALLS": "MURO",
            "WALL": "MURO"
        }
        return [synonyms.get(token, token) for token in raw_tokens]

    def _collect_section_type_names(self):
        section_names = []
        for view_type in self._collect_section_view_types():
            section_names.append(self._element_name(view_type))
        return sorted(set(name for name in section_names if name))

    def _collect_section_view_types(self):
        section_types = []
        for view_type in FilteredElementCollector(doc).OfClass(ViewFamilyType):
            try:
                if view_type.ViewFamily == ViewFamily.Section:
                    section_types.append(view_type)
            except Exception:
                continue
        return sorted(section_types, key=lambda item: self._element_name(item).lower())

    def _load_view_type_debug_rows(self):
        debug_rows = []
        for view_type in FilteredElementCollector(doc).OfClass(ViewFamilyType):
            try:
                family_name = str(view_type.ViewFamily)
            except Exception:
                family_name = "Unknown"
            debug_rows.append(ViewTypeDebugRow(self._element_name(view_type), family_name))

        self.all_view_type_debug_rows = sorted(debug_rows, key=lambda item: (item.ViewFamilyName, item.TypeName))
        self._update_view_types_button_label()

    def _update_view_types_button_label(self):
        section_count = len([row for row in self.all_view_type_debug_rows if row.ViewFamilyName == str(ViewFamily.Section)])
        total_count = len(self.all_view_type_debug_rows)
        self.ShowViewTypesButton.Content = "VIEW TYPES: {}/{}".format(section_count, total_count)

    def ShowViewTypesButton_Click(self, sender, args):
        self._load_view_type_debug_rows()
        if not self.all_view_type_debug_rows:
            forms.alert("No se encontraron View Types en este proyecto.", title="View Type Diagnostic", exitscript=False)
            return

        lines = []
        current_family = None
        for row in self.all_view_type_debug_rows:
            if row.ViewFamilyName != current_family:
                current_family = row.ViewFamilyName
                lines.append("")
                lines.append("[{}]".format(current_family))
            lines.append(" - {}".format(row.TypeName))

        message = "\n".join(lines).strip()
        if len(message) > 6000:
            message = message[:6000] + "\n\n...lista truncada. Hay mas View Types en el proyecto."

        forms.alert(message, title="View Type Diagnostic", exitscript=False)

    def EditViewTypesButton_Click(self, sender, args):
        editor = ViewTypesEditorWindow()
        editor.ShowDialog()
        if not editor.Changed:
            return

        self.section_type_names = self._collect_section_type_names()
        self._load_view_type_debug_rows()
        for current_name, new_name in editor.RenamedPairs:
            old_generated_name = self._format_section_suggestion_name(current_name)
            new_generated_name = self._format_section_suggestion_name(new_name)
            for row in self.all_rows:
                if row.SectionSuggestion == current_name:
                    row.SectionSuggestion = new_name
                    if row.ProposedName == old_generated_name:
                        row.ProposedName = new_generated_name
        self._refresh_existing_preview()

    def _is_section_related_filter(self, filter_name):
        key = self._match_key(filter_name)
        return any(token in key for token in ("SECTION", "SECCION", "SECCIONES", "CORTE", "CORTES"))

    def _find_section_type_suggestion(self, filter_name):
        if not self._is_section_related_filter(filter_name):
            return ""

        filter_tokens = set(self._name_tokens(filter_name))
        if not filter_tokens:
            return ""

        best_name = ""
        best_score = 0

        filter_key = self._match_key(filter_name)

        for section_name in self.section_type_names:
            section_tokens = set(self._name_tokens(section_name))
            if not section_tokens:
                continue

            shared = len(filter_tokens.intersection(section_tokens))
            if shared == 0:
                continue

            score = shared * 10
            section_key = self._match_key(section_name)

            if section_key in filter_key or filter_key in section_key:
                score += 5

            if "SECTION" in section_key and "SECTION" not in filter_key:
                score -= 2

            if score > best_score:
                best_score = score
                best_name = section_name

        return best_name if best_score >= 3 else ""

    def _format_section_suggestion_name(self, suggestion):
        suggestion_text = (suggestion or "").strip()
        if not suggestion_text:
            return ""
        return "SECTION ({})".format(suggestion_text)

    def _is_suggestion_already_used(self, current_name, suggestion):
        if not suggestion:
            return False

        current_key = self._match_key(current_name)
        formatted_key = self._match_key(self._format_section_suggestion_name(suggestion))

        return bool(formatted_key and current_key == formatted_key)

    def _build_generated_rows(self):
        rows = []
        search_text = self.SearchTextBox.Text or ""
        replace_text = self.ReplaceTextBox.Text or ""
        use_replace = bool(self.EnableReplaceCheckBox.IsChecked)

        for filt in self.filters:
            current_name = filt.Name

            if search_text and search_text.lower() not in current_name.lower():
                continue

            proposed_name = current_name
            section_suggestion = self._find_section_type_suggestion(current_name)

            if use_replace and search_text:
                proposed_name = current_name.replace(search_text, replace_text)
            elif section_suggestion:
                proposed_name = self._format_section_suggestion_name(section_suggestion)

            proposed_name = self._normalize_name(proposed_name)
            rows.append(FilterRenameRow(_element_id_value(filt.Id), current_name, proposed_name, section_suggestion, False))

        return rows

    def _evaluate_rows(self, rows):
        existing_names = set(f.Name for f in self.filters)
        proposed_names = {}
        ready = 0
        no_change = 0

        for row in rows:
            current_name = (row.CurrentName or "").strip()
            proposed_name = self._normalize_name(row.ProposedName)
            row.CurrentName = current_name
            row.ProposedName = proposed_name

            if not proposed_name:
                row.Status = "Empty"
                row.Include = False
                continue

            suggestion_used = self._is_suggestion_already_used(current_name, row.SectionSuggestion)
            if proposed_name == current_name:
                row.Status = "No change"
                row.Include = False
                no_change += 1
                continue

            if suggestion_used:
                row.Status = "Used"
                row.Include = False
                no_change += 1
                continue

            if not row.Include:
                row.Status = "Unchecked"
                continue

            if proposed_name in existing_names and proposed_name != current_name:
                row.Status = "Duplicate"
                row.Include = False
                continue

            if proposed_name in proposed_names:
                row.Status = "Conflict"
                row.Include = False
                proposed_names[proposed_name].Status = "Conflict"
                proposed_names[proposed_name].Include = False
                continue

            row.Status = "Ready"
            proposed_names[proposed_name] = row
            ready += 1

        issues = 0
        for row in rows:
            if row.Status in ("Duplicate", "Conflict", "Empty"):
                issues += 1

        return ready, no_change, issues

    def _current_filter_label(self):
        return self.RowFilterComboBox.SelectedItem or "All Rows"

    def _get_filtered_rows(self):
        filter_name = self._current_filter_label()
        rows = list(self.all_rows)

        if filter_name == "Only Checked":
            return [row for row in rows if row.Include]
        if filter_name == "Only Ready":
            return [row for row in rows if row.Status == "Ready"]
        if filter_name == "Only No Change":
            return [row for row in rows if row.Status in ("No change", "Used")]
        if filter_name == "Only Issues":
            return [row for row in rows if row.Status in ("Duplicate", "Conflict", "Empty")]
        if filter_name == "Only Changed":
            return [row for row in rows if row.ProposedName != row.CurrentName]
        return rows

    def _sync_check_all_checkbox(self, rows):
        if not rows:
            self.CheckAllRowsCheckBox.IsChecked = False
            return
        self.CheckAllRowsCheckBox.IsChecked = all(row.Include for row in rows)

    def _update_visible_rows(self):
        visible_rows = self._get_filtered_rows()

        self.preview_rows.Clear()
        for row in visible_rows:
            self.preview_rows.Add(row)

        ready, no_change, issues = self._evaluate_rows(self.all_rows)
        self.MatchesText.Text = str(len(visible_rows))
        self.ReadyText.Text = str(ready)
        self.NoChangeText.Text = str(no_change)
        self.IssuesText.Text = str(issues)
        self.ApplyButton.IsEnabled = ready > 0
        self._sync_check_all_checkbox(visible_rows)
        self.PreviewGrid.Items.Refresh()

    def _load_rows(self, rows):
        self.all_rows = list(rows)
        self._evaluate_rows(self.all_rows)
        self._update_visible_rows()

    def _commit_preview_edits(self):
        try:
            self.PreviewGrid.CommitEdit(DataGridEditingUnit.Cell, True)
            self.PreviewGrid.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass

    def _refresh_existing_preview(self):
        self._commit_preview_edits()
        self._evaluate_rows(self.all_rows)
        self._update_visible_rows()

    def _reload_from_inputs(self):
        rows = self._build_generated_rows()
        if not rows:
            self.all_rows = []
            self.preview_rows.Clear()
            self.MatchesText.Text = "0"
            self.ReadyText.Text = "0"
            self.NoChangeText.Text = "0"
            self.IssuesText.Text = "0"
            self.ApplyButton.IsEnabled = False
            self._sync_check_all_checkbox([])
            return
        self._load_rows(rows)

    def PreviewButton_Click(self, sender, args):
        if self.all_rows:
            self._refresh_existing_preview()
        else:
            self._reload_from_inputs()
        if not self.all_rows:
            forms.alert("No se encontraron filtros para mostrar con los criterios actuales.", exitscript=False)

    def SearchTextBox_TextChanged(self, sender, args):
        self._reload_from_inputs()

    def ReplaceTextBox_TextChanged(self, sender, args):
        if self.EnableReplaceCheckBox.IsChecked:
            self._reload_from_inputs()

    def EnableReplaceCheckBox_Click(self, sender, args):
        self._reload_from_inputs()

    def UppercaseCheckBox_Click(self, sender, args):
        if self.filters:
            self._reload_from_inputs()

    def RowFilterComboBox_SelectionChanged(self, sender, args):
        self._update_visible_rows()

    def CheckAllRowsCheckBox_Click(self, sender, args):
        visible_rows = [row for row in self.preview_rows]
        include_all = bool(self.CheckAllRowsCheckBox.IsChecked)
        for row in visible_rows:
            row.Include = include_all
        if self.all_rows:
            self._load_rows(self.all_rows)

    def PreviewGrid_CellEditEnding(self, sender, args):
        try:
            if str(args.Column.Header) not in ("NEW NAME", _ui_text("NEW NAME")):
                return
            row = args.Row.Item
            row.ProposedName = args.EditingElement.Text
            row.Include = (self._normalize_name(row.ProposedName) != self._normalize_name(row.CurrentName))
            self._load_rows(self.all_rows)
        except Exception:
            pass

    def ApplyRowCheckBox_Click(self, sender, args):
        row = sender.DataContext
        row.Include = bool(sender.IsChecked)
        self._load_rows(self.all_rows)

    def UseSuggestionButton_Click(self, sender, args):
        row = sender.DataContext
        suggestion = (row.SectionSuggestion or "").strip()
        if not suggestion:
            forms.alert("Esta fila no tiene una sugerencia de View Type para copiar.", title="Filter Standardizer", exitscript=False)
            return

        row.ProposedName = self._normalize_name(self._format_section_suggestion_name(suggestion))
        row.Include = True
        self._load_rows(self.all_rows)

    def ResetSuggestionButton_Click(self, sender, args):
        row = sender.DataContext
        row.ProposedName = row.CurrentName
        row.Include = False
        self._load_rows(self.all_rows)

    def ApplyButton_Click(self, sender, args):
        rows = list(self.all_rows)
        if not rows:
            forms.alert("Primero genera una vista previa o carga los filtros para editar.", exitscript=False)
            return

        ready, no_change, issues = self._evaluate_rows(rows)
        self._update_visible_rows()

        if ready == 0:
            forms.alert("No hay cambios validos para aplicar. Revisa nombres vacios, duplicados o conflictos.", exitscript=False)
            return

        confirm = forms.alert(
            "Se renombraran {} filtros validos.\nSin cambio: {}\nCon problemas: {}\n\n¿Deseas continuar?".format(ready, no_change, issues),
            yes=True,
            no=True,
            exitscript=False
        )
        if not confirm:
            return

        applied = 0
        skipped = []
        by_id = {_element_id_value(f.Id): f for f in self.filters}

        with revit.Transaction("Filter Standardizer"):
            for row in rows:
                if row.Status != "Ready":
                    continue
                try:
                    by_id[row.ElementId].Name = self._normalize_name(row.ProposedName)
                    applied += 1
                except Exception as exc:
                    skipped.append("{} -> {} ({})".format(row.CurrentName, row.ProposedName, exc))

        self.filters = self._collect_filters()
        self.TotalFiltersText.Text = str(len(self.filters))

        message = "Proceso terminado.\n\nRenombrados: {}".format(applied)
        if skipped:
            message += "\nOmitidos: {}\n\n{}".format(len(skipped), "\n".join(skipped[:10]))

        self._reload_from_inputs()
        forms.alert(message, title="Filter Standardizer", exitscript=False)

    def CloseButton_Click(self, sender, args):
        self.Close()


class ViewTypesEditorWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, EDIT_VIEW_TYPES_XAML)
        _localize_window_text(self)
        self.Changed = False
        self.RenamedPairs = []
        self.all_rows = []
        self.rows = ObservableCollection[object]()
        self.ViewTypesGrid.ItemsSource = self.rows
        self.ViewFamilyFilterComboBox.ItemsSource = ["All View Families"]
        self.ViewFamilyFilterComboBox.SelectedIndex = 0
        self._load_rows()

    def _element_name(self, element):
        try:
            return Element.Name.GetValue(element)
        except Exception:
            try:
                return element.Name
            except Exception:
                return ""

    def _collect_view_types(self):
        view_types = []
        for view_type in FilteredElementCollector(doc).OfClass(ViewFamilyType):
            view_types.append(view_type)
        return sorted(view_types, key=lambda item: (str(item.ViewFamily), self._element_name(item).lower()))

    def _commit_grid_edits(self):
        try:
            self.ViewTypesGrid.CommitEdit(DataGridEditingUnit.Cell, True)
            self.ViewTypesGrid.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass

    def _evaluate_rows(self):
        existing_names = set(self._element_name(view_type) for view_type in FilteredElementCollector(doc).OfClass(ViewFamilyType))
        proposed_names = {}
        ready = 0
        issues = 0

        for row in self.all_rows:
            current_name = (row.CurrentName or "").strip()
            new_name = (row.NewName or "").strip()
            row.CurrentName = current_name
            row.NewName = new_name

            if not new_name:
                row.Status = "Empty"
                row.Include = False
                issues += 1
                continue
            if new_name == current_name:
                row.Status = "No change"
                row.Include = False
                continue
            if new_name in existing_names and new_name != current_name:
                row.Status = "Duplicate"
                row.Include = False
                issues += 1
                continue
            if new_name in proposed_names:
                row.Status = "Conflict"
                row.Include = False
                proposed_names[new_name].Status = "Conflict"
                proposed_names[new_name].Include = False
                issues += 1
                continue
            if not row.Include:
                row.Status = "Unchecked"
                continue

            row.Status = "Ready"
            proposed_names[new_name] = row
            ready += 1

        self.ReadyTypesText.Text = str(ready)
        self.IssueTypesText.Text = str(issues)
        self._update_visible_rows()
        return ready, issues

    def _update_visible_rows(self):
        search_text = (self.SearchTypesTextBox.Text or "").strip().lower()
        selected_family = self.ViewFamilyFilterComboBox.SelectedItem or "All View Families"
        self.rows.Clear()
        for row in self.all_rows:
            if search_text and search_text not in (row.NewName or "").lower() and search_text not in (row.CurrentName or "").lower() and search_text not in (row.ViewFamilyName or "").lower():
                continue
            if selected_family != "All View Families" and row.ViewFamilyName != selected_family:
                continue
            self.rows.Add(row)
        self.VisibleTypesText.Text = str(len(self.rows))
        self.ViewTypesGrid.Items.Refresh()

    def _load_rows(self):
        self.all_rows = []
        for view_type in self._collect_view_types():
            current_name = self._element_name(view_type)
            try:
                family_name = str(view_type.ViewFamily)
            except Exception:
                family_name = "Unknown"
            self.all_rows.append(ViewTypeEditRow(_element_id_value(view_type.Id), current_name, current_name, family_name, False))
        families = sorted(set(row.ViewFamilyName for row in self.all_rows))
        self.ViewFamilyFilterComboBox.ItemsSource = ["All View Families"] + families
        self.ViewFamilyFilterComboBox.SelectedIndex = 0
        self.TotalTypesText.Text = str(len(self.all_rows))
        self._evaluate_rows()

    def ViewTypesGrid_CellEditEnding(self, sender, args):
        try:
            if str(args.Column.Header) not in ("TYPE NAME", _ui_text("TYPE NAME")):
                return
            row = args.Row.Item
            row.NewName = args.EditingElement.Text
            row.Include = (row.NewName or "").strip() != (row.CurrentName or "").strip()
            self._evaluate_rows()
        except Exception:
            pass

    def SearchTypesTextBox_TextChanged(self, sender, args):
        self._update_visible_rows()

    def ViewFamilyFilterComboBox_SelectionChanged(self, sender, args):
        self._update_visible_rows()

    def ViewTypeApplyCheckBox_Click(self, sender, args):
        row = sender.DataContext
        row.Include = bool(sender.IsChecked)
        self._evaluate_rows()

    def CheckChangedTypesButton_Click(self, sender, args):
        self._commit_grid_edits()
        for row in self.all_rows:
            row.Include = (row.NewName or "").strip() != (row.CurrentName or "").strip()
        self._evaluate_rows()

    def ValidateTypesButton_Click(self, sender, args):
        self._commit_grid_edits()
        self._evaluate_rows()

    def ResetTypeButton_Click(self, sender, args):
        row = sender.DataContext
        row.NewName = row.CurrentName
        row.Include = False
        self._evaluate_rows()

    def ApplyTypesButton_Click(self, sender, args):
        self._commit_grid_edits()
        ready, issues = self._evaluate_rows()
        if ready == 0:
            forms.alert("No hay View Types validos para renombrar.", title="Edit View Types", exitscript=False)
            return

        confirm = forms.alert(
            "Se renombraran {} View Types.\nCon problemas: {}\n\n¿Deseas continuar?".format(ready, issues),
            yes=True,
            no=True,
            exitscript=False
        )
        if not confirm:
            return

        by_id = {_element_id_value(view_type.Id): view_type for view_type in self._collect_view_types()}
        renamed = []
        with revit.Transaction("Rename View Types"):
            for row in self.all_rows:
                if row.Status != "Ready":
                    continue
                old_name = row.CurrentName
                new_name = row.NewName
                by_id[row.ElementId].Name = new_name
                renamed.append((old_name, new_name))

        self.Changed = self.Changed or bool(renamed)
        self.RenamedPairs.extend(renamed)
        self._load_rows()
        forms.alert("View Types renombrados: {}".format(len(renamed)), title="Edit View Types", exitscript=False)

    def CloseTypesButton_Click(self, sender, args):
        self.Close()


RenameFiltersWindow().ShowDialog()




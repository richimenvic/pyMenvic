# -*- coding: utf-8 -*-

__title__ = "Rename Filters"
__author__ = "Ricardo J. Mendieta"

import clr
import re

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Collections.ObjectModel import ObservableCollection

from Autodesk.Revit.DB import Element, FilteredElementCollector, ParameterFilterElement, ViewFamily, ViewFamilyType
from pyrevit import forms, revit, script


doc = revit.doc
XAML_FILE = script.get_bundle_file("filters_renamer.xaml")


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


class RenameFiltersWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self.filters = self._collect_filters()
        self.section_type_names = self._collect_section_type_names()
        self.all_rows = []
        self.preview_rows = ObservableCollection[object]()
        self.view_type_debug_rows = ObservableCollection[object]()
        self.PreviewGrid.ItemsSource = self.preview_rows
        self.ViewTypesDebugGrid.ItemsSource = self.view_type_debug_rows
        self.TotalFiltersText.Text = str(len(self.filters))
        self.MatchesText.Text = "0"
        self.ReadyText.Text = "0"
        self.NoChangeText.Text = "0"
        self.IssuesText.Text = "0"
        self.SearchTextBox.Text = ""
        self.ReplaceTextBox.Text = ""
        self.EnableReplaceCheckBox.IsChecked = False
        self.UseSectionTypesCheckBox.IsChecked = False
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
        for view_type in FilteredElementCollector(doc).OfClass(ViewFamilyType):
            try:
                if view_type.ViewFamily == ViewFamily.Section:
                    section_names.append(self._element_name(view_type))
            except Exception:
                continue
        return sorted(set(name for name in section_names if name))

    def _load_view_type_debug_rows(self):
        self.view_type_debug_rows.Clear()
        debug_rows = []
        for view_type in FilteredElementCollector(doc).OfClass(ViewFamilyType):
            try:
                family_name = str(view_type.ViewFamily)
            except Exception:
                family_name = "Unknown"
            debug_rows.append(ViewTypeDebugRow(self._element_name(view_type), family_name))

        for row in sorted(debug_rows, key=lambda item: (item.ViewFamilyName, item.TypeName)):
            self.view_type_debug_rows.Add(row)

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

    def _build_generated_rows(self):
        rows = []
        search_text = (self.SearchTextBox.Text or "").strip()
        replace_text = self.ReplaceTextBox.Text or ""
        use_replace = bool(self.EnableReplaceCheckBox.IsChecked)
        use_section_types = bool(self.UseSectionTypesCheckBox.IsChecked)

        for filt in self.filters:
            current_name = filt.Name

            if search_text and search_text.lower() not in current_name.lower():
                continue

            proposed_name = current_name
            section_suggestion = self._find_section_type_suggestion(current_name)

            if use_section_types and section_suggestion and self._match_key(section_suggestion) != self._match_key(current_name):
                proposed_name = section_suggestion
            if use_replace and search_text:
                proposed_name = current_name.replace(search_text, replace_text)

            proposed_name = self._normalize_name(proposed_name)
            rows.append(FilterRenameRow(filt.Id.IntegerValue, current_name, proposed_name, section_suggestion, True))

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

            if not row.Include:
                row.Status = "Excluded"
                continue

            if not proposed_name:
                row.Status = "Empty"
                continue

            if proposed_name == current_name:
                row.Status = "No change"
                no_change += 1
                continue

            if proposed_name in existing_names and proposed_name != current_name:
                row.Status = "Duplicate"
                continue

            if proposed_name in proposed_names:
                row.Status = "Conflict"
                proposed_names[proposed_name].Status = "Conflict"
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
            return [row for row in rows if row.Status == "No change"]
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
        self._sync_check_all_checkbox(visible_rows)
        self.PreviewGrid.Items.Refresh()

    def _load_rows(self, rows):
        self.all_rows = list(rows)
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
            self._sync_check_all_checkbox([])
            return
        self._load_rows(rows)

    def PreviewButton_Click(self, sender, args):
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

    def UseSectionTypesCheckBox_Click(self, sender, args):
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
        by_id = {f.Id.IntegerValue: f for f in self.filters}

        with revit.Transaction("Rename Filters"):
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
        forms.alert(message, title="Rename Filters", exitscript=False)

    def CloseButton_Click(self, sender, args):
        self.Close()


RenameFiltersWindow().ShowDialog()


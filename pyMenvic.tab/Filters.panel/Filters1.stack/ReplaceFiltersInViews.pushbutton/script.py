# -*- coding: utf-8 -*-



import os
import sys

try:
    from lib.core.branding import get_logo_path
except ImportError:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    while True:
        if os.path.basename(current_dir).lower() == "pymenvic.extension":
            lib_dir = os.path.join(current_dir, "lib")
            if lib_dir not in sys.path:
                sys.path.append(lib_dir)
            break
        parent_dir = os.path.dirname(current_dir)
        if parent_dir == current_dir:
            break
        current_dir = parent_dir
    from core.branding import get_logo_path

__title__ = "Replace Filters In Views"
__author__ = "OpenAI Codex"

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Collections.ObjectModel import ObservableCollection
from System.IO import FileStream, FileMode, FileAccess
from System.Windows import LogicalTreeHelper
from System.Windows.Controls import Button, CheckBox, DataGrid, DataGridEditingUnit, TextBlock
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption

from Autodesk.Revit.DB import (
    Element,
    ElementId,
    FilteredElementCollector,
    ParameterFilterElement,
    ScheduleSheetInstance,
    View,
    Viewport,
)
from pyrevit import forms, revit, script


XAML_FILE = script.get_bundle_file("replace_filters_in_views.xaml")
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
    """Return a stable numeric value for Revit ElementId across Revit versions."""
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


class FilterOption(object):
    def __init__(self, element_id, name):
        self.ElementId = element_id
        self.Name = name

    def __str__(self):
        return self.Name


class FilterUsageRow(object):
    def __init__(self, view_id, view_name, view_kind, sheet_info, is_template, already_has_target, source_enabled, source_visible, target_enabled, target_visible):
        self.ViewId = view_id
        self.ViewName = view_name
        self.ViewKind = view_kind
        self.SheetInfo = sheet_info
        self.IsTemplate = is_template
        self.AlreadyHasTarget = already_has_target
        self.SourceEnabled = source_enabled
        self.SourceVisible = source_visible
        self.TargetEnabled = target_enabled
        self.TargetVisible = target_visible
        self.Include = not already_has_target
        self.Status = "Ready"
        self.Match = ""
        self.Action = ""


class ReplaceFiltersWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        _localize_window_text(self)
        self._load_header_logo()
        self.filters = self._collect_filters()
        self.filter_options_by_name = {}
        for filter_option in self.filters:
            self.filter_options_by_name[filter_option.Name] = filter_option
        self.sheet_map = self._build_sheet_map()
        self.all_rows = []
        self.preview_rows = ObservableCollection[object]()
        self.ResultsGrid.ItemsSource = self.preview_rows
        filter_names = [filter_option.Name for filter_option in self.filters]
        self.SourceFilterComboBox.ItemsSource = filter_names
        self.TargetFilterComboBox.ItemsSource = filter_names
        self.IncludeTemplatesCheckBox.IsChecked = True
        self.IncludeViewsCheckBox.IsChecked = True
        self.MergeExistingCheckBox.IsChecked = True
        self.CopyVisibilityCheckBox.IsChecked = True
        self.CopyEnabledCheckBox.IsChecked = True
        self.CheckAllRowsCheckBox.IsChecked = False
        self.TotalText.Text = str(len(self.filters))
        self.VisibleText.Text = "0"
        self.ReadyText.Text = "0"
        self.IssuesText.Text = "0"
        if len(self.filters) >= 1:
            self.SourceFilterComboBox.SelectedIndex = 0
        if len(self.filters) >= 2:
            self.TargetFilterComboBox.SelectedIndex = 1
        elif len(self.filters) == 1:
            self.TargetFilterComboBox.SelectedIndex = 0
        self._reload_rows()

    def _status_text(self, text):
        return _ui_text(text)

    def _collect_filters(self):
        items = []
        for filt in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
            items.append(FilterOption(filt.Id, self._filter_name(filt)))
        return sorted(items, key=lambda item: item.Name.lower())

    def _load_header_logo(self):
        stream = None
        try:
            stream = FileStream(LOGO_FILE, FileMode.Open, FileAccess.Read)
            bitmap = BitmapImage()
            bitmap.BeginInit()
            bitmap.CacheOption = BitmapCacheOption.OnLoad
            bitmap.StreamSource = stream
            bitmap.EndInit()
            try:
                bitmap.Freeze()
            except Exception:
                pass
            self.HeaderLogoImage.Source = bitmap
        except Exception:
            pass
        finally:
            if stream:
                try:
                    stream.Close()
                except Exception:
                    pass

    def _filter_name(self, filter_element):
        try:
            return filter_element.Name
        except Exception:
            return self._element_name(filter_element)

    def _element_name(self, element):
        try:
            return Element.Name.GetValue(element)
        except Exception:
            try:
                return element.Name
            except Exception:
                return ""

    def _view_kind_name(self, view):
        try:
            return str(view.ViewType)
        except Exception:
            return "Unknown"

    def _sheet_label(self, sheet):
        try:
            number = sheet.SheetNumber or "?"
        except Exception:
            number = "?"
        try:
            name = sheet.Name or ""
        except Exception:
            name = ""
        label = "{} - {}".format(number, name).strip()
        return label.strip("- ").strip()

    def _build_sheet_map(self):
        sheet_names = {}
        for view in FilteredElementCollector(doc).OfClass(View):
            try:
                if str(view.ViewType) == "DrawingSheet":
                    sheet_names[_element_id_value(view.Id)] = self._sheet_label(view)
            except Exception:
                continue

        placed = {}

        for viewport in FilteredElementCollector(doc).OfClass(Viewport):
            try:
                view_id = _element_id_value(viewport.ViewId)
                sheet_id = _element_id_value(viewport.SheetId)
                if sheet_id in sheet_names:
                    placed.setdefault(view_id, []).append(sheet_names[sheet_id])
            except Exception:
                continue

        try:
            schedule_instances = FilteredElementCollector(doc).OfClass(ScheduleSheetInstance)
        except Exception:
            schedule_instances = []

        for schedule_instance in schedule_instances:
            try:
                if schedule_instance.IsTitleblockRevisionSchedule:
                    continue
            except Exception:
                pass
            try:
                owner_view_id = _element_id_value(schedule_instance.OwnerViewId)
                schedule_id = _element_id_value(schedule_instance.ScheduleId)
                if owner_view_id in sheet_names:
                    placed.setdefault(schedule_id, []).append(sheet_names[owner_view_id])
            except Exception:
                continue

        for view_id, labels in placed.items():
            unique_labels = []
            seen = set()
            for label in labels:
                if label in seen:
                    continue
                unique_labels.append(label)
                seen.add(label)
            placed[view_id] = unique_labels

        return placed

    def _selected_source(self):
        selected_name = self.SourceFilterComboBox.SelectedItem
        return self.filter_options_by_name.get(selected_name)

    def _selected_target(self):
        selected_name = self.TargetFilterComboBox.SelectedItem
        return self.filter_options_by_name.get(selected_name)

    def _include_templates(self):
        return bool(self.IncludeTemplatesCheckBox.IsChecked)

    def _include_views(self):
        return bool(self.IncludeViewsCheckBox.IsChecked)

    def _allow_merge_existing(self):
        return bool(self.MergeExistingCheckBox.IsChecked)

    def _copy_source_visibility(self):
        return bool(self.CopyVisibilityCheckBox.IsChecked)

    def _copy_source_enabled(self):
        return bool(self.CopyEnabledCheckBox.IsChecked)

    def _can_use_view(self, view):
        try:
            if view.IsTemplate and not self._include_templates():
                return False
            if not view.IsTemplate and not self._include_views():
                return False
        except Exception:
            return False

        try:
            if view.ViewType.ToString() in ("ProjectBrowser", "SystemBrowser", "Undefined", "Internal"):
                return False
        except Exception:
            pass

        try:
            view.GetFilters()
            return True
        except Exception:
            return False

    def _build_usage_rows(self):
        source = self._selected_source()
        target = self._selected_target()
        rows = []

        if source is None:
            return rows

        source_id = source.ElementId
        target_id = target.ElementId if target else None

        for view in FilteredElementCollector(doc).OfClass(View):
            if not self._can_use_view(view):
                continue

            try:
                filter_ids = list(view.GetFilters())
            except Exception:
                continue

            has_source = any(_element_id_value(fid) == _element_id_value(source_id) for fid in filter_ids)
            if not has_source:
                continue

            has_target = False
            if target_id is not None:
                has_target = any(_element_id_value(fid) == _element_id_value(target_id) for fid in filter_ids)

            source_visible = self._get_filter_visibility(view, source_id)
            source_enabled = self._get_filter_enabled(view, source_id)
            if source_enabled is None:
                source_enabled = True

            target_visible = False
            target_enabled = False
            if has_target and target_id is not None:
                target_visible = self._get_filter_visibility(view, target_id)
                target_enabled = self._get_filter_enabled(view, target_id)
                if target_enabled is None:
                    target_enabled = True

            sheet_labels = self.sheet_map.get(_element_id_value(view.Id), [])
            rows.append(
                FilterUsageRow(
                    _element_id_value(view.Id),
                    self._element_name(view),
                    self._view_kind_name(view),
                    ", ".join(sheet_labels) if sheet_labels else "-",
                    bool(getattr(view, "IsTemplate", False)),
                    has_target,
                    source_enabled,
                    source_visible,
                    target_enabled,
                    target_visible,
                )
            )

        return sorted(rows, key=lambda row: (row.IsTemplate, row.SheetInfo == "-", row.SheetInfo, row.ViewKind, row.ViewName.lower()))

    def _evaluate_rows(self):
        source = self._selected_source()
        target = self._selected_target()
        ready = 0
        issues = 0

        for row in self.all_rows:
            row.Status = self._status_text("Ready")
            row.Match = ""
            row.Action = ""

            if source is None:
                row.Status = self._status_text("Select source")
                row.Include = False
                row.Action = _ui_text("Select")
                issues += 1
                continue

            if target is None:
                row.Status = self._status_text("Select target")
                row.Include = False
                row.Action = _ui_text("Select")
                issues += 1
                continue

            if _element_id_value(source.ElementId) == _element_id_value(target.ElementId):
                row.Status = self._status_text("Same filter")
                row.Include = False
                row.Match = _ui_text("Same")
                row.Action = _ui_text("Skip")
                issues += 1
                continue

            if row.AlreadyHasTarget:
                differs_active = bool(row.SourceEnabled) != bool(row.TargetEnabled)
                differs_visible = bool(row.SourceVisible) != bool(row.TargetVisible)
                if differs_active and differs_visible:
                    row.Match = _ui_text("Active + Visible")
                elif differs_active:
                    row.Match = _ui_text("Active")
                elif differs_visible:
                    row.Match = _ui_text("Visible")
                else:
                    row.Match = _ui_text("Same")

                if self._allow_merge_existing():
                    row.Action = _ui_text("Merge")
                    if row.Include:
                        row.Status = self._status_text("Merge")
                        ready += 1
                    else:
                        row.Status = self._status_text("Target exists")
                    continue

                row.Status = self._status_text("Target exists")
                row.Include = False
                row.Action = _ui_text("Keep target")
                issues += 1
                continue

            if row.Include:
                ready += 1
                row.Action = _ui_text("Replace")
            else:
                row.Status = self._status_text("Skipped")
                row.Action = _ui_text("Skip")
            row.Match = _ui_text("Missing")

        return ready, issues

    def _sync_check_all_checkbox(self, rows):
        if not rows:
            self.CheckAllRowsCheckBox.IsChecked = False
            return
        self.CheckAllRowsCheckBox.IsChecked = all(row.Include for row in rows)

    def _update_visible_rows(self):
        search_text = (self.SearchTextBox.Text or "").strip().lower()
        self.preview_rows.Clear()
        for row in self.all_rows:
            if search_text:
                haystack = " | ".join([
                    row.ViewName or "",
                    row.ViewKind or "",
                    row.SheetInfo or "",
                    row.Match or "",
                    row.Action or "",
                    row.Status or "",
                ]).lower()
                if search_text not in haystack:
                    continue
            self.preview_rows.Add(row)

        ready, issues = self._evaluate_rows()
        self.VisibleText.Text = str(len(self.preview_rows))
        self.ReadyText.Text = str(ready)
        self.IssuesText.Text = str(issues)
        self._sync_check_all_checkbox([row for row in self.preview_rows])
        self.ResultsGrid.Items.Refresh()

    def _reload_rows(self):
        self.all_rows = self._build_usage_rows()
        self._row_map_by_view_id = dict((row.ViewId, row) for row in self.all_rows)
        self._update_visible_rows()

    def _commit_grid_edits(self):
        try:
            self.ResultsGrid.CommitEdit(DataGridEditingUnit.Cell, True)
            self.ResultsGrid.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass

    def _get_filter_overrides(self, view, filter_id):
        try:
            return view.GetFilterOverrides(filter_id)
        except Exception:
            return None

    def _get_filter_visibility(self, view, filter_id):
        try:
            return view.GetFilterVisibility(filter_id)
        except Exception:
            return True

    def _get_filter_enabled(self, view, filter_id):
        for method_name in ("GetIsFilterEnabled", "IsFilterEnabled"):
            try:
                return getattr(view, method_name)(filter_id)
            except Exception:
                continue
        return None

    def _set_filter_enabled(self, view, filter_id, enabled_value):
        if enabled_value is None:
            return
        for method_name in ("SetIsFilterEnabled", "SetFilterEnabled"):
            try:
                getattr(view, method_name)(filter_id, enabled_value)
                return
            except Exception:
                continue

    def _get_filter_order(self, view):
        for method_name in ("GetOrderedFilters",):
            try:
                return list(getattr(view, method_name)())
            except Exception:
                continue
        try:
            return list(view.GetFilters())
        except Exception:
            return []

    def _move_filter_to_index(self, view, filter_id, index):
        for method_name in ("MoveFilter",):
            try:
                getattr(view, method_name)(filter_id, index)
                return True
            except Exception:
                continue
        return False

    def _replace_filter_on_view(self, view, source_id, target_id):
        ordered_filters = self._get_filter_order(view)
        source_index = -1
        target_exists = False

        for idx, filter_id in enumerate(ordered_filters):
            if _element_id_value(filter_id) == _element_id_value(source_id):
                source_index = idx
            if _element_id_value(filter_id) == _element_id_value(target_id):
                target_exists = True

        if source_index < 0:
            return "No source"

        visibility = self._get_filter_visibility(view, source_id)
        overrides = self._get_filter_overrides(view, source_id)
        enabled_value = self._get_filter_enabled(view, source_id)

        if not target_exists:
            view.AddFilter(target_id)
            self._move_filter_to_index(view, target_id, source_index)

        if overrides is not None:
            view.SetFilterOverrides(target_id, overrides)

        desired_visibility = None
        desired_enabled = None

        try:
            row = self._row_map_by_view_id.get(_element_id_value(view.Id))
        except Exception:
            row = None

        if target_exists and row is not None:
            desired_visibility = row.TargetVisible
            desired_enabled = row.TargetEnabled
        else:
            if self._copy_source_visibility():
                desired_visibility = visibility
            if self._copy_source_enabled():
                desired_enabled = enabled_value

        if desired_visibility is not None:
            try:
                view.SetFilterVisibility(target_id, desired_visibility)
            except Exception:
                pass

        if desired_enabled is not None:
            self._set_filter_enabled(view, target_id, desired_enabled)
        view.RemoveFilter(source_id)

        return "Merged" if target_exists else "Replaced"

    def RefreshButton_Click(self, sender, args):
        self._reload_rows()
        if not self.all_rows:
            forms.alert(_ui_text("No views or templates were found using that filter."), title=_ui_text("Replace Filters"), exitscript=False)

    def FilterComboBox_SelectionChanged(self, sender, args):
        self._reload_rows()

    def ScopeCheckBox_Click(self, sender, args):
        self._reload_rows()

    def SearchTextBox_TextChanged(self, sender, args):
        self._update_visible_rows()

    def CheckAllRowsCheckBox_Click(self, sender, args):
        include_all = bool(self.CheckAllRowsCheckBox.IsChecked)
        for row in self.all_rows:
            row.Include = include_all
        self._update_visible_rows()

    def IncludeRowCheckBox_Click(self, sender, args):
        self._update_visible_rows()

    def ApplyButton_Click(self, sender, args):
        self._commit_grid_edits()
        ready, issues = self._evaluate_rows()
        self._update_visible_rows()

        source = self._selected_source()
        target = self._selected_target()

        if source is None or target is None:
            forms.alert(_ui_text("Select a source filter and a target filter."), title=_ui_text("Replace Filters"), exitscript=False)
            return

        if _element_id_value(source.ElementId) == _element_id_value(target.ElementId):
            forms.alert(_ui_text("The source filter and target filter cannot be the same."), title=_ui_text("Replace Filters"), exitscript=False)
            return

        if ready == 0:
            forms.alert(_ui_text("There are no ready views to replace. Review the list or enable merge."), title=_ui_text("Replace Filters"), exitscript=False)
            return

        confirm = forms.alert(
            _ui_text("Replace '{0}' with '{1}' in {2} views/templates.\nRows with issues: {3}\n\nDo you want to continue?").format(
                source.Name,
                target.Name,
                ready,
                issues,
            ),
            yes=True,
            no=True,
            exitscript=False,
        )
        if not confirm:
            return

        replaced = 0
        merged = 0
        skipped = []

        with revit.Transaction("Replace Filters In Views"):
            for row in self.all_rows:
                if row.Status not in (self._status_text("Ready"), self._status_text("Merge")) or not row.Include:
                    continue

                view = doc.GetElement(ElementId(row.ViewId))
                try:
                    result = self._replace_filter_on_view(view, source.ElementId, target.ElementId)
                    if result == "Merged":
                        merged += 1
                    elif result == "Replaced":
                        replaced += 1
                except Exception as exc:
                    skipped.append("{} ({})".format(row.ViewName, exc))
                    row.Status = self._status_text("Error")
                    row.Note = str(exc)

        self.sheet_map = self._build_sheet_map()
        self._reload_rows()

        summary = _ui_text("Process finished.\n\nReplaced: {0}\nMerged: {1}").format(replaced, merged)
        if skipped:
            summary += _ui_text("\nErrors: {0}\n\n{1}").format(len(skipped), "\n".join(skipped[:12]))
        forms.alert(summary, title=_ui_text("Replace Filters"), exitscript=False)

    def CloseButton_Click(self, sender, args):
        self.Close()


ReplaceFiltersWindow().ShowDialog()

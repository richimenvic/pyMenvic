# -*- coding: utf-8 -*-

__title__ = "Manage Filters"
__author__ = "OpenAI Codex"

import clr
import re

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import Action
from System.IO import FileStream, FileMode, FileAccess
from System.Collections.ObjectModel import ObservableCollection
from System import Enum
from System.Windows import LogicalTreeHelper
from System.Windows.Controls import Button, CheckBox, DataGrid, DataGridEditingUnit, TextBlock
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption

from Autodesk.Revit.DB import BuiltInParameter, Element, ElementId, FilteredElementCollector, LabelUtils, ParameterFilterElement, View
from pyrevit import forms, revit, script
from lib.core.branding import get_logo_path


doc = revit.doc
XAML_FILE = script.get_bundle_file("manage_filters.xaml")
LOGO_FILE = get_logo_path()
STRINGS_FILE = script.get_bundle_file("strings.py")
logger = script.get_logger()


def _load_ui_strings():
    data = {}
    try:
        execfile(STRINGS_FILE, data)
        language = data.get("LANGUAGE", "en")
        return data.get("STRINGS", {}).get(language, {})
    except Exception:
        return {}


UI_STRINGS = _load_ui_strings()


def _ui_text(text):
    return UI_STRINGS.get(text, text)


def _element_id_value(element_id):
    """Return a stable integer value for ElementId across Revit versions."""
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


class FilterRow(object):
    def __init__(self, filter_id, filter_name, used_views, used_templates, duplicate_group, duplicate_type, category_summary, content_signature):
        self.FilterId = filter_id
        self.OriginalName = filter_name
        self.FilterName = filter_name
        self.UsedViewsCount = len(used_views)
        self.UsedTemplatesCount = len(used_templates)
        self.TotalUses = self.UsedViewsCount + self.UsedTemplatesCount
        self.UsedViews = used_views
        self.UsedTemplates = used_templates
        self.DuplicateGroup = duplicate_group
        self.DuplicateType = duplicate_type
        self.DuplicateGroupDisplay = duplicate_group or "-"
        self.DuplicateTypeDisplay = duplicate_type or "Not duplicate"
        self.CategorySummary = category_summary
        self.ContentSignature = content_signature
        self.Status = self._build_status()
        self.Include = self.TotalUses == 0

    @property
    def HasPendingRename(self):
        return self._normalized_name(self.FilterName) != self._normalized_name(self.OriginalName)

    def _normalized_name(self, value):
        return (value or "").strip()

    def _build_status(self):
        if self.HasPendingRename:
            return "Rename Pending"
        is_duplicate = bool(self.DuplicateType)
        is_unused = self.TotalUses == 0
        if is_duplicate and is_unused:
            return "Duplicate + Unused"
        if is_duplicate:
            return "Duplicate"
        if is_unused:
            return "Unused"
        return "Used"


class ManageFiltersWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        _localize_window_text(self)
        self._load_header_logo()
        self._pending_edit_row = None
        self.all_rows = []
        self.preview_rows = ObservableCollection[object]()
        self.ResultsGrid.ItemsSource = self.preview_rows
        self.ScopeComboBox.ItemsSource = [
            _ui_text("All Filters"),
            _ui_text("Unused Only"),
            _ui_text("Used Only"),
            _ui_text("Duplicates Only"),
        ]
        self.ScopeComboBox.SelectedIndex = 0
        self._reload_rows()

    def _load_header_logo(self):
        stream = None
        try:
            stream = FileStream(LOGO_FILE, FileMode.Open, FileAccess.Read)
            image = BitmapImage()
            image.BeginInit()
            image.StreamSource = stream
            image.CacheOption = BitmapCacheOption.OnLoad
            image.EndInit()
            image.Freeze()
            self.HeaderLogoImage.Source = image
        except Exception:
            pass
        finally:
            if stream is not None:
                try:
                    stream.Close()
                except Exception:
                    pass

    def _element_name(self, element):
        try:
            return Element.Name.GetValue(element)
        except Exception:
            try:
                return element.Name
            except Exception:
                return ""

    def _match_key(self, text):
        cleaned = re.sub(r"[^A-Z0-9]+", " ", (text or "").upper())
        return " ".join(cleaned.split())

    def _category_signature(self, filter_element):
        try:
            category_ids = list(filter_element.GetCategories())
        except Exception:
            category_ids = []
        values = []
        for category_id in category_ids:
            try:
                values.append(str(_element_id_value(category_id)))
            except Exception:
                values.append(str(category_id))
        return "|".join(sorted(values))

    def _category_summary(self, filter_element):
        try:
            category_ids = list(filter_element.GetCategories())
        except Exception:
            category_ids = []

        names = []
        for category_id in category_ids:
            try:
                category = doc.Settings.Categories.get_Item(category_id)
                if category:
                    names.append(category.Name)
                    continue
            except Exception:
                pass
            try:
                names.append(str(_element_id_value(category_id)))
            except Exception:
                names.append(str(category_id))

        return ", ".join(sorted(names))

    def _safe_class_name(self, obj):
        try:
            return obj.GetType().Name
        except Exception:
            try:
                return obj.__class__.__name__
            except Exception:
                return str(type(obj))

    def _parameter_label(self, parameter_id):
        if parameter_id is None:
            return "Unknown parameter"

        try:
            element = doc.GetElement(parameter_id)
            if element is not None:
                return self._element_name(element)
        except Exception:
            pass

        try:
            integer_value = _element_id_value(parameter_id)
        except Exception:
            try:
                integer_value = int(parameter_id)
            except Exception:
                return str(parameter_id)

        try:
            bip = Enum.ToObject(BuiltInParameter, integer_value)
            return LabelUtils.GetLabelFor(bip)
        except Exception:
            return "ParameterId {}".format(integer_value)

    def _extract_rule_parameter_id(self, rule):
        for method_name in ("GetRuleParameter",):
            try:
                return getattr(rule, method_name)()
            except Exception:
                continue
        for property_name in ("RuleParameter", "ParameterId"):
            try:
                return getattr(rule, property_name)
            except Exception:
                continue
        return None

    def _extract_rule_value(self, rule):
        for method_name in (
            "GetStringValue",
            "GetValue",
            "GetIntegerValue",
            "GetDoubleValue",
            "GetElementIdValue",
        ):
            try:
                value = getattr(rule, method_name)()
                return value
            except Exception:
                continue
        return None

    def _describe_rule(self, rule, indent=""):
        class_name = self._safe_class_name(rule)
        parameter_name = self._parameter_label(self._extract_rule_parameter_id(rule))
        value = self._extract_rule_value(rule)
        if value is None:
            value_text = ""
        else:
            value_text = " = {}".format(value)

        evaluator_name = ""
        for property_name in ("Evaluator",):
            try:
                evaluator = getattr(rule, property_name)
                evaluator_name = " [{}]".format(self._safe_class_name(evaluator))
                break
            except Exception:
                continue
        return "{}- {}: {}{}{}".format(indent, class_name, parameter_name, value_text, evaluator_name)

    def _describe_element_filter(self, element_filter, indent=""):
        if element_filter is None:
            return []

        class_name = self._safe_class_name(element_filter)
        lines = ["{}{}".format(indent, class_name)]

        if class_name in ("LogicalAndFilter", "LogicalOrFilter"):
            for method_name in ("GetFilters",):
                try:
                    child_filters = list(getattr(element_filter, method_name)())
                    for child_filter in child_filters:
                        lines.extend(self._describe_element_filter(child_filter, indent + "  "))
                    return lines
                except Exception:
                    continue

        if class_name == "ElementParameterFilter":
            for method_name in ("GetRules",):
                try:
                    rules = list(getattr(element_filter, method_name)())
                    for rule in rules:
                        lines.append(self._describe_rule(rule, indent + "  "))
                    return lines
                except Exception:
                    continue

        lines.append("{}- {}".format(indent + "  ", str(element_filter)))
        return lines

    def _element_filter_signature(self, filter_element):
        parts = []

        for method_name in ("GetElementFilter",):
            try:
                element_filter = getattr(filter_element, method_name)()
                parts.append(str(element_filter))
                break
            except Exception:
                continue

        for method_name in ("GetRules",):
            try:
                rules = list(getattr(filter_element, method_name)())
                rule_parts = [str(rule) for rule in rules]
                parts.append("RULES:" + "|".join(rule_parts))
                break
            except Exception:
                continue

        return " || ".join(parts)

    def _filter_content_signature(self, filter_element):
        return "CATS[{0}] FILTER[{1}]".format(
            self._category_signature(filter_element),
            self._element_filter_signature(filter_element),
        )

    def _build_rows(self):
        filters = list(FilteredElementCollector(doc).OfClass(ParameterFilterElement))
        if not filters:
            return []

        duplicate_groups = {}
        content_groups = {}
        exact_groups = {}
        for filt in filters:
            name_key = self._match_key(self._element_name(filt))
            content_key = self._filter_content_signature(filt)
            exact_key = "{0} || {1}".format(name_key, content_key)
            duplicate_groups.setdefault(name_key, []).append(filt)
            content_groups.setdefault(content_key, []).append(filt)
            exact_groups.setdefault(exact_key, []).append(filt)

        filter_usage = {}
        for filt in filters:
            filter_usage[_element_id_value(filt.Id)] = {"views": [], "templates": []}

        for view in FilteredElementCollector(doc).OfClass(View):
            try:
                if not view.AreGraphicsOverridesAllowed():
                    continue
                filter_ids = list(view.GetFilters())
            except Exception:
                continue

            for filter_id in filter_ids:
                usage = filter_usage.get(_element_id_value(filter_id))
                if usage is None:
                    continue
                view_name = self._element_name(view)
                if getattr(view, "IsTemplate", False):
                    usage["templates"].append(view_name)
                else:
                    usage["views"].append(view_name)

        rows = []
        for filt in filters:
            group_key = self._match_key(self._element_name(filt))
            content_key = self._filter_content_signature(filt)
            exact_key = "{0} || {1}".format(group_key, content_key)

            duplicate_group = ""
            duplicate_type = ""
            if len(exact_groups.get(exact_key, [])) > 1:
                duplicate_group = group_key
                duplicate_type = "Exact Match"
            elif len(content_groups.get(content_key, [])) > 1:
                duplicate_group = group_key
                duplicate_type = "Rule Duplicate"
            elif len(duplicate_groups.get(group_key, [])) > 1:
                duplicate_group = group_key
                duplicate_type = "Name Duplicate"

            usage = filter_usage.get(_element_id_value(filt.Id), {"views": [], "templates": []})
            rows.append(
                FilterRow(
                    _element_id_value(filt.Id),
                    self._element_name(filt),
                    sorted(usage["views"]),
                    sorted(usage["templates"]),
                    duplicate_group,
                    duplicate_type,
                    self._category_summary(filt),
                    content_key,
                )
            )

        return sorted(rows, key=lambda row: (row.FilterName or "").lower())

    def _scope_label(self):
        return self.ScopeComboBox.SelectedItem or _ui_text("All Filters")

    def _matches_scope(self, row):
        scope = self._scope_label()
        if scope == _ui_text("Unused Only"):
            return row.TotalUses == 0
        if scope == _ui_text("Used Only"):
            return row.TotalUses > 0
        if scope == _ui_text("Duplicates Only"):
            return bool(row.DuplicateType)
        return True

    def _matches_search(self, row):
        search_text = (self.SearchTextBox.Text or "").strip().lower()
        if not search_text:
            return True
        haystack = " | ".join([row.FilterName or "", row.Status or "", row.DuplicateGroupDisplay or "", row.DuplicateTypeDisplay or ""]).lower()
        return search_text in haystack

    def _update_counts(self):
        self.TotalText.Text = str(len(self.all_rows))
        self.VisibleText.Text = str(len(self.preview_rows))
        self.UnusedText.Text = str(len([row for row in self.all_rows if row.TotalUses == 0]))
        self.DuplicateText.Text = str(len([row for row in self.all_rows if row.DuplicateType]))

    def _update_visible_rows(self):
        self.preview_rows.Clear()
        for row in self.all_rows:
            if not self._matches_scope(row):
                continue
            if not self._matches_search(row):
                continue
            self.preview_rows.Add(row)

        self._update_counts()
        self.ResultsGrid.Items.Refresh()

    def _commit_grid_edits(self):
        try:
            self.ResultsGrid.CommitEdit(DataGridEditingUnit.Cell, True)
        except Exception:
            pass
        try:
            self.ResultsGrid.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass

    def _sync_row_after_name_edit(self, row):
        if row is None:
            return
        row.FilterName = (row.FilterName or "").strip()
        if row.HasPendingRename:
            row.Include = True
        elif row.TotalUses == 0:
            row.Include = True
        row.Status = row._build_status()

    def _refresh_row_after_edit(self, row):
        self._sync_row_after_name_edit(row)
        try:
            self.ResultsGrid.Items.Refresh()
        except Exception:
            pass

    def _collect_rename_rows(self):
        self._commit_grid_edits()
        changed_rows = []
        for row in self.all_rows:
            self._sync_row_after_name_edit(row)
            if row.HasPendingRename:
                changed_rows.append(row)
        return changed_rows

    def _validate_rename_rows(self, rows):
        errors = []
        final_names = {}
        unchanged_names = {}

        for row in self.all_rows:
            unchanged_names[row.FilterId] = (row.FilterName or "").strip()

        for row in rows:
            new_name = (row.FilterName or "").strip()
            if not new_name:
                errors.append(_ui_text("Empty filter name: {0}").format(row.OriginalName))
                continue
            final_names[row.FilterId] = new_name

        if errors:
            return errors

        used_names = {}
        for row in self.all_rows:
            if row.FilterId in final_names:
                continue
            existing_name = (row.FilterName or "").strip().lower()
            if existing_name:
                used_names.setdefault(existing_name, []).append(row.OriginalName)

        for row in rows:
            new_name = final_names.get(row.FilterId, "")
            if not new_name:
                continue
            key = new_name.lower()
            if key in used_names:
                errors.append(_ui_text("Name already exists: {0}").format(new_name))
                continue
            used_names.setdefault(key, []).append(row.OriginalName)

        return errors

    def _reload_rows(self):
        self.all_rows = self._build_rows()
        self._update_visible_rows()

    def SearchTextBox_TextChanged(self, sender, args):
        self._update_visible_rows()

    def ScopeComboBox_SelectionChanged(self, sender, args):
        self._update_visible_rows()

    def RefreshButton_Click(self, sender, args):
        self._reload_rows()
        if not self.all_rows:
            forms.alert(_ui_text("No filters were found in this project."), title=_ui_text("Manage Filters"), exitscript=False)
        elif not self.preview_rows:
            forms.alert(_ui_text("No filters match the current criteria."), title=_ui_text("Manage Filters"), exitscript=False)

    def CheckVisibleButton_Click(self, sender, args):
        for row in self.preview_rows:
            row.Include = (row.TotalUses == 0) or row.HasPendingRename
        self.ResultsGrid.Items.Refresh()

    def IncludeRowCheckBox_Click(self, sender, args):
        self.ResultsGrid.Items.Refresh()

    def ResultsGrid_CellEditEnding(self, sender, args):
        try:
            row = args.Row.Item
        except Exception:
            row = None

        try:
            text = args.EditingElement.Text
            if row is not None:
                row.FilterName = text
        except Exception:
            pass

        self._pending_edit_row = row

    def ResultsGrid_CurrentCellChanged(self, sender, args):
        row = self._pending_edit_row
        self._pending_edit_row = None
        if row is None:
            return
        try:
            self.Dispatcher.BeginInvoke(Action(lambda: self._refresh_row_after_edit(row)))
        except Exception:
            self._refresh_row_after_edit(row)

    def ResultsGrid_MouseDoubleClick(self, sender, args):
        row = self.ResultsGrid.SelectedItem
        if row is None:
            return

        filter_element = doc.GetElement(ElementId(row.FilterId))
        lines = []
        lines.append("Name: {}".format(row.FilterName))
        lines.append("Status: {}".format(row.Status))
        lines.append("Duplicate Type: {}".format(row.DuplicateType or "-"))
        lines.append("Duplicate Group: {}".format(row.DuplicateGroup or "-"))
        lines.append("Categories: {}".format(row.CategorySummary or "-"))
        lines.append("")
        lines.append("[Definition]")
        try:
            element_filter = filter_element.GetElementFilter()
            definition_lines = self._describe_element_filter(element_filter, "  ")
            if definition_lines:
                lines.extend(definition_lines)
            else:
                lines.append("  No rule details available.")
        except Exception:
            lines.append("  No rule details available.")
        lines.append("")
        if row.UsedViews:
            lines.append("[Views]")
            lines.extend(" - " + item for item in row.UsedViews[:40])
        if row.UsedTemplates:
            if lines:
                lines.append("")
            lines.append("[Templates]")
            lines.extend(" - " + item for item in row.UsedTemplates[:40])
        if not lines:
            lines.append(_ui_text("No views or templates use the selected filter."))

        if row.ContentSignature:
            lines.append("")
            lines.append("[Signature]")
            signature_text = row.ContentSignature
            if len(signature_text) > 2500:
                signature_text = signature_text[:2500] + " ..."
            lines.append(signature_text)

        message = "\n".join(lines)
        if len(message) > 6000:
            message = message[:6000] + "\n\n..."
        forms.alert(message, title=_ui_text("Filter Usage Preview"), exitscript=False)

    def PurgeButton_Click(self, sender, args):
        self._commit_grid_edits()
        selected_rows = [row for row in self.all_rows if row.Include and row.TotalUses == 0]
        if not selected_rows:
            forms.alert(_ui_text("No unused checked filters are ready to purge."), title=_ui_text("Manage Filters"), exitscript=False)
            return

        confirm = forms.alert(_ui_text("Purge {0} unused filters?\n\nThis action cannot be undone.").format(len(selected_rows)), yes=True, no=True, exitscript=False)
        if not confirm:
            return

        purged = 0
        skipped = []
        with revit.Transaction("Purge Unused Filters"):
            for row in selected_rows:
                try:
                    doc.Delete(ElementId(row.FilterId))
                    purged += 1
                except Exception as exc:
                    skipped.append("{} ({})".format(row.FilterName, exc))
                    logger.error("Error purging filter: {} | {}".format(row.FilterName, exc))

        self._reload_rows()
        message = _ui_text("Process finished.\n\nPurged: {0}").format(purged)
        if skipped:
            message += _ui_text("\nSkipped: {0}\n\n{1}").format(len(skipped), "\n".join(skipped[:12]))
        forms.alert(message, title=_ui_text("Manage Filters"), exitscript=False)

    def SaveRenamesButton_Click(self, sender, args):
        rename_rows = self._collect_rename_rows()
        if not rename_rows:
            forms.alert(_ui_text("No edited filter names are ready to save."), title=_ui_text("Manage Filters"), exitscript=False)
            return

        validation_errors = self._validate_rename_rows(rename_rows)
        if validation_errors:
            forms.alert("\n".join(validation_errors[:12]), title=_ui_text("Manage Filters"), exitscript=False)
            return

        confirm = forms.alert(
            _ui_text("Save {0} filter renames?").format(len(rename_rows)),
            yes=True,
            no=True,
            exitscript=False
        )
        if not confirm:
            return

        saved = 0
        skipped = []
        with revit.Transaction("Rename Filters"):
            for row in rename_rows:
                new_name = (row.FilterName or "").strip()
                try:
                    filter_element = doc.GetElement(ElementId(row.FilterId))
                    if filter_element is None:
                        skipped.append(_ui_text("Filter not found: {0}").format(row.OriginalName))
                        continue
                    filter_element.Name = new_name
                    saved += 1
                except Exception as exc:
                    skipped.append("{} ({})".format(row.OriginalName, exc))
                    logger.error("Error renaming filter: {} | {}".format(row.OriginalName, exc))

        self._reload_rows()
        message = _ui_text("Process finished.\n\nSaved renames: {0}").format(saved)
        if skipped:
            message += _ui_text("\nSkipped: {0}\n\n{1}").format(len(skipped), "\n".join(skipped[:12]))
        forms.alert(message, title=_ui_text("Manage Filters"), exitscript=False)

    def CloseButton_Click(self, sender, args):
        self.Close()


ManageFiltersWindow().ShowDialog()

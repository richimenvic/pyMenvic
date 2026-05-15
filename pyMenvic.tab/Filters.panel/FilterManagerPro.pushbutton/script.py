# -*- coding: utf-8 -*-

__title__ = "Filter Manager Pro"
__author__ = "Ricardo J. Mendieta | pyMENVIC"

import os
import sys

try:
    from lib.filters.collectors import collect_parameter_filters
    from lib.filters.compat import element_id_value
    from lib.filters.elements import element_name
    from lib.filters.resources import get_filters_logo_path
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
    from filters.collectors import collect_parameter_filters
    from filters.compat import element_id_value
    from filters.elements import element_name
    from filters.resources import get_filters_logo_path

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
try:
    clr.AddReference("RevitAPIUI")
except Exception:
    pass

from System.Collections.ObjectModel import ObservableCollection
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System.Windows import Visibility
from System.Windows.Controls import DataGridEditingUnit
from System import Int64

from Autodesk.Revit.DB import View, ElementId, Transaction, Category, BuiltInParameter, LabelUtils
try:
    from Autodesk.Revit.DB import WorksetId
except Exception:
    WorksetId = None
try:
    from Autodesk.Revit.UI import RevitCommandId, PostableCommand
except Exception:
    RevitCommandId = None
    PostableCommand = None
from pyrevit import forms, revit, script

doc = revit.doc
XAML_FILE = script.get_bundle_file("filter_manager_pro.xaml")
LOGO_FILE = get_filters_logo_path()
TOOL_VERSION = "MVP 0.3.39"
TOOL_LABEL = "pyMENVIC Filter Manager Pro | {}".format(TOOL_VERSION)


def safe_element_id(value):
    try:
        if isinstance(value, ElementId):
            return value
    except Exception:
        pass
    try:
        return ElementId(Int64(value))
    except Exception:
        return ElementId(value)


class FilterOption(object):
    def __init__(self, element_id, name):
        self.ElementId = element_id
        self.Name = name


class AuditRow(object):
    def __init__(self, filter_id, original_name, name, categories, category_names, vc, tc, duplicate_type, duplicate_group):
        self.FilterId = filter_id
        self.OriginalName = original_name
        self.FilterName = name
        self.Categories = categories
        self.CategoryNames = category_names or []
        self.ViewCount = vc
        self.TemplateCount = tc
        self.TotalCount = vc + tc
        self.Status = "Used" if self.TotalCount > 0 else "Unused"
        self.DuplicateType = duplicate_type or "Not duplicate"
        self.DuplicateGroup = duplicate_group or "-"
        self.Duplicate = self.DuplicateType
        self.Purge = False


class RenameRow(object):
    def __init__(self, filter_id, current, proposed):
        self.FilterId = filter_id
        self.CurrentName = current
        self.ProposedName = proposed
        self.Apply = False
        self.Status = "No change"


class ReplaceRow(object):
    def __init__(self, view_id, view_name, kind, templ, hs, ht, se, sv, te, tv):
        self.ViewId = view_id
        self.ViewName = view_name
        self.ViewKind = kind
        self.IsTemplate = templ
        self.HasSource = hs
        self.HasTarget = ht
        self.SourceEnabled = se
        self.SourceVisible = sv
        self.TargetEnabled = te
        self.TargetVisible = tv
        self.Apply = hs
        self.Status = "Ready" if hs else "No source"


class FilterManagerProWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self._load_header_logo()
        try:
            self.FooterVersionTextBlock.Text = TOOL_LABEL
        except Exception:
            pass
        self.filters = self._collect_filters()
        self._rebuild_maps()
        self.audit_rows = ObservableCollection[object]()
        self.rename_rows = ObservableCollection[object]()
        self.replace_rows = ObservableCollection[object]()
        self.all_audit_rows = []
        self.all_rename_rows = []
        self.all_replace_rows = []
        self.AuditGrid.ItemsSource = self.audit_rows
        self.RenameGrid.ItemsSource = self.rename_rows
        self.ReplaceGrid.ItemsSource = self.replace_rows
        try:
            self.AuditScopeComboBox.SelectedIndex = 0
        except Exception:
            pass
        self.SourceComboBox.ItemsSource = self.filter_names
        self.TargetComboBox.ItemsSource = self.filter_names
        if self.filter_names:
            self.SourceComboBox.SelectedIndex = 0
            self.TargetComboBox.SelectedIndex = min(1, len(self.filter_names) - 1)
        self._load_audit()
        self._load_rename_rows()
        self._set_reports_status("Ready to export from current tab data.")
        self._set_audit_status("Audit is safe. Edit only filter names in the Filter column, then apply changes.")
        self._set_rename_status("Configure rename options and click Preview.")
        self._set_replace_status("Select Source and Target, then Preview Usage.")
        self._refresh_active_tab_summary()
        self._select_first_visible_audit_row()

    def _rebuild_maps(self):
        self.filter_name_to_option = {f.Name: f for f in self.filters}
        self.filter_id_to_option = {element_id_value(f.ElementId): f for f in self.filters}
        self.filter_names = sorted(self.filter_name_to_option.keys(), key=lambda x: x.lower())

    def _collect_filters(self):
        return [FilterOption(f.Id, element_name(f)) for f in collect_parameter_filters(doc, key_selector=lambda i: element_name(i).lower())]

    def _views(self):
        from Autodesk.Revit.DB import FilteredElementCollector
        out = []
        for v in FilteredElementCollector(doc).OfClass(View):
            try:
                v.GetFilters()
                out.append(v)
            except Exception:
                pass
        return out

    def _set_text(self, name, v):
        try:
            getattr(self, name).Text = str(v)
        except Exception:
            pass

    def _set_header_cards(self, cards):
        cards = cards[:4]
        for i in range(4):
            d = cards[i] if i < len(cards) else None
            self._set_text("HeaderCardLabel{}".format(i + 1), d[0] if d else "")
            self._set_text("HeaderCardValue{}".format(i + 1), d[1] if d else "")
            try:
                getattr(self, "HeaderCardBorder{}".format(i + 1)).Visibility = Visibility.Visible if d else Visibility.Collapsed
            except Exception:
                pass

    def _card(self, l, v):
        return (l, str(v))

    def _duplicate_rows(self, rows):
        return [r for r in rows if r.DuplicateType != "Not duplicate"]

    def _duplicate_group_count(self, rows):
        groups = set()
        for r in self._duplicate_rows(rows):
            if r.DuplicateGroup and r.DuplicateGroup != "-":
                groups.add(r.DuplicateGroup)
        return len(groups)

    def _refresh_active_tab_summary(self):
        h = "Audit"
        try:
            h = str(self.MainTabControl.SelectedItem.Header)
        except Exception:
            pass
        if "Audit" in h:
            used = len([r for r in self.all_audit_rows if r.TotalCount > 0])
            self._set_header_cards([
                self._card("FILTERS", len(self.all_audit_rows)),
                self._card("VISIBLE", len(self.audit_rows)),
                self._card("UNUSED", len(self.all_audit_rows) - used),
                self._card("DUP. SETS", self._duplicate_group_count(self.all_audit_rows))
            ])
        elif "Rename" in h:
            ready = len([r for r in self.rename_rows if r.Apply])
            self._set_header_cards([self._card("ROWS", len(self.rename_rows)), self._card("READY", ready)])
        elif "Replace" in h:
            self._set_header_cards([self._card("PREVIEW", len(self.replace_rows)), self._card("APPLY", len([r for r in self.replace_rows if r.Apply]))])
        else:
            self._set_header_cards([self._card("REPORTS", "CSV")])

    def _load_header_logo(self):
        s = None
        try:
            s = FileStream(LOGO_FILE, FileMode.Open, FileAccess.Read)
            i = BitmapImage()
            i.BeginInit()
            i.StreamSource = s
            i.CacheOption = BitmapCacheOption.OnLoad
            i.EndInit()
            i.Freeze()
            self.HeaderLogoImage.Source = i
        except Exception:
            pass
        finally:
            if s:
                try:
                    s.Close()
                except Exception:
                    pass

    def _safe_class_name(self, obj):
        try:
            return obj.GetType().Name
        except Exception:
            try:
                return obj.__class__.__name__
            except Exception:
                return str(type(obj))

    def _category_ids(self, filter_el):
        try:
            return list(filter_el.GetCategories())
        except Exception:
            try:
                raw = getattr(filter_el, "Categories", None)
                if raw:
                    return list(raw)
            except Exception:
                pass
        return []

    def _category_name_from_id(self, cid):
        try:
            cat = Category.GetCategory(doc, cid)
            if cat and cat.Name:
                return cat.Name
        except Exception:
            pass
        try:
            cat = doc.Settings.Categories.get_Item(cid)
            if cat and cat.Name:
                return cat.Name
        except Exception:
            pass
        try:
            cid_value = element_id_value(cid)
            for cat in doc.Settings.Categories:
                try:
                    if element_id_value(cat.Id) == cid_value:
                        return cat.Name
                except Exception:
                    pass
        except Exception:
            pass
        try:
            return "CategoryId {}".format(element_id_value(cid))
        except Exception:
            return str(cid)

    def _get_category_names(self, filter_el):
        names = []
        for cid in self._category_ids(filter_el):
            name = self._category_name_from_id(cid)
            if name:
                names.append(name)
        return sorted(set(names))

    def _format_category_summary(self, category_names):
        names = list(category_names or [])
        if not names:
            return "N/A"
        if len(names) <= 3:
            return ", ".join(names)
        return "{}, +{}".format(", ".join(names[:3]), len(names) - 3)

    def _get_categories(self, filter_el):
        return self._format_category_summary(self._get_category_names(filter_el))

    def _category_signature(self, filter_el):
        values = []
        for cid in self._category_ids(filter_el):
            try:
                values.append(str(element_id_value(cid)))
            except Exception:
                values.append(str(cid))
        return "|".join(sorted(values))

    def _param_name(self, parameter_id):
        if parameter_id is None:
            return "<unknown parameter>"
        try:
            el = doc.GetElement(parameter_id)
            if el:
                return element_name(el)
        except Exception:
            pass
        try:
            param_value = element_id_value(parameter_id)
            if param_value == -1:
                return "Category"
            try:
                label = LabelUtils.GetLabelFor(BuiltInParameter(param_value))
                if label:
                    return label
            except Exception:
                pass
            return "Built-in parameter {}".format(param_value)
        except Exception:
            return str(parameter_id)

    def _display_value(self, value):
        if value is None:
            return "<value not readable>"
        try:
            el = doc.GetElement(value)
            if el:
                return element_name(el)
        except Exception:
            pass
        try:
            raw_value = element_id_value(value)
            if raw_value == -1:
                return "<value not readable>"
            return str(raw_value)
        except Exception:
            return str(value)

    def _display_workset_value(self, value):
        try:
            if isinstance(value, str):
                return value
        except Exception:
            pass
        try:
            raw_value = element_id_value(value)
        except Exception:
            raw_value = None
        try:
            if raw_value is None:
                return str(value)
            workset_table = doc.GetWorksetTable()
            if workset_table is None:
                return str(raw_value)
            workset_id = None
            try:
                if WorksetId is not None:
                    workset_id = WorksetId(Int64(raw_value))
            except Exception:
                workset_id = None
            if workset_id is None:
                return str(raw_value)
            workset = workset_table.GetWorkset(workset_id)
            if workset and getattr(workset, "Name", None):
                return workset.Name
            return str(raw_value)
        except Exception:
            try:
                return str(raw_value)
            except Exception:
                return str(value)

    def _inner_rule(self, rule):
        if self._safe_class_name(rule) != "FilterInverseRule":
            return None
        for method_name in ("GetInnerRule", "GetRule"):
            try:
                inner = getattr(rule, method_name)()
                if inner:
                    return inner
            except Exception:
                pass
        for property_name in ("InnerRule", "Rule"):
            try:
                inner = getattr(rule, property_name)
                if inner:
                    return inner
            except Exception:
                pass
        return None

    def _rule_for_value_access(self, rule):
        inner = self._inner_rule(rule)
        return inner if inner else rule

    def _extract_rule_parameter_id(self, rule):
        access_rule = self._rule_for_value_access(rule)
        for method_name in ("GetRuleParameter", "GetParameterId", "GetParameter"):
            try:
                return getattr(access_rule, method_name)()
            except Exception:
                pass
        for property_name in ("RuleParameter", "ParameterId", "Parameter"):
            try:
                return getattr(access_rule, property_name)
            except Exception:
                pass
        return None

    def _extract_rule_value(self, rule):
        access_rule = self._rule_for_value_access(rule)
        for method_name in ("GetRuleString", "GetRuleValue", "GetStringValue", "GetValue", "GetIntegerValue", "GetDoubleValue", "GetElementIdValue"):
            try:
                value = getattr(access_rule, method_name)()
                if value is not None:
                    return value
            except Exception:
                pass
        for property_name in ("RuleString", "RuleValue", "StringValue", "Value", "IntegerValue", "DoubleValue", "ElementIdValue"):
            try:
                value = getattr(access_rule, property_name)
                if value is not None:
                    return value
            except Exception:
                pass
        return None

    def _extract_rule_evaluator_name(self, rule):
        access_rule = self._rule_for_value_access(rule)
        for method_name in ("GetEvaluator",):
            try:
                evaluator = getattr(access_rule, method_name)()
                if evaluator:
                    return self._safe_class_name(evaluator)
            except Exception:
                pass
        for property_name in ("Evaluator",):
            try:
                evaluator = getattr(access_rule, property_name)
                if evaluator:
                    return self._safe_class_name(evaluator)
            except Exception:
                pass
        return ""

    def _friendly_operator_name(self, evaluator_name, rule_name):
        raw = evaluator_name or rule_name or ""
        lookup = {
            "FilterStringEquals": "equals",
            "FilterStringContains": "contains",
            "FilterStringBeginsWith": "begins with",
            "FilterStringEndsWith": "ends with",
            "FilterNumericEquals": "equals",
            "FilterNumericGreater": "greater than",
            "FilterNumericGreaterOrEqual": "greater than or equal",
            "FilterNumericLess": "less than",
            "FilterNumericLessOrEqual": "less than or equal",
            "FilterElementIdEquals": "equals",
            "FilterElementIdNotEquals": "does not equal",
            "FilterStringRule": "equals",
            "FilterIntegerRule": "equals",
            "FilterDoubleRule": "equals",
            "FilterElementIdRule": "equals",
            "FilterInverseRule": "not"
        }
        if raw in lookup:
            return lookup[raw]
        cleaned = raw.replace("Filter", "").replace("Evaluator", "").replace("Rule", "")
        return cleaned.strip().lower() or "operator not readable"

    def _invert_operator(self, operator_name):
        lookup = {
            "equals": "does not equal",
            "contains": "does not contain",
            "begins with": "does not begin with",
            "ends with": "does not end with",
            "greater than": "is not greater than",
            "greater than or equal": "is less than",
            "less than": "is not less than",
            "less than or equal": "is greater than"
        }
        return lookup.get(operator_name, "not " + operator_name)

    def _extract_rule_operator(self, rule):
        base_operator = self._friendly_operator_name(self._extract_rule_evaluator_name(rule), self._safe_class_name(self._rule_for_value_access(rule)))
        if self._safe_class_name(rule) == "FilterInverseRule":
            return self._invert_operator(base_operator)
        return base_operator

    def _rule_signature(self, rule):
        parameter_id = self._extract_rule_parameter_id(rule)
        try:
            parameter_value = element_id_value(parameter_id)
        except Exception:
            parameter_value = str(parameter_id)
        value = self._extract_rule_value(rule)
        try:
            value = element_id_value(value)
        except Exception:
            pass
        evaluator_name = self._extract_rule_evaluator_name(rule)
        inverse = "NOT" if self._safe_class_name(rule) == "FilterInverseRule" else ""
        return "{}|{}|{}|{}|{}".format(self._safe_class_name(self._rule_for_value_access(rule)), parameter_value, value, evaluator_name, inverse)

    def _rule_detail_text(self, rule):
        raw_value = self._extract_rule_value(rule)
        parameter_name = self._param_name(self._extract_rule_parameter_id(rule))
        operator_name = self._extract_rule_operator(rule)
        if (parameter_name or "").lower() == "workset":
            value = self._display_workset_value(raw_value)
        else:
            value = self._display_value(raw_value)
        if value == "<value not readable>":
            if (parameter_name or "").lower() == "category":
                return "Category: category not readable"
            return "{}: value not readable".format(parameter_name)
        return "{} {} {}".format(parameter_name, operator_name, value)

    def _rule_detail_line(self, rule):
        return "- " + self._rule_detail_text(rule)

    def _is_unreadable_category_rule_text(self, text):
        clean = (text or "").strip()
        return clean in (
            "Category: value not readable",
            "Category: category not readable",
            "- Category: value not readable",
            "- Category: category not readable"
        )

    def _category_filter_names(self, element_filter):
        if element_filter is None:
            return []
        class_name = self._safe_class_name(element_filter)
        ids = []
        if class_name == "ElementCategoryFilter":
            for method_name in ("GetCategoryId",):
                try:
                    cid = getattr(element_filter, method_name)()
                    if cid:
                        ids.append(cid)
                except Exception:
                    pass
            for property_name in ("CategoryId",):
                try:
                    cid = getattr(element_filter, property_name)
                    if cid:
                        ids.append(cid)
                except Exception:
                    pass
        elif class_name == "ElementMulticategoryFilter":
            for method_name in ("GetCategoryIds", "GetCategories"):
                try:
                    ids.extend(list(getattr(element_filter, method_name)()))
                except Exception:
                    pass
        names = []
        for cid in ids:
            name = self._category_name_from_id(cid)
            if name:
                names.append(name)
        return sorted(set(names))

    def _element_filter_signature(self, element_filter):
        if element_filter is None:
            return ""
        class_name = self._safe_class_name(element_filter)
        if class_name in ("LogicalAndFilter", "LogicalOrFilter"):
            parts = []
            try:
                for child_filter in list(element_filter.GetFilters()):
                    parts.append(self._element_filter_signature(child_filter))
            except Exception:
                pass
            return "{}({})".format(class_name, ";".join(sorted(parts)))
        if class_name == "ElementParameterFilter":
            rule_parts = []
            try:
                for rule in list(element_filter.GetRules()):
                    rule_parts.append(self._rule_signature(rule))
            except Exception:
                pass
            return "{}({})".format(class_name, ";".join(sorted(rule_parts)))
        category_names = self._category_filter_names(element_filter)
        if category_names:
            return "{}({})".format(class_name, ";".join(category_names))
        return "{}:{}".format(class_name, str(element_filter))

    def _child_filters(self, element_filter):
        try:
            return list(element_filter.GetFilters())
        except Exception:
            return []

    def _parameter_rule_texts_from_filter(self, element_filter):
        texts = []
        if self._safe_class_name(element_filter) != "ElementParameterFilter":
            return texts
        try:
            for rule in list(element_filter.GetRules()):
                rule_text = self._rule_detail_text(rule)
                if self._is_unreadable_category_rule_text(rule_text):
                    continue
                texts.append(rule_text)
        except Exception:
            pass
        return texts

    def _element_filter_detail_lines(self, element_filter, indent="", group_label=None):
        if element_filter is None:
            return []
        class_name = self._safe_class_name(element_filter)
        lines = []
        if class_name in ("LogicalAndFilter", "LogicalOrFilter"):
            child_filters = self._child_filters(element_filter)
            if class_name == "LogicalAndFilter":
                category_names = []
                rule_texts = []
                for child_filter in child_filters:
                    category_names.extend(self._category_filter_names(child_filter))
                    rule_texts.extend(self._parameter_rule_texts_from_filter(child_filter))
                category_names = sorted(set(category_names))
                filtered_rule_texts = []
                for rule_text in rule_texts:
                    if self._is_unreadable_category_rule_text(rule_text):
                        continue
                    filtered_rule_texts.append(rule_text)
                if category_names and filtered_rule_texts:
                    for category_name in category_names:
                        for rule_text in filtered_rule_texts:
                            lines.append("{}- {} | {}".format(indent, category_name, rule_text))
                    return lines
                if filtered_rule_texts:
                    for rule_text in filtered_rule_texts:
                        lines.append("{}- {}".format(indent, rule_text))
                    return lines
            logic_label = "All rules must be true:" if class_name == "LogicalAndFilter" else "Any rule may be true:"
            lines.append("{}{}".format(indent, group_label if group_label else logic_label))
            if child_filters:
                logical_type_counts = {}
                for child_filter in child_filters:
                    child_class = self._safe_class_name(child_filter)
                    if child_class in ("LogicalAndFilter", "LogicalOrFilter"):
                        logical_type_counts[child_class] = logical_type_counts.get(child_class, 0) + 1
                logical_type_index = {}
                for child_filter in child_filters:
                    child_class = self._safe_class_name(child_filter)
                    child_group_label = None
                    if child_class in ("LogicalAndFilter", "LogicalOrFilter") and logical_type_counts.get(child_class, 0) > 1:
                        logical_type_index[child_class] = logical_type_index.get(child_class, 0) + 1
                        child_logic_label = "All rules must be true:" if child_class == "LogicalAndFilter" else "Any rule may be true:"
                        child_group_label = "Rule Set {} - {}".format(logical_type_index[child_class], child_logic_label)
                    lines.extend(self._element_filter_detail_lines(child_filter, indent + "  ", child_group_label))
            else:
                lines.append("{}  <unable to read child filters>".format(indent))
            return lines
        if class_name == "ElementParameterFilter":
            try:
                rules = list(element_filter.GetRules())
                if not rules:
                    lines.append("{}<no rules>".format(indent))
                for rule in rules:
                    rule_line = self._rule_detail_line(rule)
                    if self._is_unreadable_category_rule_text(rule_line):
                        continue
                    lines.append(indent + rule_line)
            except Exception:
                lines.append("{}<unable to read parameter rules>".format(indent))
            return lines
        category_names = self._category_filter_names(element_filter)
        if category_names:
            lines.append("{}- Category: {}".format(indent, ", ".join(category_names)))
            return lines
        lines.append("{}{}".format(indent, class_name))
        return lines

    def _filter_rule_lines(self, filter_el):
        if filter_el is None:
            return ["<filter not found>"]
        try:
            lines = self._element_filter_detail_lines(filter_el.GetElementFilter())
            if lines:
                return lines
        except Exception:
            pass
        try:
            rules = list(filter_el.GetRules())
            if rules:
                lines = []
                for rule in rules:
                    rule_line = self._rule_detail_line(rule)
                    if self._is_unreadable_category_rule_text(rule_line):
                        continue
                    lines.append(rule_line)
                if lines:
                    return lines
        except Exception:
            pass
        return ["Category-only filter. No parameter rules."]

    def _number_repeated_rule_sets(self, lines):
        target_labels = ("Any rule may be true:", "All rules must be true:")
        totals = {}
        for line in lines:
            clean = line.strip()
            if clean in target_labels:
                totals[clean] = totals.get(clean, 0) + 1
        if all(count <= 1 for count in totals.values()):
            return lines
        indexes = {}
        out = []
        for line in lines:
            clean = line.strip()
            prefix = line[:len(line) - len(line.lstrip())]
            if clean in target_labels and totals.get(clean, 0) > 1:
                indexes[clean] = indexes.get(clean, 0) + 1
                out.append("{}Rule Set {} - {}".format(prefix, indexes[clean], clean))
            else:
                out.append(line)
        return out

    def _filter_content_signature(self, filter_el):
        if filter_el is None:
            return ""
        cat_sig = self._category_signature(filter_el)
        filter_sig = ""
        try:
            filter_sig = self._element_filter_signature(filter_el.GetElementFilter())
        except Exception:
            pass
        if not filter_sig:
            rule_parts = []
            try:
                for rule in list(filter_el.GetRules()):
                    rule_parts.append(self._rule_signature(rule))
            except Exception:
                pass
            if rule_parts:
                filter_sig = "RULES({})".format(";".join(sorted(rule_parts)))
        if not cat_sig or not filter_sig:
            return ""
        return "CATS[{}] FILTER[{}]".format(cat_sig, filter_sig)

    def _name_key(self, value):
        return " ".join("".join([c if c.isalnum() else " " for c in (value or "").upper()]).split())

    def _group_label(self, label_map, key):
        if key not in label_map:
            label_map[key] = "Duplicate Set {:02d}".format(len(label_map) + 1)
        return label_map[key]

    def _load_audit(self):
        views = self._views()
        usage_by_filter_id = {}
        for view in views:
            try:
                filter_ids = list(view.GetFilters())
            except Exception:
                continue
            is_template = False
            try:
                is_template = bool(view.IsTemplate)
            except Exception:
                is_template = False
            for filter_id in filter_ids:
                fid = element_id_value(filter_id)
                if fid not in usage_by_filter_id:
                    usage_by_filter_id[fid] = [0, 0]
                if is_template:
                    usage_by_filter_id[fid][1] += 1
                else:
                    usage_by_filter_id[fid][0] += 1
        rows = []
        content_groups = {}
        name_groups = {}
        for f in self.filters:
            fid = element_id_value(f.ElementId)
            filter_el = doc.GetElement(f.ElementId)
            category_names = self._get_category_names(filter_el)
            cats = self._format_category_summary(category_names)
            sig = self._filter_content_signature(filter_el)
            counts = usage_by_filter_id.get(fid, [0, 0])
            vc = counts[0]
            tc = counts[1]
            name_key = self._name_key(f.Name)
            rows.append((fid, f.Name, cats, category_names, vc, tc, sig, name_key))
            if sig:
                content_groups.setdefault(sig, []).append(fid)
            if name_key:
                name_groups.setdefault(name_key, []).append(fid)
        self.all_audit_rows = []
        duplicate_labels = {}
        for row in rows:
            dup_type = "Not duplicate"
            dup_group = "-"
            if row[6] and len(content_groups.get(row[6], [])) > 1:
                dup_type = "Exact Definition"
                dup_group = self._group_label(duplicate_labels, "DEF:{}".format(row[6]))
            elif row[7] and len(name_groups.get(row[7], [])) > 1:
                dup_type = "Similar Name"
                dup_group = self._group_label(duplicate_labels, "NAME:{}".format(row[7]))
            self.all_audit_rows.append(AuditRow(row[0], row[1], row[1], row[2], row[3], row[4], row[5], dup_type, dup_group))
        self._filter_audit_rows()
        self._update_audit_apply_state()
        self._update_audit_purge_ui()
        self._refresh_active_tab_summary()

    def _audit_scope(self):
        try:
            selected = self.AuditScopeComboBox.SelectedItem
            return str(selected.Content) if hasattr(selected, "Content") else str(selected)
        except Exception:
            return "All Filters"

    def _is_unused_scope(self):
        return self._audit_scope() == "Unused Only"

    def _audit_row_matches_scope(self, row):
        scope = self._audit_scope()
        if scope == "Unused Only":
            return row.TotalCount == 0
        if scope == "Used Only":
            return row.TotalCount > 0
        if scope == "Duplicates Only":
            return row.DuplicateType != "Not duplicate"
        return True

    def _filter_audit_rows(self):
        term = ""
        try:
            term = (self.AuditSearchTextBox.Text or "").strip().lower()
        except Exception:
            pass
        if not self._is_unused_scope():
            for r in self.all_audit_rows:
                r.Purge = False
        self.audit_rows.Clear()
        for r in self.all_audit_rows:
            if not self._audit_row_matches_scope(r):
                continue
            if term:
                haystack = " | ".join([r.FilterName or "", r.Categories or "", " ".join(r.CategoryNames), r.Status or "", r.DuplicateType or "", r.DuplicateGroup or ""]).lower()
                if term not in haystack:
                    continue
            self.audit_rows.Add(r)
        if self._is_unused_scope():
            msg = "Visible: {} unused of {} filters. Check Purge only for unused filters you want to delete.".format(len(self.audit_rows), len(self.all_audit_rows))
        else:
            msg = "Visible: {} of {} filters. Edit names in the Filter column only. Apply activates after a name change.".format(len(self.audit_rows), len(self.all_audit_rows))
        self._set_audit_status(msg)
        self._update_audit_purge_ui()
        self._refresh_active_tab_summary()
        self._select_first_visible_audit_row()

    def _select_first_visible_audit_row(self):
        try:
            if len(self.audit_rows) > 0:
                self.AuditGrid.SelectedIndex = 0
                self.AuditGrid.ScrollIntoView(self.audit_rows[0])
                self._update_audit_details()
            else:
                self.AuditGrid.SelectedIndex = -1
                self._set_audit_details_columns("Select a filter row.", "-", "-")
        except Exception:
            pass

    def _update_audit_purge_ui(self):
        visible = Visibility.Visible if self._is_unused_scope() else Visibility.Collapsed
        try:
            self.PurgeSelectedUnusedButton.Visibility = visible
        except Exception:
            pass
        try:
            self.AuditGrid.Columns[0].Visibility = visible
        except Exception:
            pass

    def _commit_audit_edits(self):
        try:
            self.AuditGrid.CommitEdit(DataGridEditingUnit.Cell, True)
        except Exception:
            pass
        try:
            self.AuditGrid.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass

    def _get_changed_audit_rows(self):
        changed = []
        for r in self.all_audit_rows:
            if (r.FilterName or "").strip() != (r.OriginalName or "").strip():
                changed.append(r)
        return changed

    def _update_audit_apply_state(self):
        try:
            self.ApplyAuditChangesButton.IsEnabled = len(self._get_changed_audit_rows()) > 0
        except Exception:
            pass

    def AuditGrid_CellEditEnding(self, sender, args):
        try:
            header = str(args.Column.Header)
            row = args.Row.Item
            if header == "Filter":
                row.FilterName = args.EditingElement.Text
            elif header == "Purge":
                row.Purge = bool(args.EditingElement.IsChecked) if row.TotalCount == 0 else False
        except Exception:
            pass
        self._update_audit_apply_state()

    def AuditGrid_SelectionChanged(self, sender, args):
        self._update_audit_details()

    def _update_audit_details(self):
        row = None
        try:
            row = self.AuditGrid.SelectedItem
        except Exception:
            pass
        if not row:
            self._set_audit_details_columns("Select a filter row.", "-", "-")
            return
        filter_el = doc.GetElement(safe_element_id(row.FilterId))
        filter_lines = [
            "Name: {}".format(row.FilterName),
            "Usage: {} Views | {} Templates | {} Total".format(row.ViewCount, row.TemplateCount, row.TotalCount),
            "Status: {}".format(row.Status),
            "",
            "CATEGORIES"
        ]
        if row.CategoryNames:
            for category_name in row.CategoryNames:
                filter_lines.append("- {}".format(category_name))
        else:
            filter_lines.append("- N/A")
        duplicate_lines = [
            "Type: {}".format(row.DuplicateType),
            "Set: {}".format(row.DuplicateGroup)
        ]
        same_set = [r.FilterName for r in self.all_audit_rows if r.DuplicateGroup == row.DuplicateGroup and r.DuplicateGroup != "-" and r.FilterId != row.FilterId]
        if same_set:
            duplicate_lines.append("Same Duplicate Set:")
            for name in sorted(same_set):
                duplicate_lines.append("- {}".format(name))
        rule_lines = self._number_repeated_rule_sets(self._filter_rule_lines(filter_el))
        if not rule_lines:
            rule_lines = ["Category-only filter. No parameter rules."]
        self._set_audit_details_columns("\n".join(filter_lines), "\n".join(duplicate_lines), "\n".join(rule_lines))

    def ApplyAuditChangesButton_Click(self, s, a):
        self._commit_audit_edits()
        changed = self._get_changed_audit_rows()
        if not changed:
            self._set_audit_status("No filter name changes to apply.")
            self._update_audit_apply_state()
            return
        names = [(r.FilterName or "").strip() for r in changed]
        if "" in names:
            self._set_audit_status("Apply failed: filter names cannot be empty.")
            self._update_audit_apply_state()
            return
        lowered = [n.lower() for n in names]
        if len(set(lowered)) != len(lowered):
            self._set_audit_status("Apply failed: duplicate edited filter names.")
            self._update_audit_apply_state()
            return
        changed_ids = set([r.FilterId for r in changed])
        existing = set([f.Name.lower() for f in self.filters if element_id_value(f.ElementId) not in changed_ids])
        for n in names:
            if n.lower() in existing:
                self._set_audit_status("Apply failed: name already exists: {}".format(n))
                self._update_audit_apply_state()
                return
        ok = 0
        fail = 0
        tx = Transaction(doc, "Filter Manager Pro - Apply Audit Names")
        tx.Start()
        try:
            for r in changed:
                try:
                    doc.GetElement(safe_element_id(r.FilterId)).Name = (r.FilterName or "").strip()
                    ok += 1
                except Exception:
                    fail += 1
            tx.Commit()
        except Exception:
            try:
                tx.RollBack()
            except Exception:
                pass
            fail = len(changed)
        self.filters = self._collect_filters()
        self._rebuild_maps()
        self.SourceComboBox.ItemsSource = self.filter_names
        self.TargetComboBox.ItemsSource = self.filter_names
        self._load_audit()
        self._load_rename_rows()
        self._set_audit_status("Apply complete. Renamed: {} | Failed: {}".format(ok, fail))

    def PurgeSelectedUnusedButton_Click(self, s, a):
        self._commit_audit_edits()
        selected = [r for r in self.all_audit_rows if r.Purge and r.TotalCount == 0]
        if not selected:
            self._set_audit_status("No unused filters selected to purge.")
            return
        names = sorted([r.FilterName for r in selected])
        preview = "\n".join(names[:10])
        if len(names) > 10:
            preview += "\n...and {} more".format(len(names) - 10)
        confirmed = forms.alert(
            "You are about to delete {} unused filter(s). This cannot be undone.\n\n{}\n\nContinue?".format(len(selected), preview),
            title="Purge Selected Unused Filters",
            yes=True,
            no=True,
            exitscript=False
        )
        if not confirmed:
            self._set_audit_status("Purge cancelled.")
            return
        ok = 0
        fail = 0
        skipped = 0
        tx = Transaction(doc, "Filter Manager Pro - Purge Selected Unused Filters")
        tx.Start()
        try:
            for r in selected:
                if r.TotalCount != 0:
                    skipped += 1
                    continue
                try:
                    doc.Delete(safe_element_id(r.FilterId))
                    ok += 1
                except Exception:
                    fail += 1
            tx.Commit()
        except Exception:
            try:
                tx.RollBack()
            except Exception:
                pass
            fail = len(selected)
        self.filters = self._collect_filters()
        self._rebuild_maps()
        self.SourceComboBox.ItemsSource = self.filter_names
        self.TargetComboBox.ItemsSource = self.filter_names
        self._load_audit()
        self._load_rename_rows()
        self._set_audit_status("Purge complete. Deleted: {} | Skipped: {} | Failed: {}".format(ok, skipped, fail))

    def _load_rename_rows(self):
        self.all_rename_rows = [RenameRow(element_id_value(f.ElementId), f.Name, f.Name) for f in self.filters]
        self._filter_rename_rows()

    def _filter_rename_rows(self):
        term = (self.RenameSearchTextBox.Text or "").strip().lower()
        self.rename_rows.Clear()
        for r in self.all_rename_rows:
            if term and term not in r.CurrentName.lower() and term not in r.ProposedName.lower():
                continue
            self.rename_rows.Add(r)
        self._refresh_active_tab_summary()

    def PreviewRenameButton_Click(self, s, a):
        f = (self.FindTextBox.Text or "")
        rep = (self.ReplaceTextBox.Text or "")
        pre = (self.RenamePrefixTextBox.Text or "")
        suf = (self.RenameSuffixTextBox.Text or "")
        up = self.UppercaseCheckBox.IsChecked
        for r in self.all_rename_rows:
            p = r.CurrentName
            if f:
                p = p.replace(f, rep)
            p = "{}{}{}".format(pre, p, suf)
            if up:
                p = p.upper()
            r.ProposedName = p
            r.Apply = (r.CurrentName != r.ProposedName)
            r.Status = "Ready" if r.Apply else "No change"
        self._filter_rename_rows()
        self._set_rename_status("Preview ready.")

    def ResetRenameRowButton_Click(self, s, a):
        r = self.RenameGrid.SelectedItem
        if r:
            r.ProposedName = r.CurrentName
            r.Apply = False
            r.Status = "Reset"
            self.RenameGrid.Items.Refresh()
            self._refresh_active_tab_summary()

    def AuditSearchTextBox_TextChanged(self, s, a):
        self._filter_audit_rows()

    def AuditScopeComboBox_SelectionChanged(self, s, a):
        self._filter_audit_rows()

    def RenameSearchTextBox_TextChanged(self, s, a):
        self._filter_rename_rows()

    def ReplaceSearchTextBox_TextChanged(self, s, a):
        self._filter_replace_rows()

    def RefreshAuditButton_Click(self, s, a):
        self._load_audit()

    def OpenRevitFiltersButton_Click(self, s, a):
        try:
            if RevitCommandId is None or PostableCommand is None:
                raise Exception("Revit UI command API is not available.")
            cmd_id = RevitCommandId.LookupPostableCommandId(PostableCommand.Filters)
            if not cmd_id:
                raise Exception("Revit Filters command id is not available.")
            uiapp = revit.uidoc.Application
            try:
                self.Close()
            except Exception:
                pass
            uiapp.PostCommand(cmd_id)
        except Exception:
            try:
                self._set_audit_status("Could not open Revit Filters dialog. Use Manage > Filters.")
            except Exception:
                pass
            forms.alert(
                "Could not open Revit Filters dialog from this context.\n\nUse Manage > Filters.",
                title="Open Revit Filters",
                exitscript=False
            )

    def MainTabControl_SelectionChanged(self, s, a):
        self._refresh_active_tab_summary()

    def ApplyRenameButton_Click(self, s, a):
        rows = [r for r in self.all_rename_rows if r.Apply]
        if not rows:
            self._set_rename_status("Nothing selected to rename.")
            return
        names = [(r.ProposedName or "").strip() for r in rows]
        if "" in names:
            self._set_rename_status("Validation failed: empty proposed names.")
            return
        if len(set([n.lower() for n in names])) != len(names):
            self._set_rename_status("Validation failed: duplicate proposed names.")
            return
        existing = set([f.Name.lower() for f in self.filters if element_id_value(f.ElementId) not in [r.FilterId for r in rows]])
        for n in names:
            if n.lower() in existing:
                self._set_rename_status("Validation failed: conflicts with existing filters.")
                return
        ok = 0
        fail = 0
        tx = Transaction(doc, "Filter Manager Pro - Rename")
        tx.Start()
        try:
            for r in rows:
                try:
                    doc.GetElement(safe_element_id(r.FilterId)).Name = r.ProposedName
                    ok += 1
                except Exception:
                    fail += 1
            tx.Commit()
        except Exception:
            try:
                tx.RollBack()
            except Exception:
                pass
            fail = len(rows)
        self.filters = self._collect_filters()
        self._rebuild_maps()
        self.SourceComboBox.ItemsSource = self.filter_names
        self.TargetComboBox.ItemsSource = self.filter_names
        self._load_audit()
        self._load_rename_rows()
        self._set_rename_status("Apply Rename complete. Renamed: {} | Failed: {}".format(ok, fail))

    def PreviewReplaceButton_Click(self, s, a):
        self.replace_rows.Clear()
        self.all_replace_rows = []
        src = self.filter_name_to_option.get(self.SourceComboBox.SelectedItem)
        tgt = self.filter_name_to_option.get(self.TargetComboBox.SelectedItem)
        if not src or not tgt or element_id_value(src.ElementId) == element_id_value(tgt.ElementId):
            self._set_replace_status("Select different source and target filters.")
            return
        svid = element_id_value(src.ElementId)
        tvid = element_id_value(tgt.ElementId)
        inc_views = self.IncludeViewsCheckBox.IsChecked
        inc_t = self.IncludeTemplatesCheckBox.IsChecked
        for v in self._views():
            if (v.IsTemplate and not inc_t) or ((not v.IsTemplate) and not inc_views):
                continue
            try:
                ids = list(v.GetFilters())
                vals = [element_id_value(x) for x in ids]
            except Exception:
                continue
            hs = svid in vals
            ht = tvid in vals
            if not hs and not ht:
                continue
            ge = getattr(v, "GetIsFilterEnabled", None)
            se = None
            te = None
            if callable(ge):
                try:
                    se = ge(src.ElementId)
                except Exception:
                    pass
                try:
                    te = ge(tgt.ElementId)
                except Exception:
                    pass
            sv = None
            tv = None
            try:
                sv = v.GetFilterVisibility(src.ElementId)
            except Exception:
                pass
            try:
                tv = v.GetFilterVisibility(tgt.ElementId)
            except Exception:
                pass
            row = ReplaceRow(element_id_value(v.Id), element_name(v), str(v.ViewType), v.IsTemplate, hs, ht, se, sv, te, tv)
            self.all_replace_rows.append(row)
        self._filter_replace_rows()
        self._set_replace_status("Preview ready: {} rows.".format(len(self.all_replace_rows)))

    def _filter_replace_rows(self):
        term = (self.ReplaceSearchTextBox.Text or "").strip().lower()
        self.replace_rows.Clear()
        for r in self.all_replace_rows:
            if term and term not in r.ViewName.lower() and term not in r.ViewKind.lower():
                continue
            self.replace_rows.Add(r)
        self._refresh_active_tab_summary()

    def ApplyReplaceButton_Click(self, s, a):
        src = self.filter_name_to_option.get(self.SourceComboBox.SelectedItem)
        tgt = self.filter_name_to_option.get(self.TargetComboBox.SelectedItem)
        if not src or not tgt:
            self._set_replace_status("Missing source/target.")
            return
        rows = [r for r in self.all_replace_rows if r.Apply]
        if not rows:
            self._set_replace_status("No rows selected.")
            return
        merge = self.MergeExistingCheckBox.IsChecked
        copyv = self.CopyVisibilityCheckBox.IsChecked
        copye = self.CopyEnabledCheckBox.IsChecked
        ok = 0
        sk = 0
        fl = 0
        tx = Transaction(doc, "Filter Manager Pro - Replace")
        tx.Start()
        try:
            for r in rows:
                v = doc.GetElement(safe_element_id(r.ViewId))
                if not v:
                    sk += 1
                    continue
                try:
                    vals = [element_id_value(x) for x in list(v.GetFilters())]
                except Exception:
                    sk += 1
                    continue
                hs = element_id_value(src.ElementId) in vals
                ht = element_id_value(tgt.ElementId) in vals
                if not hs:
                    sk += 1
                    continue
                try:
                    o = v.GetFilterOverrides(src.ElementId)
                    if not ht:
                        v.AddFilter(tgt.ElementId)
                    elif not merge:
                        sk += 1
                        continue
                    v.SetFilterOverrides(tgt.ElementId, o)
                    if copyv:
                        v.SetFilterVisibility(tgt.ElementId, v.GetFilterVisibility(src.ElementId))
                    if copye:
                        ge = getattr(v, "GetIsFilterEnabled", None)
                        se = getattr(v, "SetIsFilterEnabled", None)
                        if callable(ge) and callable(se):
                            se(tgt.ElementId, ge(src.ElementId))
                    v.RemoveFilter(src.ElementId)
                    ok += 1
                except Exception:
                    fl += 1
            tx.Commit()
        except Exception:
            try:
                tx.RollBack()
            except Exception:
                pass
            fl += 1
        self._load_audit()
        self.PreviewReplaceButton_Click(None, None)
        self._set_replace_status("Apply Replace complete. Updated: {} | Skipped: {} | Failed: {}".format(ok, sk, fl))

    def ExportAuditCsvButton_Click(self, s, a):
        self._export_csv("audit_summary.csv", ["Filter", "Categories", "Views", "Templates", "Total", "Status", "Duplicate Type", "Duplicate Set"], self.all_audit_rows, lambda r: [r.FilterName, r.Categories, r.ViewCount, r.TemplateCount, r.TotalCount, r.Status, r.DuplicateType, r.DuplicateGroup])

    def ExportUnusedCsvButton_Click(self, s, a):
        rows = [r for r in self.all_audit_rows if r.TotalCount == 0]
        self._export_csv("unused_filters.csv", ["Filter", "Categories"], rows, lambda r: [r.FilterName, r.Categories])

    def ExportReplaceCsvButton_Click(self, s, a):
        self._export_csv("replace_preview.csv", ["View", "Type", "Template", "Source", "Target", "Apply"], self.all_replace_rows, lambda r: [r.ViewName, r.ViewKind, r.IsTemplate, r.HasSource, r.HasTarget, r.Apply])

    def _export_csv(self, filename, header, rows, rowf):
        path = forms.save_file(file_ext='csv', default_name=filename)
        if not path:
            return
        import csv
        try:
            with open(path, 'wb') as f:
                w = csv.writer(f)
                w.writerow(header)
                for r in rows:
                    w.writerow([str(x) for x in rowf(r)])
            self._set_reports_status("Exported: {}".format(path))
        except Exception as ex:
            self._set_reports_status("Export failed: {}".format(ex))

    def _set_audit_status(self, t):
        try:
            self.AuditStatusTextBlock.Text = t
        except Exception:
            pass

    def _set_audit_details_columns(self, filter_text, duplicate_text, rules_text):
        wrote_new = False
        try:
            self.AuditDetailsFilterTextBlock.Text = filter_text
            wrote_new = True
        except Exception:
            pass
        try:
            self.AuditDetailsDuplicateTextBlock.Text = duplicate_text
            wrote_new = True
        except Exception:
            pass
        try:
            self.AuditDetailsRulesTextBlock.Text = rules_text
            wrote_new = True
        except Exception:
            pass
        for viewer_name in ("AuditDetailsFilterScrollViewer", "AuditDetailsDuplicateScrollViewer", "AuditDetailsRulesScrollViewer"):
            try:
                getattr(self, viewer_name).ScrollToTop()
            except Exception:
                pass
        if not wrote_new:
            self._set_audit_details("FILTER\n{}\n\nDUPLICATE\n{}\n\nRULES\n{}".format(filter_text, duplicate_text, rules_text))

    def _set_audit_details(self, t):
        try:
            self.AuditDetailsTextBlock.Text = t
            return
        except Exception:
            pass
        try:
            self.AuditDetailsFilterTextBlock.Text = t
            self.AuditDetailsDuplicateTextBlock.Text = ""
            self.AuditDetailsRulesTextBlock.Text = ""
        except Exception:
            pass

    def _set_rename_status(self, t):
        self.RenameStatusTextBlock.Text = t

    def _set_replace_status(self, t):
        self.ReplaceStatusTextBlock.Text = t

    def _set_reports_status(self, t):
        self.ReportsStatusTextBlock.Text = t


FilterManagerProWindow().ShowDialog()

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

from System.Collections.ObjectModel import ObservableCollection
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System.Windows import Visibility

from Autodesk.Revit.DB import View, Element, ElementId, Transaction
from pyrevit import forms, revit, script


doc = revit.doc
XAML_FILE = script.get_bundle_file("filter_manager_pro.xaml")
LOGO_FILE = get_filters_logo_path()


class FilterOption(object):
    def __init__(self, element_id, name):
        self.ElementId = element_id
        self.Name = name

    def __str__(self):
        return self.Name


class AuditRow(object):
    def __init__(self, filter_name, categories, view_count, template_count):
        self.FilterName = filter_name
        self.Categories = categories
        self.ViewCount = view_count
        self.TemplateCount = template_count
        self.TotalCount = view_count + template_count
        self.Status = "Used" if self.TotalCount > 0 else "Unused"


class RenamePreviewRow(object):
    def __init__(self, current_name, proposed_name):
        self.CurrentName = current_name
        self.ProposedName = proposed_name


class ReplacePreviewRow(object):
    def __init__(self, view_name, view_kind, has_source, has_target):
        self.ViewName = view_name
        self.ViewKind = view_kind
        self.HasSource = has_source
        self.HasTarget = has_target


class FilterManagerProWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self._load_header_logo()
        self.filters = self._collect_filters()
        self.filter_name_to_option = {filt.Name: filt for filt in self.filters}
        self.filter_names = sorted(self.filter_name_to_option.keys(), key=lambda item: item.lower())
        self.audit_rows = ObservableCollection[object]()
        self.rename_rows = ObservableCollection[object]()
        self.replace_rows = ObservableCollection[object]()
        self.AuditGrid.ItemsSource = self.audit_rows
        self.RenameGrid.ItemsSource = self.rename_rows
        self.ReplaceGrid.ItemsSource = self.replace_rows
        self.SourceComboBox.ItemsSource = self.filter_names
        self.TargetComboBox.ItemsSource = self.filter_names
        if len(self.filter_names) > 0:
            self.SourceComboBox.SelectedIndex = 0
            self.TargetComboBox.SelectedIndex = 0
        self.replace_preview_ready = False
        self.replace_preview_view_ids = []
        self._load_audit()
        self._set_rename_status("Click Preview to load all filters.")
        self._set_replace_status("Select Source and Target filters, then click Preview Usage.")
        self._refresh_active_tab_summary()


    def _set_text(self, control_name, value):
        try:
            getattr(self, control_name).Text = str(value)
        except Exception:
            pass

    def _set_header_cards(self, cards):
        card_slots = 6
        for i in range(card_slots):
            data = cards[i] if i < len(cards) else None
            label = data[0] if data is not None else ""
            value = data[1] if data is not None else ""
            self._set_text("HeaderCardLabel{}".format(i + 1), label)
            self._set_text("HeaderCardValue{}".format(i + 1), value)
            try:
                border = getattr(self, "HeaderCardBorder{}".format(i + 1))
                border.Visibility = Visibility.Visible if data is not None else Visibility.Collapsed
            except Exception:
                pass

    def _card(self, label, value):
        return (label, str(value))

    def _refresh_active_tab_summary(self):
        selected = self.MainTabControl.SelectedItem
        header = ""
        try:
            header = str(selected.Header)
        except Exception:
            pass
        if "Audit" in header:
            total = len(self.audit_rows)
            used = len([x for x in self.audit_rows if x.TotalCount > 0])
            unused = len([x for x in self.audit_rows if x.TotalCount == 0])
            views = sum([x.ViewCount for x in self.audit_rows])
            templates = sum([x.TemplateCount for x in self.audit_rows])
            cards = [self._card("FILTERS", total), self._card("USED", used)]
            if unused > 0:
                cards.append(self._card("UNUSED", unused))
            if views > 0:
                cards.append(self._card("VIEWS", views))
            if templates > 0:
                cards.append(self._card("TEMPLATES", templates))
            self._set_header_cards(cards)
        elif "Rename" in header:
            total = len(self.rename_rows)
            no_change = len([x for x in self.rename_rows if x.CurrentName == x.ProposedName])
            ready = total - no_change
            issues = 0
            cards = [self._card("TOTAL", total)]
            if ready > 0:
                cards.append(self._card("READY", ready))
            if no_change > 0:
                cards.append(self._card("NO CHANGE", no_change))
            if issues > 0:
                cards.append(self._card("ISSUES", issues))
            self._set_header_cards(cards)
        elif "Replace" in header:
            source_name = self.SourceComboBox.SelectedItem
            target_name = self.TargetComboBox.SelectedItem
            source = self.filter_name_to_option.get(source_name) if source_name else None
            target = self.filter_name_to_option.get(target_name) if target_name else None
            ready = 1 if (source is not None and target is not None and element_id_value(source.ElementId) != element_id_value(target.ElementId)) else 0
            issues = 0 if ready else 1
            affected = len(self.replace_rows)
            cards = [self._card("FILTERS", len(self.filters))]
            if affected > 0:
                cards.append(self._card("AFFECTED", affected))
            if ready > 0:
                cards.append(self._card("READY", ready))
            if issues > 0:
                cards.append(self._card("ISSUES", issues))
            self._set_header_cards(cards)
        else:
            cards = [self._card("REPORTS", "Available")]
            self._set_header_cards(cards)

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

    def _collect_filters(self):
        items = []
        for filt in collect_parameter_filters(doc, key_selector=lambda item: element_name(item).lower()):
            items.append(FilterOption(filt.Id, element_name(filt)))
        return items

    def _view_supports_filters(self, view):
        try:
            view.GetFilters()
            return True
        except Exception:
            return False

    def _load_audit(self):
        self.audit_rows.Clear()
        views = collect_views_with_filters(doc)
        for filt in self.filters:
            view_count = 0
            template_count = 0
            for view in views:
                try:
                    ids = list(view.GetFilters())
                except Exception:
                    continue
                values = [element_id_value(x) for x in ids]
                if element_id_value(filt.ElementId) not in values:
                    continue
                if view.IsTemplate:
                    template_count += 1
                else:
                    view_count += 1
            self.audit_rows.Add(AuditRow(filt.Name, self._get_filter_categories_text(filt.ElementId), view_count, template_count))
        self._refresh_active_tab_summary()

    def _preview_rename(self):
        self.rename_rows.Clear()
        prefix = (self.RenamePrefixTextBox.Text or "").strip()
        for filt in self.filters:
            proposed = filt.Name
            if prefix:
                proposed = "{}{}".format(prefix, filt.Name)
            self.rename_rows.Add(RenamePreviewRow(filt.Name, proposed))
        if prefix:
            self._set_rename_status("Previewing {} filters with prefix '{}'".format(len(self.filters), prefix))
        else:
            self._set_rename_status("Previewing {} filters. Prefix is empty, so proposed names match current names.".format(len(self.filters)))
        self._refresh_active_tab_summary()

    def _preview_replace(self):
        self.replace_rows.Clear()
        self.replace_preview_ready = False
        self.replace_preview_view_ids = []
        source_name = self.SourceComboBox.SelectedItem
        target_name = self.TargetComboBox.SelectedItem
        source = self.filter_name_to_option.get(source_name) if source_name else None
        target = self.filter_name_to_option.get(target_name) if target_name else None
        if source is None or target is None:
            self._set_replace_status("Select both Source and Target filters.")
            self._refresh_active_tab_summary()
            return
        if element_id_value(source.ElementId) == element_id_value(target.ElementId):
            self._set_replace_status("Source and Target are the same filter. Select different filters to compare usage.")
            self._refresh_active_tab_summary()
            return
        source_id = element_id_value(source.ElementId)
        target_id = element_id_value(target.ElementId)
        for view in collect_views_with_filters(doc):
            try:
                ids = [element_id_value(x) for x in list(view.GetFilters())]
            except Exception:
                continue
            has_source = source_id in ids
            has_target = target_id in ids
            if not has_source and not has_target:
                continue
            try:
                view_kind = str(view.ViewType)
            except Exception:
                view_kind = "Unknown"
            self.replace_rows.Add(ReplacePreviewRow(element_name(view), view_kind, has_source, has_target))
        self.replace_preview_ready = True
        self.replace_preview_view_ids = [element_id_value(v.Id) for v in collect_views_with_filters(doc) if self._is_view_in_replace_preview(v, source_id, target_id)]
        self._set_replace_status("Preview shows {} views/templates affected by Source or Target.".format(len(self.replace_rows)))
        self._refresh_active_tab_summary()


    def _is_view_in_replace_preview(self, view, source_id, target_id):
        try:
            ids = [element_id_value(x) for x in list(view.GetFilters())]
        except Exception:
            return False
        return (source_id in ids) or (target_id in ids)

    def _build_rename_plan(self):
        plan = []
        touched_ids = set()
        for row in self.rename_rows:
            current = (row.CurrentName or "").strip()
            proposed = (row.ProposedName or "").strip()
            if current == proposed:
                continue
            option = self.filter_name_to_option.get(current)
            if option is None:
                continue
            plan.append((option, current, proposed))
            touched_ids.add(element_id_value(option.ElementId))
        return plan, touched_ids

    def ApplyRenameButton_Click(self, sender, args):
        plan, touched_ids = self._build_rename_plan()
        if len(plan) == 0:
            self._set_rename_status("No rename actions to apply.")
            return

        proposed_names = []
        for _, _, proposed in plan:
            if not proposed:
                self._set_rename_status("Apply failed: Proposed name cannot be empty.")
                return
            proposed_names.append(proposed)

        if len(set([x.lower() for x in proposed_names])) != len(proposed_names):
            self._set_rename_status("Apply failed: Proposed names must be unique.")
            return

        existing = {}
        for filt in collect_parameter_filters(doc, key_selector=lambda item: element_name(item).lower()):
            existing[element_id_value(filt.Id)] = element_name(filt)
        existing_outside = [name.lower() for fid, name in existing.items() if fid not in touched_ids]
        for proposed in proposed_names:
            if proposed.lower() in existing_outside:
                self._set_rename_status("Apply failed: Proposed names conflict with existing filters outside the rename set.")
                return

        renamed = 0
        skipped = 0
        failed = 0
        tx = Transaction(doc, "Filter Manager Pro - Apply Rename")
        tx.Start()
        try:
            for option, current, proposed in plan:
                elem = doc.GetElement(option.ElementId)
                if elem is None:
                    skipped += 1
                    continue
                try:
                    elem.Name = proposed
                    renamed += 1
                except Exception:
                    failed += 1
            tx.Commit()
        except Exception:
            try:
                tx.RollBack()
            except Exception:
                pass
            failed = len(plan)

        self.filters = self._collect_filters()
        self.filter_name_to_option = {filt.Name: filt for filt in self.filters}
        self.filter_names = sorted(self.filter_name_to_option.keys(), key=lambda item: item.lower())
        self.SourceComboBox.ItemsSource = self.filter_names
        self.TargetComboBox.ItemsSource = self.filter_names
        self._load_audit()
        self._preview_rename()
        self._set_rename_status("Apply Rename complete. Renamed: {} | Skipped: {} | Failed: {}".format(renamed, skipped, failed))

    def ApplyReplaceButton_Click(self, sender, args):
        if not self.replace_preview_ready or len(self.replace_rows) == 0:
            self._set_replace_status("Run Preview Usage first before applying replace.")
            return

        source_name = self.SourceComboBox.SelectedItem
        target_name = self.TargetComboBox.SelectedItem
        source = self.filter_name_to_option.get(source_name) if source_name else None
        target = self.filter_name_to_option.get(target_name) if target_name else None
        if source is None or target is None:
            self._set_replace_status("Select both Source and Target filters.")
            return
        source_id = source.ElementId
        target_id = target.ElementId
        if element_id_value(source_id) == element_id_value(target_id):
            self._set_replace_status("Source and Target must be different filters.")
            return

        updated = 0
        skipped = 0
        failed = 0
        tx = Transaction(doc, "Filter Manager Pro - Apply Replace")
        tx.Start()
        try:
            for vid in self.replace_preview_view_ids:
                view = doc.GetElement(ElementId(vid))
                if view is None:
                    skipped += 1
                    continue
                try:
                    current_ids = list(view.GetFilters())
                except Exception:
                    skipped += 1
                    continue
                values = [element_id_value(x) for x in current_ids]
                has_source = element_id_value(source_id) in values
                has_target = element_id_value(target_id) in values
                if not has_source:
                    skipped += 1
                    continue
                try:
                    source_overrides = view.GetFilterOverrides(source_id)
                except Exception:
                    source_overrides = None
                try:
                    source_visible = view.GetFilterVisibility(source_id)
                except Exception:
                    source_visible = None
                get_enabled = getattr(view, "GetIsFilterEnabled", None)
                set_enabled = getattr(view, "SetIsFilterEnabled", None)
                source_enabled = None
                if callable(get_enabled):
                    try:
                        source_enabled = get_enabled(source_id)
                    except Exception:
                        source_enabled = None
                try:
                    if not has_target:
                        view.AddFilter(target_id)
                    if source_overrides is not None:
                        try:
                            view.SetFilterOverrides(target_id, source_overrides)
                        except Exception:
                            pass
                    if source_visible is not None:
                        try:
                            view.SetFilterVisibility(target_id, source_visible)
                        except Exception:
                            pass
                    if source_enabled is not None and callable(set_enabled):
                        try:
                            set_enabled(target_id, source_enabled)
                        except Exception:
                            pass
                    view.RemoveFilter(source_id)
                    updated += 1
                except Exception:
                    failed += 1
            tx.Commit()
        except Exception:
            try:
                tx.RollBack()
            except Exception:
                pass
            failed += 1

        self.filters = self._collect_filters()
        self.filter_name_to_option = {filt.Name: filt for filt in self.filters}
        self.filter_names = sorted(self.filter_name_to_option.keys(), key=lambda item: item.lower())
        self.SourceComboBox.ItemsSource = self.filter_names
        self.TargetComboBox.ItemsSource = self.filter_names
        self._load_audit()
        self._preview_replace()
        self._set_replace_status("Apply Replace complete. Updated: {} | Skipped: {} | Failed: {}".format(updated, skipped, failed))
    def _set_rename_status(self, text):
        self.RenameStatusTextBlock.Text = text

    def _set_replace_status(self, text):
        self.ReplaceStatusTextBlock.Text = text

    def _get_filter_categories_text(self, filter_id):
        try:
            filt = doc.GetElement(filter_id)
            category_ids = []
            get_categories = getattr(filt, "GetCategories", None)
            if callable(get_categories):
                try:
                    category_ids = list(get_categories())
                except Exception:
                    category_ids = []
            if not category_ids:
                categories_prop = getattr(filt, "Categories", None)
                if categories_prop is not None:
                    try:
                        category_ids = list(categories_prop)
                    except Exception:
                        category_ids = []
            names = []
            for cat_id in category_ids:
                cat_name = self._resolve_category_name(cat_id)
                if cat_name:
                    names.append(cat_name)
            if not names:
                return "N/A"
            names = sorted(list(set(names)))
            if len(names) <= 3:
                return ", ".join(names)
            return "{} categories".format(len(names))
        except Exception:
            return "N/A"

    def _resolve_category_name(self, category_ref):
        try:
            if category_ref is None:
                return None
            category_obj = None
            if hasattr(category_ref, "Name"):
                category_obj = category_ref
            else:
                cat_id = category_ref
                if hasattr(category_ref, "Id"):
                    cat_id = category_ref.Id
                try:
                    category_obj = doc.Settings.Categories.get_Item(cat_id)
                except Exception:
                    category_obj = None
                if category_obj is None:
                    try:
                        cat_elem = doc.GetElement(cat_id)
                        if cat_elem is not None and hasattr(cat_elem, "Name"):
                            return cat_elem.Name
                    except Exception:
                        pass
            if category_obj is not None and getattr(category_obj, "Name", None):
                return category_obj.Name
            return None
        except Exception:
            return None

    def RefreshAuditButton_Click(self, sender, args):
        self.replace_preview_ready = False
        self.replace_preview_view_ids = []
        self._load_audit()

    def PreviewRenameButton_Click(self, sender, args):
        self._preview_rename()

    def PreviewReplaceButton_Click(self, sender, args):
        self._preview_replace()

    def MainTabControl_SelectionChanged(self, sender, args):
        self._refresh_active_tab_summary()


def collect_views_with_filters(current_doc):
    views = []
    for view in AutodeskViews(current_doc):
        try:
            if view.ViewType.ToString() in ("ProjectBrowser", "SystemBrowser", "Undefined", "Internal"):
                continue
        except Exception:
            pass
        try:
            view.GetFilters()
            views.append(view)
        except Exception:
            continue
    return views


def AutodeskViews(current_doc):
    from Autodesk.Revit.DB import FilteredElementCollector
    return list(FilteredElementCollector(current_doc).OfClass(View))


FilterManagerProWindow().ShowDialog()

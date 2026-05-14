# -*- coding: utf-8 -*-

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

from Autodesk.Revit.DB import View, Element
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
    def __init__(self, filter_name, view_count, template_count):
        self.FilterName = filter_name
        self.ViewCount = view_count
        self.TemplateCount = template_count
        self.TotalCount = view_count + template_count


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
        self.audit_rows = ObservableCollection[object]()
        self.rename_rows = ObservableCollection[object]()
        self.replace_rows = ObservableCollection[object]()
        self.AuditGrid.ItemsSource = self.audit_rows
        self.RenameGrid.ItemsSource = self.rename_rows
        self.ReplaceGrid.ItemsSource = self.replace_rows
        self.SourceComboBox.ItemsSource = self.filters
        self.TargetComboBox.ItemsSource = self.filters
        if len(self.filters) > 0:
            self.SourceComboBox.SelectedIndex = 0
            self.TargetComboBox.SelectedIndex = 0
        self._load_audit()

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
        for filt in self.filters:
            view_count = 0
            template_count = 0
            for view in collect_views_with_filters(doc):
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
            self.audit_rows.Add(AuditRow(filt.Name, view_count, template_count))

    def _preview_rename(self):
        self.rename_rows.Clear()
        prefix = (self.RenamePrefixTextBox.Text or "").strip()
        for filt in self.filters:
            proposed = filt.Name
            if prefix:
                proposed = "{}{}".format(prefix, filt.Name)
            self.rename_rows.Add(RenamePreviewRow(filt.Name, proposed))

    def _preview_replace(self):
        self.replace_rows.Clear()
        source = self.SourceComboBox.SelectedItem
        target = self.TargetComboBox.SelectedItem
        if source is None or target is None:
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

    def RefreshAuditButton_Click(self, sender, args):
        self._load_audit()

    def PreviewRenameButton_Click(self, sender, args):
        self._preview_rename()

    def PreviewReplaceButton_Click(self, sender, args):
        self._preview_replace()


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

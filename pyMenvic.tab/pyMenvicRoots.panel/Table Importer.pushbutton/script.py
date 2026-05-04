# -*- coding: utf-8 -*-
__title__ = "Table Importer"
__author__ = "Menvic"

import os
import sys
import shutil

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System")

from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Window, Visibility, MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Forms import OpenFileDialog, DialogResult
import System
from System import Uri
from System.Windows.Controls import DataGridRow, ListBoxItem, StackPanel, TextBlock
from System.Windows.Input import MouseButtonEventHandler
from System.Windows.Media import VisualTreeHelper
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption

from pyrevit import script

script_dir = os.path.dirname(__file__)
if script_dir not in sys.path:
    sys.path.append(script_dir)

TOOL_VERSION = "MVP 0.2.1"

from models import TableEntry
from storage import load_entries, save_entries, get_storage_path
from excel_reader import (
    get_excel_worksheets,
    get_last_modified,
    get_file_name_without_extension,
    get_excel_regions,
)

USED_RANGE_DISPLAY = u"Full Worksheet Used Range"
DEFAULT_IMPORT_TYPES = ["Excel Link", "Excel Import", "Image"]
DEFAULT_VIEW_TYPES = ["Drafting View", "Legend View"]
DEFAULT_SCALES = ["1", "2", "5", "10", "20", "25", "50", "75", "100"]
DEFAULT_DPI = ["72", "96", "150", "200", "300", "600"]
EXCEL_FILTER = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"


def safe_unicode(value):
    if value is None:
        return u""
    try:
        if isinstance(value, unicode):
            return value
    except Exception:
        pass
    try:
        if isinstance(value, str):
            for enc in ("utf-8", "cp1252", "latin-1"):
                try:
                    return value.decode(enc)
                except Exception:
                    pass
            return value.decode("utf-8", "replace")
    except Exception:
        pass
    try:
        return unicode(value)
    except Exception:
        try:
            return unicode(value.ToString())
        except Exception:
            return u""


def safe_bool(value):
    try:
        if isinstance(value, bool):
            return value
    except Exception:
        pass
    text = safe_unicode(value).strip().lower()
    return text in (u"1", u"true", u"yes", u"si", u"sí", u"y")


def load_xaml(xaml_path):
    stream = FileStream(xaml_path, FileMode.Open, FileAccess.Read)
    try:
        return XamlReader.Load(stream)
    finally:
        stream.Close()


def get_storage_folder():
    try:
        return os.path.dirname(get_storage_path())
    except Exception:
        appdata = os.getenv("APPDATA") or script_dir
        return os.path.join(appdata, "pyMenvic", "TableImporter")


def resolve_entry_path(file_path):
    path = safe_unicode(file_path)
    if not path:
        return ""
    if os.path.isabs(path):
        return path
    return os.path.abspath(os.path.join(get_storage_folder(), path))


def find_visual_parent(child, parent_type):
    try:
        parent = VisualTreeHelper.GetParent(child)
        while parent is not None:
            if isinstance(parent, parent_type):
                return parent
            parent = VisualTreeHelper.GetParent(parent)
    except Exception:
        pass
    return None


def open_with_windows_shell(path):
    if not path:
        return False
    try:
        os.startfile(path)
        return True
    except Exception:
        pass
    try:
        psi = System.Diagnostics.ProcessStartInfo()
        psi.FileName = path
        psi.UseShellExecute = True
        System.Diagnostics.Process.Start(psi)
        return True
    except Exception:
        return False


def same_path(a, b):
    try:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))
    except Exception:
        return safe_unicode(a).lower() == safe_unicode(b).lower()


def get_combo_text(combo):
    try:
        item = combo.SelectedItem
        if item is None:
            return u""
        return safe_unicode(item)
    except Exception:
        return u""


def fill_combo(combo, values, selected=None):
    combo.Items.Clear()
    for value in values:
        combo.Items.Add(safe_unicode(value))
    if selected:
        for i in range(combo.Items.Count):
            if safe_unicode(combo.Items[i]) == safe_unicode(selected):
                combo.SelectedIndex = i
                return
    if combo.Items.Count:
        combo.SelectedIndex = 0

def get_shared_logo_path():
    extension_dir = os.path.abspath(os.path.join(script_dir, "..", "..", ".."))
    return os.path.join(extension_dir, "_resources", "logos", "menvic_logo.png")


def load_logo_into_image(image_control):
    try:
        if image_control is None:
            return
        logo_path = get_shared_logo_path()
        if not os.path.exists(logo_path):
            image_control.Source = None
            return
        bitmap = BitmapImage()
        bitmap.BeginInit()
        bitmap.UriSource = Uri(logo_path, System.UriKind.Absolute)
        bitmap.CacheOption = BitmapCacheOption.OnLoad
        bitmap.EndInit()
        image_control.Source = bitmap
    except Exception:
        try:
            image_control.Source = None
        except Exception:
            pass


class AddTableDialog(object):
    def __init__(self, owner=None):
        self.xaml_path = os.path.join(script_dir, "AddTableWindow.xaml")
        self.window = load_xaml(self.xaml_path)
        if owner:
            self.window.Owner = owner

        self.file_paths = []
        self.result_entries = []
        self.active_file_path = None

        self.FilePathTextBox = self.window.FindName("FilePathTextBox")
        self.FilesListBox = self.window.FindName("FilesListBox")
        self.NoFilesTextBlock = self.window.FindName("NoFilesTextBlock")
        self.BrowseButton = self.window.FindName("BrowseButton")
        self.WorksheetComboBox = self.window.FindName("WorksheetComboBox")
        self.RegionComboBox = self.window.FindName("RegionComboBox")
        self.ImportTypeComboBox = self.window.FindName("ImportTypeComboBox")
        self.ViewTypeComboBox = self.window.FindName("ViewTypeComboBox")
        self.ViewScaleComboBox = self.window.FindName("ViewScaleComboBox")
        self.DpiComboBox = self.window.FindName("DpiComboBox")
        self.BlackAndWhiteCheckBox = self.window.FindName("BlackAndWhiteCheckBox")
        self.AutoSyncCheckBox = self.window.FindName("AutoSyncCheckBox")
        self.CopiesTextBox = self.window.FindName("CopiesTextBox")
        self.StatusTextBlock = self.window.FindName("StatusTextBlock")
        self.OkButton = self.window.FindName("OkButton")
        self.CancelButton = self.window.FindName("CancelButton")
        self.LogoImage = self.window.FindName("LogoImage")
        self.VersionTextBlock = self.window.FindName("VersionTextBlock")
        if self.VersionTextBlock:
            self.VersionTextBlock.Text = TOOL_VERSION

        self.setup_defaults()
        self.bind_events()

    def setup_defaults(self):
        fill_combo(self.RegionComboBox, [USED_RANGE_DISPLAY], USED_RANGE_DISPLAY)
        fill_combo(self.ImportTypeComboBox, DEFAULT_IMPORT_TYPES, "Excel Link")
        fill_combo(self.ViewTypeComboBox, DEFAULT_VIEW_TYPES, "Drafting View")
        fill_combo(self.ViewScaleComboBox, DEFAULT_SCALES, "1")
        fill_combo(self.DpiComboBox, DEFAULT_DPI, "150")
        self.BlackAndWhiteCheckBox.IsChecked = True
        self.AutoSyncCheckBox.IsChecked = False
        self.CopiesTextBox.Text = "1"
        load_logo_into_image(self.LogoImage)
        self.update_file_controls()
        self.OkButton.IsEnabled = False



    def is_valid_excel_file(self, file_path):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            return os.path.exists(file_path) and ext in (".xlsx", ".xlsm", ".xls")
        except Exception:
            return False

    def get_valid_file_paths(self):
        result = []
        for path in self.file_paths:
            if self.is_valid_excel_file(path):
                result.append(path)
        return result

    def update_file_controls(self):
        valid_count = len(self.get_valid_file_paths())
        self.OkButton.IsEnabled = valid_count > 0
        if self.NoFilesTextBlock is not None:
            self.NoFilesTextBlock.Visibility = Visibility.Visible if valid_count == 0 else Visibility.Collapsed

    def add_file_list_item(self, file_path):
        item = ListBoxItem()
        item.Tag = file_path
        item.ToolTip = file_path
        panel = StackPanel()
        name_text = TextBlock()
        name_text.Text = os.path.basename(file_path)
        name_text.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
        path_text = TextBlock()
        path_text.Text = file_path
        path_text.FontSize = 11
        path_text.Foreground = System.Windows.Media.Brushes.Gray
        path_text.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
        panel.Children.Add(name_text)
        panel.Children.Add(path_text)
        item.Content = panel
        self.FilesListBox.Items.Add(item)

    def get_selected_file_path(self):
        try:
            item = self.FilesListBox.SelectedItem
            if item is not None and item.Tag is not None:
                return safe_unicode(item.Tag)
        except Exception:
            pass
        return u""
    def bind_events(self):
        self.BrowseButton.Click += self.on_browse
        self.FilesListBox.SelectionChanged += self.on_file_selected
        self.WorksheetComboBox.SelectionChanged += self.on_worksheet_changed
        self.OkButton.Click += self.on_ok
        self.CancelButton.Click += self.on_cancel

    def on_browse(self, sender, args):
        dialog = OpenFileDialog()
        dialog.Title = "Select Excel files"
        dialog.Filter = EXCEL_FILTER
        dialog.Multiselect = True
        result = dialog.ShowDialog()
        if result == DialogResult.OK:
            self.file_paths = [safe_unicode(x) for x in dialog.FileNames]
            self.refresh_file_list()

    def refresh_file_list(self):
        self.FilesListBox.Items.Clear()
        valid_paths = self.get_valid_file_paths()
        for path in valid_paths:
            self.add_file_list_item(path)
        self.update_file_controls()
        if valid_paths:
            self.FilePathTextBox.Text = "%s valid file(s) selected" % len(valid_paths)
            self.FilesListBox.SelectedIndex = 0
            self.load_worksheets(valid_paths[0])
            self.StatusTextBlock.Text = "%s file(s) loaded." % len(valid_paths)
        else:
            self.FilePathTextBox.Text = ""
            fill_combo(self.WorksheetComboBox, [], None)
            fill_combo(self.RegionComboBox, [USED_RANGE_DISPLAY], USED_RANGE_DISPLAY)
            self.StatusTextBlock.Text = "No files selected."

    def on_file_selected(self, sender, args):
        try:
            path = self.get_selected_file_path()
            if path:
                self.load_worksheets(path)
        except Exception:
            pass

    def load_worksheets(self, file_path):
        self.active_file_path = file_path
        worksheets = get_excel_worksheets(file_path)
        if not worksheets:
            self.WorksheetComboBox.Items.Clear()
            fill_combo(self.RegionComboBox, [USED_RANGE_DISPLAY], USED_RANGE_DISPLAY)
            self.StatusTextBlock.Text = "No worksheets found or Excel could not be opened."
            return
        fill_combo(self.WorksheetComboBox, worksheets, worksheets[0])
        self.load_region_for_selected_worksheet()
        self.StatusTextBlock.Text = "%s worksheet(s) loaded for %s." % (len(worksheets), os.path.basename(file_path))

    def on_worksheet_changed(self, sender, args):
        self.load_region_for_selected_worksheet()

    def load_region_for_selected_worksheet(self):
        worksheet = get_combo_text(self.WorksheetComboBox)
        if not self.active_file_path or not worksheet:
            fill_combo(self.RegionComboBox, [USED_RANGE_DISPLAY], USED_RANGE_DISPLAY)
            return
        regions = get_excel_regions(self.active_file_path, worksheet)
        if not regions:
            regions = [USED_RANGE_DISPLAY]
        fill_combo(self.RegionComboBox, regions, regions[0])

    def get_copies(self):
        try:
            value = int(self.CopiesTextBox.Text)
            if value < 1:
                return 1
            return value
        except Exception:
            return 1

    def get_sheet_for_file(self, file_path, requested_sheet):
        sheets = get_excel_worksheets(file_path)
        if requested_sheet in sheets:
            return requested_sheet
        if sheets:
            return sheets[0]
        return u""

    def get_region_for_file(self, file_path, worksheet, requested_region):
        regions = get_excel_regions(file_path, worksheet)
        if not regions:
            regions = [USED_RANGE_DISPLAY]
        if requested_region in regions:
            return requested_region, regions
        return regions[0], regions

    def on_ok(self, sender, args):
        valid_paths = self.get_valid_file_paths()
        if not valid_paths:
            self.StatusTextBlock.Text = "Please select one or more Excel files."
            return

        requested_sheet = get_combo_text(self.WorksheetComboBox)
        requested_region = get_combo_text(self.RegionComboBox) or USED_RANGE_DISPLAY
        import_type = get_combo_text(self.ImportTypeComboBox) or "Excel Link"
        view_type = get_combo_text(self.ViewTypeComboBox) or "Drafting View"
        view_scale = get_combo_text(self.ViewScaleComboBox) or "1"
        dpi = get_combo_text(self.DpiComboBox) or "150"
        black_white = safe_bool(self.BlackAndWhiteCheckBox.IsChecked)
        auto_sync = safe_bool(self.AutoSyncCheckBox.IsChecked)
        copies = self.get_copies()

        entries = []
        for file_path in valid_paths:
            if not file_path or not os.path.exists(file_path):
                continue
            worksheet = self.get_sheet_for_file(file_path, requested_sheet)
            if not worksheet:
                continue
            region, regions = self.get_region_for_file(file_path, worksheet, requested_region)
            base_name = get_file_name_without_extension(file_path)
            for index in range(copies):
                suffix = "" if copies == 1 else " %s" % (index + 1)
                view_name = "%s - %s%s" % (base_name, worksheet, suffix)
                entries.append(TableEntry(
                    selected=True,
                    status="Not Created",
                    source=os.path.basename(file_path),
                    import_type=import_type,
                    view_name=view_name,
                    dpi=dpi,
                    auto_sync=auto_sync,
                    black_and_white=black_white,
                    last_modified=get_last_modified(file_path),
                    worksheet=worksheet,
                    region=region,
                    region_options=list(regions),
                    view_type=view_type,
                    view_scale=view_scale,
                    file_path=file_path,
                    path_mode="Absolute",
                    revit_view_id=None,
                ))

        if not entries:
            self.StatusTextBlock.Text = "No valid table rows could be created."
            return

        self.result_entries = entries
        self.window.DialogResult = True
        self.window.Close()

    def on_cancel(self, sender, args):
        self.result_entries = []
        self.window.DialogResult = False
        self.window.Close()

    def show_dialog(self):
        self.window.ShowDialog()
        return self.result_entries


class TableImporterWindow(object):
    def __init__(self):
        self.xaml_path = os.path.join(script_dir, "TableImporter.xaml")
        self.window = load_xaml(self.xaml_path)
        self.all_entries = []
        self.entries = ObservableCollection[object]()
        self.current_search = u""

        self.TablesDataGrid = self.window.FindName("TablesDataGrid")
        self.AddTablesButton = self.window.FindName("AddTablesButton")
        self.EmptyStateButton = self.window.FindName("EmptyStateButton")
        self.EmptyStatePanel = self.window.FindName("EmptyStatePanel")
        self.BatchActionsButton = self.window.FindName("BatchActionsButton")
        self.RefreshButton = self.window.FindName("RefreshButton")
        self.SearchTextBox = self.window.FindName("SearchTextBox")
        self.FooterStatusTextBlock = self.window.FindName("FooterStatusTextBlock")
        self.CompletedTextBlock = self.window.FindName("CompletedTextBlock")
        self.CloseButton = self.window.FindName("CloseButton")
        self.ApplyButton = self.window.FindName("ApplyButton")
        self.LogoImage = self.window.FindName("LogoImage")
        self.VersionTextBlock = self.window.FindName("VersionTextBlock")
        if self.VersionTextBlock:
            self.VersionTextBlock.Text = TOOL_VERSION

        self.BatchActionsContextMenu = self.BatchActionsButton.ContextMenu
        self.RowActionsContextMenu = self.TablesDataGrid.ContextMenu
        self.bind_menu_items()
        self.load_logo_image()
        self.load_saved_entries()
        self.bind_events()

    def bind_menu_items(self):
        names = [
            "UpdateViews", "DuplicateViews", "ReloadFrom", "AbsolutePath", "RelativePath",
            "OpenFiles", "OpenFolders", "DeleteViews", "UnlinkView", "OpenView"
        ]
        for name in names:
            setattr(self, "Batch%sMenuItem" % name, self.window.FindName("Batch%sMenuItem" % name))
            setattr(self, "Row%sMenuItem" % name, self.window.FindName("Row%sMenuItem" % name))

    def load_logo_image(self):
        load_logo_into_image(self.LogoImage)

    def bind_events(self):
        self.AddTablesButton.Click += self.on_add_tables
        self.EmptyStateButton.Click += self.on_add_tables
        self.RefreshButton.Click += self.on_refresh
        self.CloseButton.Click += self.on_close
        self.ApplyButton.Click += self.on_apply
        self.BatchActionsButton.Click += self.on_batch_actions
        self.SearchTextBox.TextChanged += self.on_search_changed
        self.TablesDataGrid.PreviewMouseLeftButtonDown += self.on_datagrid_single_click
        self.TablesDataGrid.CellEditEnding += self.on_cell_edit_ending
        self.TablesDataGrid.Drop += self.on_drop_files
        try:
            self.EmptyStatePanel.Drop += self.on_drop_files
            self.EmptyStateButton.Drop += self.on_drop_files
        except Exception:
            pass

        for prefix in ("Batch", "Row"):
            self.window.FindName(prefix + "UpdateViewsMenuItem").Click += self.on_update_views
            self.window.FindName(prefix + "DuplicateViewsMenuItem").Click += self.on_duplicate_views
            self.window.FindName(prefix + "ReloadFromMenuItem").Click += self.on_reload_from
            self.window.FindName(prefix + "AbsolutePathMenuItem").Click += self.on_absolute_path
            self.window.FindName(prefix + "RelativePathMenuItem").Click += self.on_relative_path
            self.window.FindName(prefix + "OpenFilesMenuItem").Click += self.on_open_files
            self.window.FindName(prefix + "OpenFoldersMenuItem").Click += self.on_open_folders
            self.window.FindName(prefix + "DeleteViewsMenuItem").Click += self.on_delete_views
            self.window.FindName(prefix + "UnlinkViewMenuItem").Click += self.on_unlink_view
            self.window.FindName(prefix + "OpenViewMenuItem").Click += self.on_open_view

    def on_datagrid_single_click(self, sender, args):
        try:
            row = find_visual_parent(args.OriginalSource, DataGridRow)
            if row is not None and not row.IsEditing:
                self.TablesDataGrid.BeginEdit()
        except Exception:
            pass

    def on_cell_edit_ending(self, sender, args):
        try:
            if args.EditAction.ToString() == "Commit":
                self.save_current_entries()
                self.FooterStatusTextBlock.Text = "Changes saved."
        except Exception:
            pass

    def load_saved_entries(self):
        self.all_entries = []
        for entry in load_entries():
            self.prepare_entry(entry)
            self.all_entries.append(entry)
        self.apply_filter()

    def prepare_entry(self, entry):
        try:
            if not entry.Source:
                entry.Source = os.path.basename(resolve_entry_path(entry.FilePath))
        except Exception:
            pass
        self.populate_region_options(entry)

    def populate_region_options(self, entry):
        try:
            path = resolve_entry_path(entry.FilePath)
            if path and os.path.exists(path) and entry.Worksheet:
                options = get_excel_regions(path, entry.Worksheet)
            else:
                options = []
            current = safe_unicode(entry.Region)
            if current and current not in options:
                options.insert(0, current)
            if not options:
                options = [USED_RANGE_DISPLAY]
            entry.RegionOptions = options
            if not entry.Region:
                entry.Region = options[0]
        except Exception:
            try:
                entry.RegionOptions = [entry.Region or USED_RANGE_DISPLAY]
            except Exception:
                pass

    def save_current_entries(self):
        save_entries(self.all_entries)

    def update_footer(self):
        total = len(self.all_entries)
        visible = self.entries.Count
        selected = len(self.get_target_entries(False))
        linked = 0
        drafting = 0
        legend = 0
        for entry in self.all_entries:
            try:
                if entry.RevitViewId:
                    linked += 1
                if entry.ViewType == "Legend View":
                    legend += 1
                else:
                    drafting += 1
            except Exception:
                pass
        self.FooterStatusTextBlock.Text = "Total %s | Visible %s | Selected %s | Linked %s | Drafting %s | Legends %s" % (total, visible, selected, linked, drafting, legend)
        pct = 0
        if total:
            pct = int((float(linked) / float(total)) * 100.0)
        self.CompletedTextBlock.Text = "Completed %s%%" % pct
        self.update_empty_state()

    def update_empty_state(self):
        if self.EmptyStatePanel is None:
            return
        if self.entries.Count == 0:
            self.EmptyStatePanel.Visibility = Visibility.Visible
        else:
            self.EmptyStatePanel.Visibility = Visibility.Collapsed

    def matches_search(self, entry, text):
        if not text:
            return True
        values = [entry.ViewName, entry.Worksheet, entry.Region, entry.FilePath]
        for value in values:
            if text in safe_unicode(value).lower():
                return True
        return False

    def apply_filter(self):
        self.entries.Clear()
        text = safe_unicode(self.current_search).lower().strip()
        for entry in self.all_entries:
            if self.matches_search(entry, text):
                self.entries.Add(entry)
        self.TablesDataGrid.ItemsSource = self.entries
        self.TablesDataGrid.Items.Refresh()
        self.update_footer()

    def add_entries(self, new_entries):
        for entry in new_entries:
            self.prepare_entry(entry)
            self.all_entries.append(entry)
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "%s table row(s) added." % len(new_entries)

    def on_add_tables(self, sender, args):
        dialog = AddTableDialog(owner=self.window)
        entries = dialog.show_dialog()
        if entries:
            self.add_entries(entries)

    def on_drop_files(self, sender, args):
        try:
            if not args.Data.GetDataPresent(System.Windows.DataFormats.FileDrop):
                return
            paths = list(args.Data.GetData(System.Windows.DataFormats.FileDrop))
            excel_paths = []
            for path in paths:
                ext = os.path.splitext(path)[1].lower()
                if ext in (".xlsx", ".xlsm", ".xls"):
                    excel_paths.append(path)
            if not excel_paths:
                self.FooterStatusTextBlock.Text = "Drop Excel files only."
                return
            entries = []
            for path in excel_paths:
                sheets = get_excel_worksheets(path)
                if not sheets:
                    continue
                sheet = sheets[0]
                regions = get_excel_regions(path, sheet)
                region = regions[0] if regions else USED_RANGE_DISPLAY
                base = get_file_name_without_extension(path)
                entries.append(TableEntry(
                    selected=True,
                    status="Not Created",
                    source=os.path.basename(path),
                    import_type="Excel Link",
                    view_name="%s - %s" % (base, sheet),
                    dpi="150",
                    auto_sync=False,
                    black_and_white=True,
                    last_modified=get_last_modified(path),
                    worksheet=sheet,
                    region=region,
                    region_options=regions,
                    view_type="Drafting View",
                    view_scale="1",
                    file_path=path,
                    path_mode="Absolute",
                    revit_view_id=None,
                ))
            if entries:
                self.add_entries(entries)
        except Exception as ex:
            self.FooterStatusTextBlock.Text = "Drop failed: %s" % safe_unicode(ex)

    def on_refresh(self, sender, args):
        updated = 0
        missing = 0
        for entry in self.all_entries:
            path = resolve_entry_path(entry.FilePath)
            if not path or not os.path.exists(path):
                entry.Status = "Missing File"
                entry.LastModified = ""
                missing += 1
                continue
            self.populate_region_options(entry)
            old_date = entry.LastModified
            new_date = get_last_modified(path)
            entry.LastModified = new_date
            entry.Source = os.path.basename(path)
            if old_date and new_date and old_date != new_date:
                entry.Status = "Modified"
                updated += 1
            elif entry.RevitViewId:
                entry.Status = "OK"
            else:
                entry.Status = "Not Created"
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "Refresh complete. %s modified, %s missing." % (updated, missing)

    def on_apply(self, sender, args):
        self.save_current_entries()
        self.FooterStatusTextBlock.Text = "Settings saved. Revit view creation is still a placeholder."

    def on_batch_actions(self, sender, args):
        try:
            self.BatchActionsContextMenu.PlacementTarget = self.BatchActionsButton
            self.BatchActionsContextMenu.IsOpen = True
        except Exception:
            self.FooterStatusTextBlock.Text = "Batch Actions menu could not be opened."

    def get_target_entries(self, allow_selected_rows=True):
        result = []
        try:
            for entry in self.all_entries:
                if entry.Selected:
                    result.append(entry)
        except Exception:
            pass
        if result or not allow_selected_rows:
            return result
        try:
            for item in self.TablesDataGrid.SelectedItems:
                if item not in result:
                    result.append(item)
        except Exception:
            pass
        return result

    def require_targets(self):
        targets = self.get_target_entries(True)
        if not targets:
            self.FooterStatusTextBlock.Text = "Select one or more tables first."
        return targets

    def on_update_views(self, sender, args):
        targets = self.require_targets()
        if targets:
            self.save_current_entries()
            self.FooterStatusTextBlock.Text = "Update Views placeholder. %s table(s) selected." % len(targets)

    def duplicate_entry(self, entry):
        copied = TableEntry.from_dict(entry.to_dict())
        copied.Selected = True
        copied.Status = "Not Created"
        copied.ViewName = "%s Copy" % safe_unicode(entry.ViewName)
        copied.RevitViewId = None
        return copied

    def on_duplicate_views(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        new_entries = []
        for entry in targets:
            try:
                new_entries.append(self.duplicate_entry(entry))
            except Exception:
                pass
        for entry in new_entries:
            self.all_entries.append(entry)
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "%s duplicate table row(s) added." % len(new_entries)

    def on_reload_from(self, sender, args):
        targets = self.require_targets()
        if targets:
            self.FooterStatusTextBlock.Text = "Reload From placeholder. %s table(s) selected." % len(targets)

    def on_absolute_path(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        changed = 0
        base_folder = get_storage_folder()
        for entry in targets:
            try:
                path = safe_unicode(entry.FilePath)
                if path and not os.path.isabs(path):
                    entry.FilePath = os.path.abspath(os.path.join(base_folder, path))
                    changed += 1
                entry.PathMode = "Absolute"
            except Exception:
                pass
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "%s path(s) set to absolute." % changed

    def on_relative_path(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        changed = 0
        base_folder = get_storage_folder()
        for entry in targets:
            try:
                path = safe_unicode(entry.FilePath)
                if path and os.path.isabs(path):
                    entry.FilePath = os.path.relpath(path, base_folder)
                    changed += 1
                entry.PathMode = "Relative"
            except Exception:
                pass
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "%s path(s) set to relative." % changed

    def on_open_files(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        opened = 0
        missing = 0
        for entry in targets:
            path = resolve_entry_path(entry.FilePath)
            if path and os.path.exists(path) and open_with_windows_shell(path):
                opened += 1
            else:
                missing += 1
        self.FooterStatusTextBlock.Text = "Opened %s file(s). %s missing." % (opened, missing)

    def on_open_folders(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        opened_folders = []
        missing = 0
        for entry in targets:
            path = resolve_entry_path(entry.FilePath)
            folder = os.path.dirname(path) if path else ""
            if folder and os.path.exists(folder):
                already = False
                for open_folder in opened_folders:
                    if same_path(open_folder, folder):
                        already = True
                        break
                if not already and open_with_windows_shell(folder):
                    opened_folders.append(folder)
            else:
                missing += 1
        self.FooterStatusTextBlock.Text = "Opened %s folder(s). %s missing." % (len(opened_folders), missing)

    def on_delete_views(self, sender, args):
        targets = self.require_targets()
        if not targets:
            return
        result = MessageBox.Show(
            "Remove %s selected table row(s)?\n\nThis does not delete Revit views yet." % len(targets),
            "Delete Views",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
        )
        if result != MessageBoxResult.Yes:
            self.FooterStatusTextBlock.Text = "Delete cancelled."
            return
        removed = 0
        for entry in list(targets):
            try:
                if entry in self.all_entries:
                    self.all_entries.remove(entry)
                    removed += 1
            except Exception:
                pass
        self.save_current_entries()
        self.apply_filter()
        self.FooterStatusTextBlock.Text = "%s table row(s) removed." % removed

    def on_unlink_view(self, sender, args):
        targets = self.require_targets()
        if targets:
            self.FooterStatusTextBlock.Text = "Unlink View placeholder. %s table(s) selected." % len(targets)

    def on_open_view(self, sender, args):
        targets = self.require_targets()
        if targets:
            self.FooterStatusTextBlock.Text = "Open View placeholder. %s table(s) selected." % len(targets)

    def on_search_changed(self, sender, args):
        self.current_search = safe_unicode(self.SearchTextBox.Text)
        self.apply_filter()
        if self.current_search:
            self.FooterStatusTextBlock.Text = "Search: %s | %s visible." % (self.current_search, self.entries.Count)

    def on_close(self, sender, args):
        self.save_current_entries()
        self.window.Close()

    def show(self):
        self.window.ShowDialog()


if __name__ == "__main__":
    try:
        ui = TableImporterWindow()
        ui.show()
    except Exception as ex:
        output = script.get_output()
        output.print_md("## Table Importer Error")
        output.print_md("```")
        output.print_md(safe_unicode(ex))
        output.print_md("```")











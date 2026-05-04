# -*- coding: utf-8 -*-
__title__ = "Table Importer"
__author__ = "Menvic"

import os
import sys

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System")

from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Forms import OpenFileDialog, DialogResult
import System
from System.Windows.Controls import DataGridRow

from System.Windows.Media import VisualTreeHelper

from pyrevit import script

# Local imports
script_dir = os.path.dirname(__file__)
if script_dir not in sys.path:
    sys.path.append(script_dir)

from models import TableEntry
from storage import load_entries, save_entries
from excel_reader import (
    get_excel_worksheets,
    get_last_modified,
    get_file_name_without_extension,
    get_used_range_address,
    get_excel_regions,
)

USED_RANGE_KEY = u"Used Range"
USED_RANGE_DISPLAY = u"Full Worksheet Used Range"


def safe_unicode(value):
    """Return safe Unicode text in IronPython, including Windows-encoded accents."""
    if value is None:
        return u""

    # IronPython 2: unicode is text, str is bytes.
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
            try:
                return value.decode("utf-8", "replace")
            except Exception:
                return u""
    except Exception:
        pass

    # .NET strings / COM objects
    try:
        return unicode(value)
    except Exception:
        pass

    try:
        return unicode(value.ToString())
    except Exception:
        pass

    try:
        raw = str(value)
        for enc in ("utf-8", "cp1252", "latin-1"):
            try:
                return raw.decode(enc)
            except Exception:
                pass
    except Exception:
        pass

    return u""


def safe_ascii_text(value):
    text = safe_unicode(value)
    if not text:
        return u""
    try:
        import unicodedata
        text = unicodedata.normalize('NFKD', text)
        text = u''.join([c for c in text if not unicodedata.combining(c)])
    except Exception:
        pass
    replacements = {
        u"Ñ": u"N", u"ñ": u"n",
        u"Á": u"A", u"É": u"E", u"Í": u"I", u"Ó": u"O", u"Ú": u"U",
        u"á": u"a", u"é": u"e", u"í": u"i", u"ó": u"o", u"ú": u"u",
    }
    for a, b in replacements.items():
        text = text.replace(a, b)
    cleaned = []
    for ch in text:
        try:
            code = ord(ch)
            if 32 <= code <= 126:
                cleaned.append(ch)
        except Exception:
            pass
    return u"".join(cleaned).strip()


def load_xaml(xaml_path):
    stream = FileStream(xaml_path, FileMode.Open, FileAccess.Read)
    try:
        return XamlReader.Load(stream)
    finally:
        stream.Close()




def get_storage_folder():
    try:
        from storage import get_storage_path
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
    """Walk up the visual tree to find a parent of a given type."""
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
    """Open a file or folder using Windows shell without importing Process directly."""
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

class AddTableDialog(object):
    def __init__(self, owner=None):
        self.xaml_path = os.path.join(script_dir, "AddTableWindow.xaml")
        self.window = load_xaml(self.xaml_path)

        if owner:
            self.window.Owner = owner

        self.file_path = None
        self.result_entry = None

        self.FilePathTextBox = self.window.FindName("FilePathTextBox")
        self.BrowseButton = self.window.FindName("BrowseButton")
        self.WorksheetComboBox = self.window.FindName("WorksheetComboBox")
        self.RegionComboBox = self.window.FindName("RegionComboBox")
        self.ViewTypeComboBox = self.window.FindName("ViewTypeComboBox")
        self.ViewScaleComboBox = self.window.FindName("ViewScaleComboBox")
        self.CopiesTextBox = self.window.FindName("CopiesTextBox")
        self.StatusTextBlock = self.window.FindName("StatusTextBlock")
        self.OkButton = self.window.FindName("OkButton")
        self.CancelButton = self.window.FindName("CancelButton")

        self.setup_defaults()
        self.bind_events()

    def setup_defaults(self):
        # Region is populated after an Excel file and worksheet are selected.
        # Keep a safe default so the UI never starts empty.
        self.RegionComboBox.Items.Add(USED_RANGE_DISPLAY)
        self.RegionComboBox.SelectedIndex = 0

        self.ViewTypeComboBox.Items.Add("Drafting View")
        self.ViewTypeComboBox.Items.Add("Legend View")
        self.ViewTypeComboBox.SelectedIndex = 0

        scales = ["1", "2", "5", "10", "20", "25", "50", "75", "100"]
        for scale in scales:
            self.ViewScaleComboBox.Items.Add(scale)
        self.ViewScaleComboBox.SelectedIndex = 0

        self.CopiesTextBox.Text = "1"

    def bind_events(self):
        self.BrowseButton.Click += self.on_browse
        self.WorksheetComboBox.SelectionChanged += self.on_worksheet_changed
        self.OkButton.Click += self.on_ok
        self.CancelButton.Click += self.on_cancel

    def on_browse(self, sender, args):
        dialog = OpenFileDialog()
        dialog.Title = "Select Excel file"
        dialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls"
        dialog.Multiselect = False

        result = dialog.ShowDialog()

        if result == DialogResult.OK:
            self.file_path = dialog.FileName
            self.FilePathTextBox.Text = self.file_path
            self.load_worksheets()

    def load_worksheets(self):
        self.WorksheetComboBox.Items.Clear()

        worksheets = get_excel_worksheets(self.file_path)

        if not worksheets:
            self.StatusTextBlock.Text = "No worksheets found or Excel could not be opened."
            return

        for worksheet in worksheets:
            self.WorksheetComboBox.Items.Add(worksheet)

        self.WorksheetComboBox.SelectedIndex = 0
        self.load_region_for_selected_worksheet()
        self.StatusTextBlock.Text = "%s worksheet(s) loaded." % len(worksheets)

    def on_worksheet_changed(self, sender, args):
        self.load_region_for_selected_worksheet()

    def load_region_for_selected_worksheet(self):
        self.RegionComboBox.Items.Clear()

        worksheet = self.get_selected_combo_text(self.WorksheetComboBox)
        if not self.file_path or not worksheet:
            self.RegionComboBox.Items.Add(USED_RANGE_DISPLAY)
            self.RegionComboBox.SelectedIndex = 0
            return

        regions = get_excel_regions(self.file_path, worksheet)
        if not regions:
            regions = [USED_RANGE_DISPLAY]

        for region in regions:
            self.RegionComboBox.Items.Add(safe_ascii_text(region))
        self.RegionComboBox.SelectedIndex = 0

    def get_selected_combo_text(self, combo):
        item = combo.SelectedItem
        if item is None:
            return ""
        return safe_unicode(item)

    def get_copies(self):
        try:
            value = int(self.CopiesTextBox.Text)
            if value < 1:
                return 1
            return value
        except Exception:
            return 1

    def on_ok(self, sender, args):
        file_path = self.FilePathTextBox.Text

        if not file_path or not os.path.exists(file_path):
            self.StatusTextBlock.Text = "Please select a valid Excel file."
            return

        worksheet = self.get_selected_combo_text(self.WorksheetComboBox)
        if not worksheet:
            self.StatusTextBlock.Text = "Please select a worksheet."
            return

        view_type = self.get_selected_combo_text(self.ViewTypeComboBox)
        view_scale = self.get_selected_combo_text(self.ViewScaleComboBox)

        region_address = safe_ascii_text(self.get_selected_combo_text(self.RegionComboBox))
        if not region_address:
            region_address = USED_RANGE_DISPLAY

        base_name = get_file_name_without_extension(file_path)
        view_name = "%s - %s" % (base_name, worksheet)

        self.result_entry = TableEntry(
            selected=True,
            status="Not Created",
            view_name=view_name,
            auto_sync=False,
            last_modified=get_last_modified(file_path),
            worksheet=worksheet,
            region=region_address,
            region_options=[safe_ascii_text(x) for x in self.RegionComboBox.Items],
            view_type=view_type,
            view_scale=view_scale,
            file_path=file_path,
            revit_view_id=None,
        )

        self.window.DialogResult = True
        self.window.Close()

    def on_cancel(self, sender, args):
        self.result_entry = None
        self.window.DialogResult = False
        self.window.Close()

    def show_dialog(self):
        self.window.ShowDialog()
        return self.result_entry


class TableImporterWindow(object):
    def __init__(self):
        self.xaml_path = os.path.join(script_dir, "TableImporter.xaml")
        self.window = load_xaml(self.xaml_path)

        self.entries = ObservableCollection[object]()

        self.TablesDataGrid = self.window.FindName("TablesDataGrid")
        self.AddTablesButton = self.window.FindName("AddTablesButton")
        self.BatchActionsButton = self.window.FindName("BatchActionsButton")
        self.RefreshButton = self.window.FindName("RefreshButton")
        self.SearchTextBox = self.window.FindName("SearchTextBox")
        self.FooterStatusTextBlock = self.window.FindName("FooterStatusTextBlock")
        self.CloseButton = self.window.FindName("CloseButton")
        self.ApplyButton = self.window.FindName("ApplyButton")

        self.BatchActionsContextMenu = self.BatchActionsButton.ContextMenu
        self.BatchUpdateViewsMenuItem = self.window.FindName("BatchUpdateViewsMenuItem")
        self.BatchDuplicateViewsMenuItem = self.window.FindName("BatchDuplicateViewsMenuItem")
        self.BatchAbsolutePathMenuItem = self.window.FindName("BatchAbsolutePathMenuItem")
        self.BatchRelativePathMenuItem = self.window.FindName("BatchRelativePathMenuItem")
        self.BatchOpenFilesMenuItem = self.window.FindName("BatchOpenFilesMenuItem")
        self.BatchOpenFoldersMenuItem = self.window.FindName("BatchOpenFoldersMenuItem")
        self.BatchDeleteViewsMenuItem = self.window.FindName("BatchDeleteViewsMenuItem")

        self.load_saved_entries()
        self.bind_events()

    def bind_events(self):
        self.AddTablesButton.Click += self.on_add_tables
        self.RefreshButton.Click += self.on_refresh
        self.CloseButton.Click += self.on_close
        self.ApplyButton.Click += self.on_apply
        self.BatchActionsButton.Click += self.on_batch_actions
        self.SearchTextBox.TextChanged += self.on_search_changed

        self.BatchUpdateViewsMenuItem.Click += self.on_batch_update_views
        self.BatchDuplicateViewsMenuItem.Click += self.on_batch_duplicate_views
        self.BatchAbsolutePathMenuItem.Click += self.on_batch_absolute_path
        self.BatchRelativePathMenuItem.Click += self.on_batch_relative_path
        self.BatchOpenFilesMenuItem.Click += self.on_batch_open_files
        self.BatchOpenFoldersMenuItem.Click += self.on_batch_open_folders
        self.BatchDeleteViewsMenuItem.Click += self.on_batch_delete_views

        # Single-click to begin editing
        self.TablesDataGrid.PreviewMouseLeftButtonDown += self.on_datagrid_single_click
        # Auto-save when a cell finishes editing
        self.TablesDataGrid.CellEditEnding += self.on_cell_edit_ending

    def on_datagrid_single_click(self, sender, args):
        """Begin editing on single click instead of requiring double-click."""
        try:
            hit = args.OriginalSource
            row = find_visual_parent(hit, DataGridRow)
            if row is not None:
                if not row.IsEditing:
                    self.TablesDataGrid.BeginEdit()
        except Exception:
            pass

    def on_cell_edit_ending(self, sender, args):
        """Auto-save to disk whenever a cell edit is committed."""
        try:
            if args.EditAction.ToString() == "Commit":
                self.save_current_entries()
                self.FooterStatusTextBlock.Text = "Changes saved."
        except Exception:
            pass

    def load_saved_entries(self):
        saved_entries = load_entries()

        for entry in saved_entries:
            self.populate_region_options(entry)
            self.entries.Add(entry)

        self.TablesDataGrid.ItemsSource = self.entries
        self.update_footer()

    def populate_region_options(self, entry):
        """Reload available regions for an existing row so Region can be changed later."""
        try:
            if entry.FilePath and os.path.exists(entry.FilePath) and entry.Worksheet:
                options = [safe_ascii_text(x) for x in get_excel_regions(entry.FilePath, entry.Worksheet)]
            else:
                options = []

            current = safe_ascii_text(entry.Region)
            if current and current not in options:
                options.insert(0, current)
            if not options:
                options = [USED_RANGE_DISPLAY]

            entry.RegionOptions = [safe_ascii_text(x) for x in options]
            if not entry.Region:
                entry.Region = options[0]
        except Exception:
            try:
                entry.RegionOptions = [entry.Region or USED_RANGE_DISPLAY]
            except Exception:
                pass

    def save_current_entries(self):
        save_entries(list(self.entries))

    def update_footer(self):
        count = self.entries.Count
        self.FooterStatusTextBlock.Text = "%s table(s) loaded." % count

    def on_add_tables(self, sender, args):
        dialog = AddTableDialog(owner=self.window)
        entry = dialog.show_dialog()

        if entry:
            self.populate_region_options(entry)
            self.entries.Add(entry)
            self.save_current_entries()
            self.update_footer()

    def on_refresh(self, sender, args):
        updated = 0

        for entry in self.entries:
            if not entry.FilePath or not os.path.exists(entry.FilePath):
                entry.Status = "Missing File"
                entry.LastModified = ""
                continue

            self.populate_region_options(entry)

            old_date = entry.LastModified
            new_date = get_last_modified(entry.FilePath)
            entry.LastModified = new_date

            if old_date and new_date and old_date != new_date:
                entry.Status = "Modified"
                updated += 1
            elif entry.RevitViewId:
                entry.Status = "OK"
            else:
                entry.Status = "Not Created"

        self.TablesDataGrid.Items.Refresh()
        self.save_current_entries()
        self.FooterStatusTextBlock.Text = "Refresh complete. %s modified table(s)." % updated

    def on_apply(self, sender, args):
        self.save_current_entries()
        self.FooterStatusTextBlock.Text = "Settings saved. Revit view creation comes in the next phase."

    def on_batch_actions(self, sender, args):
        """Open the Batch Actions menu."""
        try:
            self.BatchActionsContextMenu.PlacementTarget = self.BatchActionsButton
            self.BatchActionsContextMenu.IsOpen = True
        except Exception:
            self.FooterStatusTextBlock.Text = "Batch Actions menu could not be opened."

    def get_target_entries(self):
        """Use checked rows first. If none are checked, use selected DataGrid rows."""
        result = []
        try:
            for entry in self.entries:
                if entry.Selected:
                    result.append(entry)
        except Exception:
            pass

        if result:
            return result

        try:
            for item in self.TablesDataGrid.SelectedItems:
                result.append(item)
        except Exception:
            pass

        return result

    def on_batch_update_views(self, sender, args):
        targets = self.get_target_entries()
        if not targets:
            self.FooterStatusTextBlock.Text = "Select one or more tables first."
            return

        self.save_current_entries()
        self.FooterStatusTextBlock.Text = "Update Views is ready for the next phase. %s table(s) selected." % len(targets)

    def on_batch_duplicate_views(self, sender, args):
        targets = self.get_target_entries()
        if not targets:
            self.FooterStatusTextBlock.Text = "Select one or more tables first."
            return

        added = 0
        for entry in targets:
            try:
                copied = TableEntry(
                    selected=True,
                    status="Not Created",
                    view_name="%s Copy" % safe_unicode(entry.ViewName),
                    auto_sync=entry.AutoSync,
                    last_modified=entry.LastModified,
                    worksheet=entry.Worksheet,
                    region=entry.Region,
                    region_options=list(entry.RegionOptions),
                    view_type=entry.ViewType,
                    view_scale=entry.ViewScale,
                    file_path=entry.FilePath,
                    revit_view_id=None,
                )
                self.entries.Add(copied)
                added += 1
            except Exception:
                pass

        self.TablesDataGrid.Items.Refresh()
        self.save_current_entries()
        self.update_footer()
        self.FooterStatusTextBlock.Text = "%s duplicate table(s) added." % added

    def on_batch_absolute_path(self, sender, args):
        targets = self.get_target_entries()
        if not targets:
            self.FooterStatusTextBlock.Text = "Select one or more tables first."
            return

        changed = 0
        base_folder = get_storage_folder()
        for entry in targets:
            try:
                path = safe_unicode(entry.FilePath)
                if path and not os.path.isabs(path):
                    entry.FilePath = os.path.abspath(os.path.join(base_folder, path))
                    changed += 1
            except Exception:
                pass

        self.save_current_entries()
        self.TablesDataGrid.Items.Refresh()
        self.FooterStatusTextBlock.Text = "%s path(s) converted to absolute." % changed

    def on_batch_relative_path(self, sender, args):
        targets = self.get_target_entries()
        if not targets:
            self.FooterStatusTextBlock.Text = "Select one or more tables first."
            return

        changed = 0
        base_folder = get_storage_folder()
        for entry in targets:
            try:
                path = safe_unicode(entry.FilePath)
                if path and os.path.isabs(path):
                    entry.FilePath = os.path.relpath(path, base_folder)
                    changed += 1
            except Exception:
                pass

        self.save_current_entries()
        self.TablesDataGrid.Items.Refresh()
        self.FooterStatusTextBlock.Text = "%s path(s) converted to relative." % changed

    def on_batch_open_files(self, sender, args):
        targets = self.get_target_entries()
        if not targets:
            self.FooterStatusTextBlock.Text = "Select one or more tables first."
            return

        opened = 0
        missing = 0
        for entry in targets:
            try:
                path = resolve_entry_path(entry.FilePath)
                if path and os.path.exists(path):
                    open_with_windows_shell(path)
                    opened += 1
                else:
                    missing += 1
            except Exception:
                missing += 1

        self.FooterStatusTextBlock.Text = "Opened %s file(s). %s missing." % (opened, missing)

    def on_batch_open_folders(self, sender, args):
        targets = self.get_target_entries()
        if not targets:
            self.FooterStatusTextBlock.Text = "Select one or more tables first."
            return

        opened_folders = []
        missing = 0
        for entry in targets:
            try:
                path = resolve_entry_path(entry.FilePath)
                folder = os.path.dirname(path) if path else ""
                if folder and os.path.exists(folder) and folder not in opened_folders:
                    open_with_windows_shell(folder)
                    opened_folders.append(folder)
                elif not folder or not os.path.exists(folder):
                    missing += 1
            except Exception:
                missing += 1

        self.FooterStatusTextBlock.Text = "Opened %s folder(s). %s missing." % (len(opened_folders), missing)

    def on_batch_delete_views(self, sender, args):
        targets = self.get_target_entries()
        if not targets:
            self.FooterStatusTextBlock.Text = "Select one or more tables first."
            return

        cleared = 0
        for entry in targets:
            try:
                if entry.RevitViewId:
                    entry.RevitViewId = None
                    entry.Status = "Not Created"
                    cleared += 1
            except Exception:
                pass

        self.TablesDataGrid.Items.Refresh()
        self.save_current_entries()
        if cleared:
            self.FooterStatusTextBlock.Text = "%s created view reference(s) cleared. Revit deletion comes in the next phase." % cleared
        else:
            self.FooterStatusTextBlock.Text = "No created views selected. Revit deletion comes in the next phase."

    def on_search_changed(self, sender, args):
        text = self.SearchTextBox.Text
        if text:
            self.FooterStatusTextBlock.Text = "Search: %s" % safe_unicode(text)
        else:
            self.update_footer()

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

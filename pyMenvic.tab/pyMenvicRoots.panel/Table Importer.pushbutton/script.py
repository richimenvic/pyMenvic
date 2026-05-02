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

from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Forms import OpenFileDialog, DialogResult

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
)


def load_xaml(xaml_path):
    stream = FileStream(xaml_path, FileMode.Open, FileAccess.Read)
    try:
        return XamlReader.Load(stream)
    finally:
        stream.Close()


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
        self.RegionComboBox.Items.Add("Used Range")
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
        self.StatusTextBlock.Text = "%s worksheet(s) loaded." % len(worksheets)

    def get_selected_combo_text(self, combo):
        item = combo.SelectedItem
        if item is None:
            return ""
        return str(item)

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

        region_address = get_used_range_address(file_path, worksheet)
        if not region_address:
            region_address = "Used Range"

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

        self.load_saved_entries()
        self.bind_events()

    def bind_events(self):
        self.AddTablesButton.Click += self.on_add_tables
        self.RefreshButton.Click += self.on_refresh
        self.CloseButton.Click += self.on_close
        self.ApplyButton.Click += self.on_apply

        self.BatchActionsButton.Click += self.on_batch_actions
        self.SearchTextBox.TextChanged += self.on_search_changed

    def load_saved_entries(self):
        saved_entries = load_entries()

        for entry in saved_entries:
            self.entries.Add(entry)

        self.TablesDataGrid.ItemsSource = self.entries
        self.update_footer()

    def save_current_entries(self):
        save_entries(list(self.entries))

    def update_footer(self):
        count = self.entries.Count
        self.FooterStatusTextBlock.Text = "%s table(s) loaded." % count

    def on_add_tables(self, sender, args):
        dialog = AddTableDialog(owner=self.window)
        entry = dialog.show_dialog()

        if entry:
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
        # MVP 0.1: no Revit drawing yet.
        # For now, this only saves the current grid configuration.
        self.save_current_entries()
        self.FooterStatusTextBlock.Text = "Settings saved. Revit view creation comes in the next phase."

    def on_batch_actions(self, sender, args):
        self.FooterStatusTextBlock.Text = "Batch Actions will be added in the next phase."

    def on_search_changed(self, sender, args):
        # Simple MVP behavior: search is visual placeholder for now.
        text = self.SearchTextBox.Text
        if text:
            self.FooterStatusTextBlock.Text = "Search typed: %s" % text
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
        output.print_md(str(ex))
        output.print_md("```")
# -*- coding: utf-8 -*-

__title__ = "Levels + Grids to Standard Workset"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
LEVELS + GRIDS TO STANDARD WORKSET
_____________________________________________________

Description:

Moves Levels, Grids and related datum elements to the selected discipline standard workset.

_____________________________________________________
What the tool does:

• loads an XAML window from the button folder
• lists all user worksets in the current project
• audits Levels, Grids and Scope Boxes across the project
• groups selected workset contents by Category / Family / Type
• suggests a destination workset based on discipline and datum rules
• lets you assign a single action per row
• applies changes only to rows marked Apply with a valid destination

_____________________________________________________
Usage:

1. Open the tool
2. Select discipline and workset
3. Click Scan
4. Review grouped rows
5. Apply valid changes

_____________________________________________________
"""

import os
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')
clr.AddReference('System.Core')

from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind

from pyrevit import revit, DB, forms, script


doc = revit.doc


def get_logo_path():
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        while True:
            if os.path.basename(current_dir).lower() == "pymenvic.extension":
                logo_path = os.path.join(current_dir, "_resources", "logos", "menvic_logo.png")
                if os.path.exists(logo_path):
                    return logo_path
                return None
            parent_dir = os.path.dirname(current_dir)
            if parent_dir == current_dir:
                break
            current_dir = parent_dir
    except Exception:
        pass
    return None


def is_workshared_document(current_doc):
    try:
        return current_doc is not None and current_doc.IsWorkshared
    except Exception:
        return False
output = script.get_output()

TOOL_NAME = 'LEVELS + GRIDS TO STANDARD WORKSET'
DISCIPLINE_LABELS = ['ARCHITECTURE', 'STRUCTURE', 'MECHANICAL', 'ELECTRICAL', 'PLUMBING', 'SITE']
DISCIPLINE_CODE_BY_LABEL = {
    'ARCHITECTURE': 'ARC',
    'STRUCTURE': 'STR',
    'MECHANICAL': 'MECH',
    'ELECTRICAL': 'ELE',
    'PLUMBING': 'PLM',
    'SITE': 'SITE',
}
IGNORE_CATEGORY_NAMES = set([
    'Constraints',
    'Reference Planes',
])
SCOPE_BOX_CATEGORY_NAME = 'Scope Boxes'
DATUM_CLASSES = (DB.Level, DB.Grid)


# ==================================================
# HELPERS
# ==================================================

def first_line(ex):
    try:
        return str(ex).splitlines()[0]
    except:
        return 'Unknown error'


def get_user_worksets(current_doc):
    worksets = []
    collector = DB.FilteredWorksetCollector(current_doc).OfKind(DB.WorksetKind.UserWorkset)
    for ws in collector:
        worksets.append(ws)
    return sorted(worksets, key=lambda x: x.Name)


def get_workset_by_name(worksets, name):
    for ws in worksets:
        if ws.Name == name:
            return ws
    return None


def get_category_name(elem):
    try:
        if elem.Category:
            return elem.Category.Name or '<No Category>'
    except:
        pass
    return '<No Category>'


def get_family_name(elem):
    try:
        if elem.Category and elem.Category.Name == 'Levels':
            return 'Level'
        if elem.Category and elem.Category.Name == 'Grids':
            return 'Grid'
    except:
        pass

    try:
        etype = doc.GetElement(elem.GetTypeId())
        if etype and hasattr(etype, 'FamilyName'):
            return etype.FamilyName or '<System Family>'
    except:
        pass

    try:
        fam = getattr(elem, 'Symbol', None)
        if fam and fam.Family:
            return fam.Family.Name or '<Family>'
    except:
        pass

    return '<System / No Family>'


def get_type_name(elem):
    try:
        if hasattr(elem, 'Name') and elem.Name:
            return elem.Name
    except:
        pass

    try:
        etype = doc.GetElement(elem.GetTypeId())
        if etype and hasattr(etype, 'Name') and etype.Name:
            return etype.Name
    except:
        pass

    return '<Unnamed Type>'


def get_element_display_name(elem):
    try:
        if elem.Name:
            return elem.Name
    except:
        pass
    try:
        p = elem.get_Parameter(DB.BuiltInParameter.DATUM_TEXT)
        if p and p.HasValue:
            return p.AsString()
    except:
        pass
    return 'Element {}'.format(elem.Id.IntegerValue)


def get_all_non_type_elements(current_doc):
    return DB.FilteredElementCollector(current_doc).WhereElementIsNotElementType().ToElements()


def get_all_scope_boxes(current_doc):
    results = []
    try:
        for e in get_all_non_type_elements(current_doc):
            if get_category_name(e) == SCOPE_BOX_CATEGORY_NAME:
                results.append(e)
    except:
        pass
    return results


def should_ignore_element(elem):
    if elem is None:
        return True
    cname = get_category_name(elem)
    if cname in IGNORE_CATEGORY_NAMES:
        return True
    return False


def is_scope_box(elem):
    return get_category_name(elem) == SCOPE_BOX_CATEGORY_NAME


def is_datum_or_scopebox(elem):
    if isinstance(elem, DATUM_CLASSES):
        return True
    if is_scope_box(elem):
        return True
    return False


def get_discipline_code(discipline_label):
    if not discipline_label:
        return ""
    return DISCIPLINE_CODE_BY_LABEL.get(discipline_label, discipline_label)

def get_standard_workset_name(discipline, suffix):
    code = get_discipline_code(discipline)
    return "{}_{}".format(code, suffix)

def get_target_levels_grids_name(discipline_label):
    return get_standard_workset_name(discipline_label, 'LEVELS_GRIDS')


def get_missing_workset_message(workset_name):
    return 'Target workset not found in this model: {}'.format(workset_name)


def create_workset(current_doc, workset_name):
    if not workset_name:
        raise Exception('Workset name is empty.')
    t = DB.Transaction(current_doc, 'MENVIC - Create Workset')
    t.Start()
    try:
        workset = DB.Workset.Create(current_doc, workset_name)
        t.Commit()
        return workset
    except Exception:
        t.RollBack()
        raise

def suggest_destination_name(discipline, elem, available_names):
    if not discipline:
        return ''

    level_grid_name = get_standard_workset_name(discipline, 'LEVELS_GRIDS')
    model_name = get_standard_workset_name(discipline, 'MODEL')

    if is_datum_or_scopebox(elem):
        if level_grid_name in available_names:
            return level_grid_name
    else:
        if model_name in available_names:
            return model_name

    return ''


def get_workset_param(elem):
    try:
        return elem.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
    except:
        return None


def get_current_workset_name(elem, workset_name_by_id):
    try:
        wsid = elem.WorksetId.IntegerValue
        if wsid in workset_name_by_id:
            return workset_name_by_id[wsid]
    except:
        pass
    return '<Unknown Workset>'


def build_available_destinations(current_workset_name, workset_names, suggested_name):
    names = []
    for ws_name in workset_names:
        if ws_name == current_workset_name:
            continue
        names.append(ws_name)

    if suggested_name and suggested_name != current_workset_name and suggested_name not in names:
        names.insert(0, suggested_name)

    return names


def row_requires_attention(row):
    if row is None:
        return False
    if row.SelectedAction == 'Ignore':
        return False
    if row.SelectedAction == 'Review':
        return True
    if row.SelectedAction == 'Apply':
        if not row.SelectedDestination:
            return True
        return row.SelectedDestination != row.CurrentWorkset
    return True


def build_workset_maps(worksets):
    by_id = {}
    by_name = {}
    names = []
    for ws in worksets:
        by_id[ws.Id.IntegerValue] = ws.Name
        by_name[ws.Name] = ws
        names.append(ws.Name)
    return by_id, by_name, names


def load_logo(window, base_dir):
    image = getattr(window, 'logoImage', None)
    if image is None:
        return

    preferred_names = []
    candidate = None

    for name in preferred_names:
        path = os.path.join(base_dir, name)
        if os.path.exists(path):
            candidate = path
            break

    if candidate is None:
        try:
            for fname in os.listdir(base_dir):
                lower = fname.lower()
                if lower.endswith('.png') or lower.endswith('.jpg') or lower.endswith('.jpeg') or lower.endswith('.bmp'):
                    candidate = os.path.join(base_dir, fname)
                    break
        except:
            candidate = None

    if candidate:
        try:
            bmp = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource = Uri(candidate, UriKind.Absolute)
            bmp.EndInit()
            image.Source = bmp
        except:
            pass


# ==================================================
# DATA MODELS
# ==================================================

class WorksetRow(object):
    def __init__(self):
        self.CategoryName = ''
        self.FamilyName = ''
        self.TypeName = ''
        self.Count = 0
        self.CurrentWorkset = ''
        self.SuggestedDestination = ''
        self.SelectedDestination = ''
        self.AvailableDestinations = []
        self.AvailableActions = ['Apply', 'Review', 'Ignore']
        self.SelectedAction = 'Review'
        self.IsOmitted = False
        self.TargetWorksetExists = True
        self.MissingTargetWorksetName = ''
        self.ElementIds = []


# ==================================================
# CORE SCAN
# ==================================================

def audit_project(discipline, workset_names):
    rows = []

    levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level).WhereElementIsNotElementType().ToElements())
    grids = list(DB.FilteredElementCollector(doc).OfClass(DB.Grid).WhereElementIsNotElementType().ToElements())
    scope_boxes = list(get_all_scope_boxes(doc))

    expected_name = get_standard_workset_name(discipline, 'LEVELS_GRIDS') if discipline else ''

    levels_out = 0
    grids_out = 0
    scope_boxes_out = 0

    for e in levels:
        if get_current_workset_name(e, WINDOW.ws_name_by_id) != expected_name:
            levels_out += 1
    for e in grids:
        if get_current_workset_name(e, WINDOW.ws_name_by_id) != expected_name:
            grids_out += 1
    for e in scope_boxes:
        if get_current_workset_name(e, WINDOW.ws_name_by_id) != expected_name:
            scope_boxes_out += 1

    return {
        'levels_total': len(levels),
        'grids_total': len(grids),
        'scope_boxes_total': len(scope_boxes),
        'levels_out': levels_out,
        'grids_out': grids_out,
        'scope_boxes_out': scope_boxes_out,
        'expected_levels_grids': expected_name,
    }


def inspect_workset(discipline, selected_workset_name, workset_names):
    grouped = {}
    skipped_count = 0
    warnings = []

    expected_name = selected_workset_name or get_target_levels_grids_name(discipline)
    target_exists = expected_name in WINDOW.ws_by_name

    if not expected_name:
        return [], 0, ['No target workset is defined for the selected discipline.']

    if not target_exists:
        warnings.append(get_missing_workset_message(expected_name))

    all_elems = get_all_non_type_elements(doc)

    for elem in all_elems:
        try:
            if should_ignore_element(elem):
                skipped_count += 1
                continue

            if not is_datum_or_scopebox(elem):
                skipped_count += 1
                continue

            cat = get_category_name(elem)
            if cat == '<No Category>':
                skipped_count += 1
                continue

            curr = get_current_workset_name(elem, WINDOW.ws_name_by_id)
            fam = get_family_name(elem)
            typ = get_type_name(elem)

            suggested = expected_name if target_exists else 'Not found: {}'.format(expected_name)
            key = (cat, fam, typ, curr, suggested)

            if key not in grouped:
                row = WorksetRow()
                row.CategoryName = cat
                row.FamilyName = fam
                row.TypeName = typ
                row.Count = 0
                row.CurrentWorkset = curr
                row.SuggestedDestination = suggested
                row.TargetWorksetExists = target_exists
                row.MissingTargetWorksetName = '' if target_exists else expected_name
                row.AvailableDestinations = build_available_destinations(curr, workset_names, expected_name if target_exists else '')
                row.SelectedDestination = expected_name if (target_exists and expected_name != curr) else ''

                if not target_exists:
                    row.SelectedAction = 'Review'
                elif curr == expected_name:
                    row.SelectedAction = 'Ignore'
                else:
                    row.SelectedAction = 'Apply'
                grouped[key] = row

            grouped[key].Count += 1
            grouped[key].ElementIds.append(elem.Id)

        except Exception as ex:
            warnings.append(first_line(ex))

    rows = list(grouped.values())
    for row in rows:
        row.IsOmitted = not row_requires_attention(row)
    rows.sort(key=lambda r: (r.IsOmitted, r.CategoryName, r.FamilyName, r.TypeName))

    return rows, skipped_count, warnings


# ==================================================
# UI
# ==================================================

class InspectorWindow(Window):
    def __init__(self, xaml_path):
        fs = FileStream(xaml_path, FileMode.Open, FileAccess.Read)
        try:
            window = XamlReader.Load(fs)
        finally:
            fs.Close()

        self.Content = window.Content
        self.Title = window.Title
        self.Height = window.Height
        self.Width = window.Width
        self.MinHeight = window.MinHeight
        self.MinWidth = window.MinWidth
        self.Background = window.Background
        self.Foreground = window.Foreground
        self.ResizeMode = window.ResizeMode
        self.WindowStartupLocation = window.WindowStartupLocation

        # connect named controls
        for name in [
            'logoImage', 'discipline_combo', 'inspect_workset_combo',
            'levels_out_text', 'grids_out_text', 'scopeboxes_out_text',
            'rows_found_text', 'warnings_count_text', 'status_text', 'totals_summary_text',
            'results_grid', 'footer_text', 'show_omitted_checkbox',
            'scan_button', 'apply_all_button', 'close_button'
        ]:
            try:
                setattr(self, name, window.FindName(name))
            except:
                setattr(self, name, None)

        self.base_dir = os.path.dirname(xaml_path)
        self.worksets = get_user_worksets(doc)
        self.ws_name_by_id, self.ws_by_name, self.ws_names = build_workset_maps(self.worksets)
        self.rows = ObservableCollection[object]()
        self.all_rows = []
        if self.results_grid is not None:
            self.results_grid.ItemsSource = self.rows

        self._setup_ui()
        load_logo(self, self.base_dir)

        self.scan_button.Click += self.on_scan
        if self.show_omitted_checkbox is not None:
            self.show_omitted_checkbox.Checked += self.on_show_omitted_changed
            self.show_omitted_checkbox.Unchecked += self.on_show_omitted_changed
        self.apply_all_button.Click += self.on_apply_all
        self.close_button.Click += self.on_close
        self.discipline_combo.SelectionChanged += self.on_discipline_changed

    def activate_window(self):
        try:
            if self.WindowState:
                pass
        except:
            pass
        try:
            self.Activate()
            self.Focus()
        except:
            pass

    def show_message(self, text, title=None, image=MessageBoxImage.Information):
        try:
            MessageBox.Show(self, text, title or TOOL_NAME, MessageBoxButton.OK, image)
        except:
            forms.alert(text, title=title or TOOL_NAME)
        self.activate_window()

    def ask_yes_no(self, text, title=None, image=MessageBoxImage.Question):
        try:
            result = MessageBox.Show(self, text, title or TOOL_NAME, MessageBoxButton.YesNo, image)
            self.activate_window()
            return result == MessageBoxResult.Yes
        except:
            self.activate_window()
            return False

    def ensure_target_workset(self, workset_name):
        if not workset_name:
            return False

        self.refresh_worksets()
        if workset_name in self.ws_by_name:
            return True

        if not self.ask_yes_no(
            'The standard target workset does not exist yet:\n\n{}\n\nDo you want to create it now?'.format(workset_name)
        ):
            return False

        try:
            create_workset(doc, workset_name)
            self.refresh_worksets()
            if self.inspect_workset_combo is not None:
                self.inspect_workset_combo.SelectedItem = workset_name
            self.show_message('Created workset: {}'.format(workset_name), image=MessageBoxImage.Information)
            return True
        except Exception as ex:
            self.show_message(first_line(ex), image=MessageBoxImage.Error)
            return False

    def _setup_ui(self):
        if self.discipline_combo is not None:
            self.discipline_combo.ItemsSource = DISCIPLINE_LABELS
            self.discipline_combo.SelectedIndex = 0
        if self.inspect_workset_combo is not None:
            self.refresh_worksets()
        if self.status_text is not None:
            self.status_text.Text = 'Ready'
        if self.footer_text is not None:
            self.footer_text.Text = 'Select a discipline, review the target workset context, then scan. Correct rows stay hidden until you enable Show omitted items.'
        self.clear_stats()

    def clear_stats(self):
        for control in [
            self.levels_out_text, self.grids_out_text, self.scopeboxes_out_text,
            self.rows_found_text, self.warnings_count_text
        ]:
            if control is not None:
                control.Text = '0'
        if self.totals_summary_text is not None:
            self.totals_summary_text.Text = 'Totals: Levels 0 | Grids 0 | Scope Boxes 0'

    def get_selected_discipline(self):
        try:
            return str(self.discipline_combo.SelectedItem)
        except:
            return ''

    def get_selected_workset(self):
        try:
            return str(self.inspect_workset_combo.SelectedItem)
        except:
            return ''

    def refresh_worksets(self):
        current_selection = self.get_selected_workset()
        self.worksets = get_user_worksets(doc)
        self.ws_name_by_id, self.ws_by_name, self.ws_names = build_workset_maps(self.worksets)
        if self.inspect_workset_combo is None:
            return

        expected = get_target_levels_grids_name(self.get_selected_discipline())
        combo_items = list(self.ws_names)
        if expected and expected not in combo_items:
            combo_items.insert(0, expected)

        self.inspect_workset_combo.ItemsSource = None
        self.inspect_workset_combo.ItemsSource = combo_items

        if expected:
            self.inspect_workset_combo.SelectedItem = expected
        elif current_selection and current_selection in combo_items:
            self.inspect_workset_combo.SelectedItem = current_selection
        elif combo_items:
            self.inspect_workset_combo.SelectedIndex = 0

    def on_discipline_changed(self, sender, args):
        self.refresh_worksets()

    def get_show_omitted(self):
        try:
            return bool(self.show_omitted_checkbox.IsChecked)
        except:
            return False

    def repopulate_rows(self, scan_rows):
        if self.rows is None:
            return
        self.all_rows = list(scan_rows)
        self.rows.Clear()
        show_omitted = self.get_show_omitted()
        for row in self.all_rows:
            row.IsOmitted = not row_requires_attention(row)
            if show_omitted or not row.IsOmitted:
                self.rows.Add(row)

    def on_show_omitted_changed(self, sender, args):
        if self.get_show_omitted() and not self.all_rows:
            self.on_scan(sender, args)
            return
        self.repopulate_rows(self.all_rows)

    def on_scan(self, sender, args):
        try:
            self.refresh_worksets()
            discipline = self.get_selected_discipline()
            selected_workset = self.get_selected_workset()

            if not discipline:
                self.show_message('Select a discipline first.', image=MessageBoxImage.Warning)
                return
            if not selected_workset:
                self.show_message('Select a workset to inspect.', image=MessageBoxImage.Warning)
                return

            if selected_workset not in self.ws_by_name:
                created = self.ensure_target_workset(selected_workset)
                self.refresh_worksets()
                if created:
                    selected_workset = self.get_selected_workset()

            project_audit = audit_project(discipline, self.ws_names)
            scan_rows, ignored_count, warnings = inspect_workset(discipline, selected_workset, self.ws_names)

            self.repopulate_rows(scan_rows)

            element_count = 0
            visible_rows = 0
            omitted_rows = 0
            for row in scan_rows:
                element_count += row.Count
                if row.IsOmitted:
                    omitted_rows += 1
                else:
                    visible_rows += 1

            self.levels_out_text.Text = str(project_audit['levels_out'])
            self.grids_out_text.Text = str(project_audit['grids_out'])
            self.scopeboxes_out_text.Text = str(project_audit['scope_boxes_out'])
            self.rows_found_text.Text = str(visible_rows)
            self.warnings_count_text.Text = str(len(warnings))

            status = 'Scan complete'
            target_name = selected_workset
            if target_name not in self.ws_by_name:
                status = get_missing_workset_message(target_name)
            elif warnings:
                status = warnings[0]
            if self.status_text is not None:
                self.status_text.Text = '{} | Elements: {} | Ignored: {}'.format(status, element_count, ignored_count)
            if self.totals_summary_text is not None:
                self.totals_summary_text.Text = 'Totals: Levels {} | Grids {} | Scope Boxes {}'.format(
                    project_audit['levels_total'],
                    project_audit['grids_total'],
                    project_audit['scope_boxes_total']
                )
            if self.footer_text is not None:
                self.footer_text.Text = 'Discipline: {} | Workset: {} | Visible rows: {} | Omitted rows: {}'.format(discipline, selected_workset, visible_rows, omitted_rows)

        except Exception as ex:
            self.show_message(first_line(ex), image=MessageBoxImage.Error)
            if self.status_text is not None:
                self.status_text.Text = 'Scan failed'

    def collect_rows_to_apply(self):
        result = []
        for row in list(self.all_rows):
            if row is None:
                continue
            if row.SelectedAction != 'Apply':
                continue
            if not getattr(row, 'TargetWorksetExists', True):
                continue
            if not row.SelectedDestination:
                continue
            if row.SelectedDestination == row.CurrentWorkset:
                continue
            result.append(row)
        return result

    def apply_rows(self, rows_to_apply):
        if not rows_to_apply:
            self.show_message('No valid Apply rows are ready to move.', image=MessageBoxImage.Warning)
            return

        moved = 0
        skipped = 0
        failed = 0
        fail_rows = []

        t = DB.Transaction(doc, 'MENVIC - Workset Inspector Apply')
        t.Start()
        try:
            for row in rows_to_apply:
                target_ws = self.ws_by_name.get(row.SelectedDestination)
                if target_ws is None:
                    failed += len(row.ElementIds)
                    fail_rows.append(['<group>', row.CategoryName, row.TypeName, 'Missing destination workset'])
                    continue

                target_id = target_ws.Id.IntegerValue

                for elem_id in row.ElementIds:
                    try:
                        elem = doc.GetElement(elem_id)
                        if elem is None:
                            failed += 1
                            continue
                        param = get_workset_param(elem)
                        if param is None or param.IsReadOnly:
                            failed += 1
                            fail_rows.append([get_element_display_name(elem), row.CategoryName, row.TypeName, 'Workset parameter unavailable'])
                            continue
                        if elem.WorksetId.IntegerValue == target_id:
                            skipped += 1
                            continue
                        param.Set(target_id)
                        moved += 1
                    except Exception as ex:
                        failed += 1
                        fail_rows.append([get_element_display_name(elem), row.CategoryName, row.TypeName, first_line(ex)])
            t.Commit()
        except Exception:
            t.RollBack()
            raise

        self.status_text.Text = 'Moved: {} | Skipped: {} | Failed: {}'.format(moved, skipped, failed)
        self.show_message(
            'Apply complete.\n\nMoved: {}\nSkipped: {}\nFailed: {}'.format(moved, skipped, failed),
            image=MessageBoxImage.Information
        )

        if fail_rows:
            output.print_md('## Workset Inspector - Failures')
            output.print_table(
                table_data=fail_rows,
                columns=['Element', 'Category', 'Type', 'Reason']
            )
            self.activate_window()

        self.on_scan(None, None)

    def on_apply_all(self, sender, args):
        try:
            rows_to_apply = self.collect_rows_to_apply()
            self.apply_rows(rows_to_apply)
        except Exception as ex:
            self.show_message(first_line(ex), image=MessageBoxImage.Error)

    def on_close(self, sender, args):
        self.Close()


# ==================================================
# BOOTSTRAP
# ==================================================

if not is_workshared_document(doc):
    forms.alert(
        "This tool requires a workshared model with worksets enabled.\n\nEnable Worksharing first and run the tool again.",
        title="pyMENVIC | Worksets Required",
        warn_icon=True
    )
    raise SystemExit

script_dir = os.path.dirname(__file__)
xaml_path = os.path.join(script_dir, 'WorksetInspector.xaml')

if not os.path.exists(xaml_path):
    forms.alert('WorksetInspector.xaml was not found next to script.py.', title=TOOL_NAME, exitscript=True)

WINDOW = InspectorWindow(xaml_path)
WINDOW.ShowDialog()

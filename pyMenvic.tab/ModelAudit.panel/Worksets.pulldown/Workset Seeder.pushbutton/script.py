# -*- coding: utf-8 -*-

__title__ = "Workset Seeder"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
WORKSET SEEDER
_____________________________________________________

Description:

Create standard user worksets from a discipline-based office profile.
The tool supports SIMPLE and EXTENDED profiles, previews the target
worksets, and creates only the missing ones.

_____________________________________________________
What the tool does:

• lets the user choose discipline and profile level
• previews the worksets that will be created
• optionally includes coordination and link worksets
• creates only missing worksets safely and without duplicates

_____________________________________________________
Output:

Apply summary in pyRevit output and status feedback in the window.

_____________________________________________________
Usage:

1. Open the pyRevit button
2. Select discipline and profile
3. Review preview list and create missing worksets

_____________________________________________________

Author: Ricardo J. Mendieta
"""

from pyrevit import revit, DB, forms, script
from lib.core.branding import get_logo_path
from System.Collections.ObjectModel import ObservableCollection
from System import Uri, UriKind
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
import os


doc = revit.doc


def is_workshared_document(current_doc):
    try:
        return current_doc is not None and current_doc.IsWorkshared
    except Exception:
        return False
output = script.get_output()


# ==================================================
# CONFIG
# ==================================================

DISCIPLINES = [
    "ARCHITECTURE",
    "STRUCTURE",
    "MECHANICAL",
    "ELECTRICAL",
    "PLUMBING",
    "SITE"
]

PROFILE_LEVELS = [
    "SIMPLE",
    "EXTENDED"
]

COORDINATION_WORKSETS = [
    "SHARED_LEVELS_GRIDS",
    "COORD_MODEL",
    "COORD_REFERENCE"
]

LINK_WORKSETS = [
    "LINK_ARC",
    "LINK_STR",
    "LINK_MECH",
    "LINK_ELE",
    "LINK_PLM",
    "LINK_SITE",
    "LINK_CAD",
    "LINK_REF"
]

PROFILE_STANDARDS = {
    "SIMPLE": {
        "ARCHITECTURE": [
            "ARC_LEVELS_GRIDS",
            "ARC_MODEL",
            "ARC_INTERIORS",
            "ARC_FINISHES"
        ],
        "STRUCTURE": [
            "STR_LEVELS_GRIDS",
            "STR_MODEL",
            "STR_REBARS"
        ],
        "MECHANICAL": [
            "MECH_LEVELS_GRIDS",
            "MECH_MODEL",
            "MECH_EQUIPMENT",
            "MECH_DUCTS"
        ],
        "ELECTRICAL": [
            "ELE_LEVELS_GRIDS",
            "ELE_MODEL",
            "ELE_LIGHTING",
            "ELE_POWER",
            "ELE_LOW_CURRENT"
        ],
        "PLUMBING": [
            "PLM_LEVELS_GRIDS",
            "PLM_MODEL",
            "PLM_WATER",
            "PLM_DRAINAGE"
        ],
        "SITE": [
            "SITE_LEVELS_GRIDS",
            "SITE_MODEL",
            "SITE_TOPO"
        ]
    },
    "EXTENDED": {
        "ARCHITECTURE": [
            "ARC_LEVELS_GRIDS",
            "ARC_MODEL_CORE",
            "ARC_MODEL_FINISHES",
            "ARC_MODEL_INTERIORS"
        ],
        "STRUCTURE": [
            "STR_LEVELS_GRIDS",
            "STR_MODEL_CORE",
            "STR_MODEL_FRAMING",
            "STR_MODEL_REBARS"
        ],
        "MECHANICAL": [
            "MECH_LEVELS_GRIDS",
            "MECH_MODEL_CORE",
            "MECH_MODEL_DUCTS",
            "MECH_MODEL_EQUIPMENT"
        ],
        "ELECTRICAL": [
            "ELE_LEVELS_GRIDS",
            "ELE_MODEL_CORE",
            "ELE_MODEL_LIGHTING",
            "ELE_MODEL_POWER",
            "ELE_MODEL_LOW_CURRENT"
        ],
        "PLUMBING": [
            "PLM_LEVELS_GRIDS",
            "PLM_MODEL_CORE",
            "PLM_MODEL_WATER",
            "PLM_MODEL_DRAINAGE"
        ],
        "SITE": [
            "SITE_LEVELS_GRIDS",
            "SITE_MODEL",
            "SITE_MODEL_TOPO"
        ]
    }
}


# ==================================================
# HELPERS
# ==================================================


def safe_str(value):
    try:
        return str(value)
    except Exception:
        try:
            return unicode(value)
        except Exception:
            return ""



def safe_upper(value):
    if value is None:
        return ""
    try:
        return str(value).strip().upper()
    except Exception:
        try:
            return unicode(value).strip().upper()
        except Exception:
            return ""



def get_existing_user_worksets():
    worksets = []
    collector = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
    for ws in collector:
        worksets.append(ws)
    worksets.sort(key=lambda x: safe_upper(x.Name))
    return worksets



def get_existing_workset_name_set():
    names = set()
    for ws in get_existing_user_worksets():
        names.add(safe_upper(ws.Name))
    return names



def unique_append(target_list, name):
    name = safe_upper(name)
    if (not name) or name in target_list:
        return
    target_list.append(name)



def build_target_worksets(profile_level, discipline, include_coordination, include_links):
    result = []

    profile_level = safe_upper(profile_level)
    discipline = safe_upper(discipline)

    base_items = PROFILE_STANDARDS.get(profile_level, {}).get(discipline, [])
    for item in base_items:
        unique_append(result, item)

    if include_coordination:
        for item in COORDINATION_WORKSETS:
            unique_append(result, item)

    if include_links:
        for item in LINK_WORKSETS:
            unique_append(result, item)

    return result



def create_workset_if_missing(name):
    existing_names = get_existing_workset_name_set()
    target = safe_upper(name)

    if target in existing_names:
        return False, "Already exists"

    if hasattr(DB.WorksetTable, "IsWorksetNameUnique"):
        if not DB.WorksetTable.IsWorksetNameUnique(doc, target):
            return False, "Name is not unique"

    DB.Workset.Create(doc, target)
    return True, "Created"



def load_logo_if_available(window):
    try:
        bundle_dir = os.path.dirname(__file__)
        logo_path = get_logo_path()
        if not os.path.exists(logo_path):
            return

        bmp = BitmapImage()
        bmp.BeginInit()
        bmp.CacheOption = BitmapCacheOption.OnLoad
        bmp.UriSource = Uri(logo_path, UriKind.Absolute)
        bmp.EndInit()
        window.HeaderLogoImage.Source = bmp
    except Exception:
        pass


# ==================================================
# SCAN
# ==================================================


def get_preview_rows(profile_level, discipline, include_coordination, include_links):
    rows = []
    existing = get_existing_workset_name_set()

    for name in build_target_worksets(profile_level, discipline, include_coordination, include_links):
        status = "EXISTS" if name in existing else "MISSING"
        rows.append((name, status))

    return rows


# ==================================================
# DATA MODEL
# ==================================================


class PreviewRow(object):
    def __init__(self, name, status):
        self.Name = name
        self.Status = status
        self.Selected = status == "MISSING"
        self.IsSelectable = status == "MISSING"


# ==================================================
# PROCESS
# ==================================================


class WorksetSeederWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)

        self._rows = ObservableCollection[object]()

        load_logo_if_available(self)
        self.setup_controls()
        self.bind_events()
        self.refresh_preview()
        self.ShowDialog()

    def setup_controls(self):
        profile_items = ObservableCollection[object]()
        discipline_items = ObservableCollection[object]()

        for item in PROFILE_LEVELS:
            profile_items.Add(item)
        for item in DISCIPLINES:
            discipline_items.Add(item)

        self.ProfileCombo.ItemsSource = profile_items
        self.DisciplineCombo.ItemsSource = discipline_items
        self.PreviewList.ItemsSource = self._rows

        self.ProfileCombo.SelectedIndex = 0
        self.DisciplineCombo.SelectedIndex = 0
        self.IncludeCoordinationCheck.IsChecked = True
        self.IncludeLinksCheck.IsChecked = True

    def bind_events(self):
        self.ProfileCombo.SelectionChanged += self.on_options_changed
        self.DisciplineCombo.SelectionChanged += self.on_options_changed
        self.IncludeCoordinationCheck.Checked += self.on_options_changed
        self.IncludeCoordinationCheck.Unchecked += self.on_options_changed
        self.IncludeLinksCheck.Checked += self.on_options_changed
        self.IncludeLinksCheck.Unchecked += self.on_options_changed
        self.RefreshButton.Click += self.on_refresh
        self.CreateButton.Click += self.on_create
        self.CloseButton.Click += self.on_close

    def get_profile_level(self):
        try:
            return safe_upper(self.ProfileCombo.SelectedItem)
        except Exception:
            return "SIMPLE"

    def get_discipline(self):
        try:
            return safe_upper(self.DisciplineCombo.SelectedItem)
        except Exception:
            return "ARCHITECTURE"

    def include_coordination(self):
        try:
            return bool(self.IncludeCoordinationCheck.IsChecked)
        except Exception:
            return False

    def include_links(self):
        try:
            return bool(self.IncludeLinksCheck.IsChecked)
        except Exception:
            return False

    def update_status(self, text):
        self.StatusText.Text = text

    def refresh_preview(self):
        preview_rows = get_preview_rows(
            self.get_profile_level(),
            self.get_discipline(),
            self.include_coordination(),
            self.include_links()
        )

        self._rows.Clear()

        total_count = 0
        missing_count = 0
        existing_count = 0
        selected_count = 0

        for name, status in preview_rows:
            total_count += 1
            if status == "MISSING":
                missing_count += 1
            else:
                existing_count += 1
            row = PreviewRow(name, status)
            if row.Selected:
                selected_count += 1
            self._rows.Add(row)

        self.TotalCountText.Text = str(total_count)
        self.MissingCountText.Text = str(missing_count)
        self.ExistingCountText.Text = str(existing_count)
        self.update_status("Preview ready. Selected: {}".format(selected_count))

    def on_options_changed(self, sender, args):
        self.refresh_preview()

    def on_refresh(self, sender, args):
        self.refresh_preview()
        self.update_status("Preview refreshed.")

    def on_create(self, sender, args):
        targets = []
        for row in self._rows:
            if row.Selected and row.Status == "MISSING":
                targets.append(row.Name)

        if not targets:
            forms.alert("No missing worksets selected.", title="Workset Seeder")
            return

        confirm = forms.alert(
            "Create missing worksets for:\n\nProfile: {}\nDiscipline: {}\nInclude Coordination: {}\nInclude Links: {}".format(
                self.get_profile_level(),
                self.get_discipline(),
                "Yes" if self.include_coordination() else "No",
                "Yes" if self.include_links() else "No"
            ),
            yes=True,
            no=True
        )
        if not confirm:
            self.update_status("Creation canceled by user.")
            return

        created = []
        skipped = []
        failed = []

        t = DB.Transaction(doc, "pyMENVIC | Seed Worksets")
        t.Start()
        try:
            for name in targets:
                try:
                    was_created, message = create_workset_if_missing(name)
                    if was_created:
                        created.append(name)
                    else:
                        skipped.append("{} | {}".format(name, message))
                except Exception as ex:
                    failed.append("{} | {}".format(name, safe_str(ex).splitlines()[0]))
            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            forms.alert("Could not create worksets:\n{}".format(safe_str(ex).splitlines()[0]), title="Workset Seeder")
            self.update_status("Creation failed.")
            return

        self.refresh_preview()

        output.print_md("# pyMENVIC | WORKSET SEEDER — {}".format(self.get_profile_level()))
        output.print_md("")
        output.print_md("## Resumen")
        output.print_md("")
        output.print_md("- **Disciplina:** {}".format(self.get_discipline()))
        output.print_md("- **Coordinación incluida:** {}".format("Sí" if self.include_coordination() else "No"))
        output.print_md("- **Links incluidos:** {}".format("Sí" if self.include_links() else "No"))
        output.print_md("- **Creados:** {}".format(len(created)))
        output.print_md("- **Omitidos / ya existentes:** {}".format(len(skipped)))
        output.print_md("- **Fallos:** {}".format(len(failed)))
        output.print_md("")

        if created:
            output.print_md("## Creados")
            output.print_md("")
            for item in created:
                output.print_md("- {}".format(item))
            output.print_md("")

        if skipped:
            output.print_md("## Omitidos")
            output.print_md("")
            for item in skipped:
                output.print_md("- {}".format(item))
            output.print_md("")

        if failed:
            output.print_md("## Fallos")
            output.print_md("")
            for item in failed:
                output.print_md("- {}".format(item))
            output.print_md("")

        if not created and not failed:
            output.print_md("✅ Todo ya estaba estandarizado.")

        self.update_status("Creation completed.")
        forms.alert("Workset seeding completed.", title="Workset Seeder")

    def on_close(self, sender, args):
        self.Close()


# ==================================================
# CLEANUP
# ==================================================


if not is_workshared_document(doc):
    forms.alert(
        "This tool requires a workshared model with worksets enabled.\n\nEnable Worksharing first and run the tool again.",
        title="pyMENVIC | Worksets Required",
        warn_icon=True
    )
    raise SystemExit


# ==================================================
# REPORT
# ==================================================


xaml_path = os.path.join(os.path.dirname(__file__), "WorksetSeeder.xaml")
WorksetSeederWindow(xaml_path)

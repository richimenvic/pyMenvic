# -*- coding: utf-8 -*-
__title__ = "Export Selected Families"
__doc__ = "Tree selector by category/family, export by category, and report for categories outside the base list."

import os
import re
import shutil
import tempfile
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Thickness
from System.Windows.Controls import TreeViewItem, CheckBox

from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()
SCRIPT_DIR = os.path.dirname(__file__)
XAML_PATH = os.path.join(SCRIPT_DIR, "FamilySelector.xaml")


# =========================================================
# LISTA BASE DE CATEGORIAS
# =========================================================

VALID_CATEGORIES = set([

# CORE MODEL
"Abutments",
"Abutment Foundations",
"Abutment Piles",
"Abutment Walls",
"Approach Slabs",
"Air Terminals",
"Arches",
"Bearings",
"Bridge Cables",
"Bridge Decks",
"Bridge Framing",
"Balusters",
"Boundary Conditions",
"Cable Tray Fittings",
"Casework",
"Columns",
"Communication Devices",
"Conduit Fittings",
"Cross Bracing",
"Data Devices",
"Doors",
"Duct Accessories",
"Duct Fittings",
"Electrical Equipment",
"Electrical Fixtures",
"Entourage",
"Expansion Joints",
"Fire Alarm Devices",
"Fire Protection",
"Food Service Equipment",
"Furniture",
"Furniture Systems",
"Generic Models",
"Hardscape",
"Lighting Devices",
"Lighting Fixtures",
"Mass",
"Mechanical Control Devices",
"Mechanical Equipment",
"Medical Equipment",
"MEP Ancillary Framing",
"MEP Fabrication Ductwork Stiffeners",
"Parking",
"Pipe Accessories",
"Pipe Fittings",
"Planting",
"Plumbing Equipment",
"Plumbing Fixtures",
"Railings",
"Roads",
"Security Devices",
"Signage",
"Site",
"Specialty Equipment",
"Sprinklers",
"Structural Columns",
"Structural Connections",
"Structural Foundations",
"Structural Framing",
"Structural Stiffeners",
"Structural Tendons",
"Structural Trusses",
"Telephone Devices",
"Temporary Structures",
"Vertical Circulation",
"Vibration Dampers",
"Vibration Isolators",
"Vibration Management",
"Windows",

# ANNOTATIONS
"Annotation Symbols",
"Detail Items",
"Generic Annotations",
"Text Notes",
"Dimensions",
"Revision Clouds",
"Stairs Paths",

# VIEW / MARKERS
"Section Marks",
"Section Heads",
"Section Tails",
"Callout Heads",
"Level Heads",
"Grid Heads",
"Elevation Marks",
"View Titles",
"View Reference",

# TITLEBLOCKS
"Title Blocks",

# SPOT SYMBOLS
"Spot Elevation Symbols",
"Spot Coordinate Symbols",
"Spot Slope Symbols",

# TAGS
"Area Tags",
"Assembly Tags",
"Casework Tags",
"Curtain Wall Mullion Tags",
"Detail Item Tags",
"Door Tags",
"Window Tags",
"Wall Tags",
"Room Tags",
"Space Tags",
"Material Tags",
"Multi-Category Tags",
"Structural Framing Tags",
"Structural Column Tags",
"Structural Foundation Tags",
"Structural Beam System Tags",
"Structural Truss Tags",
"Structural Rebar Tags",
"Structural Rebar Coupler Tags",
"Structural Rebar Couplers",
"Structural Area Reinforcement Tags",
"Structural Path Reinforcement Tags",
"Structural Fabric Reinforcement Tags",
"Plumbing Fixture Tags",
"Plumbing Equipment Tags",
"Mechanical Equipment Tags",
"Electrical Equipment Tags",
"Electrical Fixture Tags",
"Lighting Device Tags",
"Lighting Fixture Tags",
"Security Device Tags",
"Specialty Equipment Tags",
"Telephone Device Tags",
"Nurse Call Device Tags",
"Furniture Tags",
"Furniture System Tags",
"Pipe Tags",
"Pipe Fitting Tags",
"Pipe Accessory Tags",
"Pipe Insulation Tags",
"Duct Tags",
"Duct Fitting Tags",
"Duct Accessory Tags",
"Duct Insulation Tags",
"Flex Duct Tags",
"Flex Pipe Tags",
"Railing Tags",
"Plate Tags",
"Part Tags",
"Keynote Tags",
"Revision Cloud Tags",
"Stair Tags",
"Path of Travel Tags",
"Property Line Segment Tags",
"Zone Tags",
"Mass Tags",
"Mass Floor Tags",
"Toposolid Tags",
"Toposolid Link Tags",

# STRUCTURAL SYMBOLS
"Brace in Plan View Symbols",
"Span Direction Symbol",
"Span Direction Symbols",
"Structural Area Reinforcement Symbols",
"Structural Path Reinforcement Symbols",
"Structural Fabric Reinforcement Symbols",

# OTHER
"Division Profiles",
"Profiles",
"Profile Families",
"Rebar Shape",
"Weld Tags",
"Wire Tick Marks",
"Air Terminal Tags",
"Anchor Tags",
"Area Based Load Tags",
"Ceiling Tags",
"Connection Symbols",
"Curtain Panel Tags",
"Fire Alarm Device Tags",
"Floor Tags",
"Generic Model Tags",
"MEP Fabrication Pipework Tags",
"Parking Tags",
"Shear Stud Tags",
"Supports",
"Terminations",
])

DISCOVERED_EXTRA_CATEGORIES = set()

INVALID_RE = re.compile(r'[<>:"/\\|?*]+')
VERSION_RE = re.compile(r'([ _-]?V[ _-]?\d{2})$', re.IGNORECASE)


# =========================================================
# HELPERS
# =========================================================

def sanitize_file_name(text):
    if not text:
        return "Unnamed"

    text = INVALID_RE.sub("_", text)
    text = text.replace("\n", " ").replace("\r", " ").strip()
    text = re.sub(r"\s+", " ", text)
    text = text.strip(" .")

    return text if text else "Unnamed"


def sanitize_folder_name(text):
    if not text:
        return "ZZ_Unclassified"

    text = INVALID_RE.sub("_", text)
    text = text.replace("\n", " ").replace("\r", " ").strip()
    text = re.sub(r"\s+", "_", text)
    text = text.strip(" .")

    return text if text else "ZZ_Unclassified"


def strip_version(name):
    if not name:
        return "Unnamed"

    base = name.strip()
    while True:
        new_base = VERSION_RE.sub("", base).strip()
        if new_base == base:
            break
        base = new_base

    return base


def get_version_suffix():
    try:
        ver = doc.Application.VersionNumber
    except:
        ver = ""

    if ver == "2024":
        return "_V24"
    elif ver == "2025":
        return "_V25"
    elif ver.isdigit() and len(ver) >= 4:
        return "_V{}".format(ver[-2:])
    else:
        return "_VXX"


def normalize_code(name):
    base = strip_version(name)
    base = sanitize_file_name(base)
    base = base.replace(" ", "").replace("_", "").replace("-", "")
    return base.lower()



def get_element_id_value(element_id):
    """Return a stable integer value for Revit ElementId across API versions."""
    if element_id is None:
        return None

    try:
        return int(element_id.Value)
    except:
        pass

    try:
        return int(element_id.IntegerValue)
    except:
        pass

    try:
        return int(str(element_id))
    except:
        return None


def ensure_folder(path):
    if not os.path.exists(path):
        os.makedirs(path)


def get_family_category_name(family):
    try:
        if family and family.FamilyCategory:
            cname = family.FamilyCategory.Name
            if cname and cname not in VALID_CATEGORIES:
                DISCOVERED_EXTRA_CATEGORIES.add(cname)
            return cname if cname else "ZZ_Unclassified"
    except:
        pass

    return "ZZ_Unclassified"


def find_existing_file(folder, family_name):
    if not os.path.exists(folder):
        return None

    target = normalize_code(family_name)

    for fname in os.listdir(folder):
        fullpath = os.path.join(folder, fname)
        if not os.path.isfile(fullpath):
            continue

        root, ext = os.path.splitext(fname)
        if ext.lower() != ".rfa":
            continue

        if normalize_code(root) == target:
            return fullpath

    return None


# =========================================================
# COLLECTION
# =========================================================

def get_editable_families():
    fams = []
    collector = DB.FilteredElementCollector(doc).OfClass(DB.Family)

    for fam in collector:
        try:
            if fam is None:
                continue
            if fam.IsInPlace:
                continue
            if not fam.IsEditable:
                continue
            fams.append(fam)
        except:
            pass

    fams = sorted(fams, key=lambda x: (get_family_category_name(x), x.Name.lower()))
    return fams


def build_category_map(families):
    catmap = {}
    for fam in families:
        cat = get_family_category_name(fam)
        if cat not in catmap:
            catmap[cat] = []
        catmap[cat].append(fam)

    for cat in catmap.keys():
        catmap[cat] = sorted(catmap[cat], key=lambda x: x.Name.lower())

    return catmap


# =========================================================
# WPF TREE SELECTOR
# =========================================================

# XAML UI is stored externally as FamilySelector.xaml in the same bundle folder.
class FamilyTreeWindow(forms.WPFWindow):
    def __init__(self, xaml_source, category_map):
        forms.WPFWindow.__init__(self, xaml_source)

        self.category_map = category_map
        self.family_by_id = {}
        self.checked_ids = set()
        self.result = None
        self._updating = False

        for cat_name, fams in self.category_map.items():
            for fam in fams:
                self.family_by_id[get_element_id_value(fam.Id)] = fam

        self.tbSearch.TextChanged += self.on_search_changed
        self.btnCheckAll.Click += self.on_check_all
        self.btnUncheckAll.Click += self.on_uncheck_all
        self.btnExpandAll.Click += self.on_expand_all
        self.btnCollapseAll.Click += self.on_collapse_all
        self.btnAccept.Click += self.on_accept

        self.build_tree("")

    def build_tree(self, filter_text):
        self._updating = True
        self.tvFamilies.Items.Clear()

        filter_text = (filter_text or "").strip().lower()

        ordered_categories = sorted(self.category_map.keys(), key=lambda x: x.lower())

        for cat_name in ordered_categories:
            all_fams = self.category_map[cat_name]

            visible_fams = []
            for fam in all_fams:
                if not filter_text:
                    visible_fams.append(fam)
                else:
                    if filter_text in cat_name.lower() or filter_text in fam.Name.lower():
                        visible_fams.append(fam)

            if not visible_fams:
                continue

            cat_item = TreeViewItem()
            cat_item.IsExpanded = True if filter_text else False

            cat_cb = CheckBox()
            cat_cb.Content = u"{} ({})".format(cat_name, len(all_fams))
            cat_cb.Tag = ("category", cat_name)
            cat_cb.IsThreeState = True
            cat_cb.Margin = Thickness(1)
            cat_cb.Checked += self.on_category_checked
            cat_cb.Unchecked += self.on_category_unchecked

            cat_item.Header = cat_cb

            for fam in visible_fams:
                fam_item = TreeViewItem()

                fam_cb = CheckBox()
                fam_cb.Content = fam.Name
                fam_cb.Tag = ("family", get_element_id_value(fam.Id), cat_name)
                fam_cb.Margin = Thickness(1)

                if get_element_id_value(fam.Id) in self.checked_ids:
                    fam_cb.IsChecked = True
                else:
                    fam_cb.IsChecked = False

                fam_cb.Checked += self.on_family_toggled
                fam_cb.Unchecked += self.on_family_toggled

                fam_item.Header = fam_cb
                cat_item.Items.Add(fam_item)

            self.tvFamilies.Items.Add(cat_item)
            self.update_category_checkbox(cat_item)

        self._updating = False
        self.update_status()

    def update_category_checkbox(self, cat_item):
        cat_cb = cat_item.Header
        tag = cat_cb.Tag
        cat_name = tag[1]

        fam_ids = [get_element_id_value(f.Id) for f in self.category_map.get(cat_name, [])]
        checked_count = sum(1 for fid in fam_ids if fid in self.checked_ids)

        self._updating = True
        if checked_count == 0:
            cat_cb.IsChecked = False
        elif checked_count == len(fam_ids):
            cat_cb.IsChecked = True
        else:
            cat_cb.IsChecked = None
        self._updating = False

    def update_all_category_states(self):
        for item in self.tvFamilies.Items:
            self.update_category_checkbox(item)

    def update_status(self):
        total = len(self.checked_ids)
        if total == 1:
            self.txtStatus.Text = u"1 family selected."
        else:
            self.txtStatus.Text = u"{} families selected.".format(total)

    def on_search_changed(self, sender, args):
        self.build_tree(self.tbSearch.Text)

    def on_category_checked(self, sender, args):
        if self._updating:
            return

        tag = sender.Tag
        cat_name = tag[1]

        for fam in self.category_map.get(cat_name, []):
            self.checked_ids.add(get_element_id_value(fam.Id))

        self.build_tree(self.tbSearch.Text)

    def on_category_unchecked(self, sender, args):
        if self._updating:
            return

        tag = sender.Tag
        cat_name = tag[1]

        for fam in self.category_map.get(cat_name, []):
            fid = get_element_id_value(fam.Id)
            if fid in self.checked_ids:
                self.checked_ids.remove(fid)

        self.build_tree(self.tbSearch.Text)

    def on_family_toggled(self, sender, args):
        if self._updating:
            return

        tag = sender.Tag
        family_id = tag[1]

        if sender.IsChecked:
            self.checked_ids.add(family_id)
        else:
            if family_id in self.checked_ids:
                self.checked_ids.remove(family_id)

        self.update_all_category_states()
        self.update_status()

    def on_check_all(self, sender, args):
        for fid in self.family_by_id.keys():
            self.checked_ids.add(fid)
        self.build_tree(self.tbSearch.Text)

    def on_uncheck_all(self, sender, args):
        self.checked_ids = set()
        self.build_tree(self.tbSearch.Text)

    def on_expand_all(self, sender, args):
        for item in self.tvFamilies.Items:
            item.IsExpanded = True

    def on_collapse_all(self, sender, args):
        for item in self.tvFamilies.Items:
            item.IsExpanded = False

    def on_accept(self, sender, args):
        selected = []
        for fid in sorted(self.checked_ids):
            fam = self.family_by_id.get(fid)
            if fam:
                selected.append(fam)

        self.result = selected
        self.Close()



# =========================================================
# EXPORT
# =========================================================

def export_family(family, root_folder, version_suffix, stats):
    fam_doc = None
    temp_dir = None

    try:
        base_name = strip_version(family.Name)
        final_name = sanitize_file_name(base_name) + version_suffix

        category_name = sanitize_folder_name(get_family_category_name(family))
        category_folder = os.path.join(root_folder, category_name)
        ensure_folder(category_folder)

        target_path = os.path.join(category_folder, final_name + ".rfa")
        existing_path = find_existing_file(category_folder, base_name)

        fam_doc = doc.EditFamily(family)

        temp_dir = tempfile.mkdtemp(prefix="pyMenvic_")
        temp_path = os.path.join(temp_dir, final_name + ".rfa")

        save_opts = DB.SaveAsOptions()
        save_opts.OverwriteExistingFile = True
        fam_doc.SaveAs(temp_path, save_opts)

        if existing_path:
            existing_norm = os.path.normcase(os.path.normpath(existing_path))
            target_norm = os.path.normcase(os.path.normpath(target_path))

            if existing_norm != target_norm and os.path.exists(target_path):
                os.remove(target_path)

            shutil.copy2(temp_path, target_path)

            if existing_norm != target_norm and os.path.exists(existing_path):
                os.remove(existing_path)

            stats["updated"] += 1
            output.print_md(u"Updated: **{}**".format(family.Name))
        else:
            shutil.copy2(temp_path, target_path)
            stats["created"] += 1
            output.print_md(u"Created: **{}**".format(family.Name))

    except Exception as ex:
        stats["failed"] += 1
        output.print_md(u"Error with **{}**: `{}`".format(family.Name, str(ex)))
    finally:
        if fam_doc:
            try:
                fam_doc.Close(False)
            except:
                pass

        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except:
                pass


def print_missing_categories_report():
    if DISCOVERED_EXTRA_CATEGORIES:
        output.print_md("## Categories found outside the base list")
        for cname in sorted(DISCOVERED_EXTRA_CATEGORIES):
            output.print_md("- {}".format(cname))
    else:
        output.print_md("## No categories outside the base list were detected")


# =========================================================
# MAIN
# =========================================================

def main():
    families = get_editable_families()

    if not families:
        forms.alert("No editable/exportable families were found in this project.")
        return

    category_map = build_category_map(families)

    if not os.path.exists(XAML_PATH):
        forms.alert("FamilySelector.xaml was not found in the ExportFamilies.pushbutton folder.")
        return

    selector = FamilyTreeWindow(XAML_PATH, category_map)
    selector.ShowDialog()
    selected_families = selector.result

    if not selected_families:
        return

    folder = forms.pick_folder(title="Select destination folder to export families")
    if not folder:
        return

    confirm = forms.alert(
        "{} families will be exported.\n\nContinue?".format(len(selected_families)),
        yes=True,
        no=True
    )

    if not confirm:
        return

    version_suffix = get_version_suffix()
    stats = {"created": 0, "updated": 0, "failed": 0}

    for fam in selected_families:
        export_family(fam, folder, version_suffix, stats)

    print_missing_categories_report()

    extra_count = len(DISCOVERED_EXTRA_CATEGORIES)
    extra_msg = ""
    if extra_count > 0:
        extra_msg = "\n\nCategories outside the base list detected: {}".format(extra_count)

    forms.alert(
        "Process completed.\n\nCreated: {}\nUpdated: {}\nErrors: {}{}".format(
            stats["created"], stats["updated"], stats["failed"], extra_msg
        )
    )


if __name__ == "__main__":
    main()
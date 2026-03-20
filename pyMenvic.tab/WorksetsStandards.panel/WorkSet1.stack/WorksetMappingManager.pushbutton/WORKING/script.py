# -*- coding: utf-8 -*-

__title__ = "Workset Mapping Manager"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
WORKSET MAPPING MANAGER
_____________________________________________________

Description:

Maps non-standard user worksets into pyMENVIC standard
worksets, previews the proposed mapping, applies
the reassignment safely, and allows deleting empty
non-standard worksets.

_____________________________________________________
What the tool does:

• suggests standard worksets by discipline
• analyzes current user worksets
• maps elements from source worksets into target worksets
• deletes empty non-standard worksets

_____________________________________________________
Output:

Preview report and apply summary in pyRevit output.

_____________________________________________________
Usage:

1. Open the tool
2. Select discipline and review mapping
3. Preview or apply changes

_____________________________________________________

Author: Ricardo J. Mendieta
"""

from pyrevit import revit, DB, forms, script
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from System.Windows import MessageBox
from System.Windows.Data import CollectionViewSource
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind
import os

doc = revit.doc
output = script.get_output()

# ==================================================
# CONFIG
# ==================================================

PLACEHOLDER_TARGET = "-- Select --"

DISCIPLINES = [
    "Architecture",
    "Structure",
    "Mechanical",
    "Electrical",
    "Plumbing",
    "Site",
    "Coordination",
    "All Disciplines"
]

DISCIPLINE_DISPLAY_TO_INTERNAL = {
    "ARCHITECTURE": "Architecture",
    "STRUCTURE": "Structure",
    "MECHANICAL": "Mechanical",
    "ELECTRICAL": "Electrical",
    "PLUMBING": "Plumbing",
    "SITE": "Site",
    "COORDINATION": "Coordination",
    "ALL DISCIPLINES": "All Disciplines"
}

DISCIPLINE_PREFIX = {
    "Architecture": "ARC",
    "Structure": "STR",
    "Mechanical": "MECH",
    "Electrical": "ELE",
    "Plumbing": "PLM",
    "Site": "SITE"
}

ALWAYS_CORE = [
    "ARC_MODEL",
    "ARC_LEVELS_GRIDS"
]

ALL_CORE = [
    "ARC_MODEL",
    "ARC_LEVELS_GRIDS",
    "STR_MODEL",
    "STR_LEVELS_GRIDS",
    "MECH_MODEL",
    "MECH_LEVELS_GRIDS",
    "ELE_MODEL",
    "ELE_LEVELS_GRIDS",
    "PLM_MODEL",
    "PLM_LEVELS_GRIDS",
    "SITE_MODEL",
    "SITE_LEVELS_GRIDS"
]

ALL_LINKS = [
    "LINK_ARC",
    "LINK_STR",
    "LINK_MECH",
    "LINK_ELE",
    "LINK_PLM",
    "LINK_SITE",
    "LINK_CAD",
    "LINK_REF"
]

ALL_TARGETS = [PLACEHOLDER_TARGET] + ALL_CORE + ALL_LINKS
ALL_ACTIONS = ["ASSIGN", "KEEP", "IGNORE", "REVIEW"]

STANDARD_CHECKBOX_TO_WORKSET = {
    "chkARC_MODEL": "ARC_MODEL",
    "chkARC_LEVELS_GRIDS": "ARC_LEVELS_GRIDS",
    "chkSTR_MODEL": "STR_MODEL",
    "chkSTR_LEVELS_GRIDS": "STR_LEVELS_GRIDS",
    "chkMECH_MODEL": "MECH_MODEL",
    "chkMECH_LEVELS_GRIDS": "MECH_LEVELS_GRIDS",
    "chkELE_MODEL": "ELE_MODEL",
    "chkELE_LEVELS_GRIDS": "ELE_LEVELS_GRIDS",
    "chkPLM_MODEL": "PLM_MODEL",
    "chkPLM_LEVELS_GRIDS": "PLM_LEVELS_GRIDS",
    "chkSITE_MODEL": "SITE_MODEL",
    "chkSITE_LEVELS_GRIDS": "SITE_LEVELS_GRIDS",
    "chkLINK_ARC": "LINK_ARC",
    "chkLINK_STR": "LINK_STR",
    "chkLINK_MECH": "LINK_MECH",
    "chkLINK_ELE": "LINK_ELE",
    "chkLINK_PLM": "LINK_PLM",
    "chkLINK_SITE": "LINK_SITE",
    "chkLINK_CAD": "LINK_CAD",
    "chkLINK_REF": "LINK_REF"
}

# ==================================================
# HELPERS
# ==================================================

def safe_str(ex):
    try:
        return str(ex)
    except Exception:
        return "Unknown error"

def normalize_name(name):
    try:
        return name.strip().upper()
    except Exception:
        return ""

def tokenize_name(name):
    n = normalize_name(name)
    n = n.replace("-", "_").replace(" ", "_")
    return [p for p in n.split("_") if p]

def starts_with_any(text, prefixes):
    for prefix in prefixes:
        if text.startswith(prefix):
            return True
    return False

def is_target_empty(value):
    return (not value) or value == PLACEHOLDER_TARGET

def get_workset_table_by_name():
    result = {}
    collector = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
    for ws in collector:
        result[ws.Name] = ws
    return result

def create_workset_if_missing(name):
    ws_map = get_workset_table_by_name()
    if name in ws_map:
        return False
    DB.Workset.Create(doc, name)
    return True

def build_element_count_by_workset():
    counts = {}
    fec = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()

    for el in fec:
        try:
            p = el.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
            if p and p.HasValue:
                wsid = p.AsInteger()
                counts[wsid] = counts.get(wsid, 0) + 1
        except Exception:
            pass

    return counts

def get_true_workset_element_count(workset_id):
    try:
        return DB.FilteredElementCollector(doc) \
                 .WherePasses(DB.ElementWorksetFilter(workset_id, False)) \
                 .GetElementCount()
    except Exception:
        count = 0
        for el in DB.FilteredElementCollector(doc).WhereElementIsNotElementType():
            try:
                p = el.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                if p and p.HasValue and p.AsInteger() == workset_id.IntegerValue:
                    count += 1
            except Exception:
                pass
        return count

def get_reassignable_elements_by_workset_id(source_wsid):
    result = []
    fec = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()

    for el in fec:
        try:
            p = el.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
            if not p:
                continue
            if not p.HasValue:
                continue
            if p.AsInteger() != source_wsid:
                continue
            result.append(el)
        except Exception:
            pass

    return result

def checkout_worksets_if_possible(workset_ids):
    checked_out = []
    failed = []

    if not workset_ids:
        return checked_out, failed

    ids = List[DB.WorksetId]()
    seen = set()

    for wsid in workset_ids:
        try:
            key = wsid.IntegerValue
        except Exception:
            key = str(wsid)

        if key in seen:
            continue
        seen.add(key)

        try:
            ids.Add(wsid)
        except Exception as ex:
            failed.append((wsid, "Could not add workset id: {}".format(safe_str(ex))))

    if ids.Count == 0:
        return checked_out, failed

    try:
        checked_collection = DB.WorksharingUtils.CheckoutWorksets(doc, ids)
        try:
            for wsid in checked_collection:
                checked_out.append(wsid)
        except Exception:
            pass
    except Exception as ex:
        reason = safe_str(ex)
        for wsid in ids:
            failed.append((wsid, reason))

    return checked_out, failed

def get_base_model_target_for_discipline(discipline):
    if discipline in DISCIPLINE_PREFIX:
        return "{}_MODEL".format(DISCIPLINE_PREFIX[discipline])
    return None

def rename_workset_safe(workset, new_name):
    if not workset or not new_name:
        return False, "Invalid rename input"

    try:
        if workset.Name == new_name:
            return True, "Already named"
    except Exception:
        pass

    try:
        if hasattr(DB.WorksetTable, "IsWorksetNameUnique"):
            if not DB.WorksetTable.IsWorksetNameUnique(doc, new_name):
                return False, "Target workset name already exists"
    except Exception:
        pass

    try:
        if hasattr(DB.WorksetTable, "RenameWorkset"):
            DB.WorksetTable.RenameWorkset(doc, workset.Id, new_name)
            return True, ""
        return False, "RenameWorkset API not available"
    except Exception as ex:
        return False, safe_str(ex)

def move_elements_to_workset(source_wsid, target_wsid):
    changed = 0
    skipped = 0
    failed = []

    for el in get_reassignable_elements_by_workset_id(source_wsid):
        try:
            p = el.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
            if not p:
                skipped += 1
                continue
            p.Set(target_wsid)
            changed += 1
        except Exception as ex:
            failed.append((el.Id.IntegerValue, safe_str(ex)))

    return changed, skipped, failed

def consolidate_workset_into_existing(source_name, target_name):
    changed = 0
    skipped = 0
    failed = []
    deleted = False
    delete_reason = ""

    ws_map = get_workset_table_by_name()
    source_ws = ws_map.get(source_name)
    target_ws = ws_map.get(target_name)

    if not source_ws:
        return changed, skipped, failed, deleted, "Source workset not found"
    if not target_ws:
        return changed, skipped, failed, deleted, "Target workset not found"

    source_id = source_ws.Id
    target_id = target_ws.Id

    changed, skipped, failed = move_elements_to_workset(source_id.IntegerValue, target_id.IntegerValue)

    remaining = get_true_workset_element_count(source_id)
    if remaining <= 0:
        try:
            if can_delete_workset_safe(source_id):
                delete_workset_safe(source_id)
                deleted = True
            else:
                move_settings = build_delete_move_settings(target_id)
                if move_settings and can_delete_workset_safe(source_id, move_settings):
                    delete_workset_safe(source_id, move_settings)
                    deleted = True
                else:
                    delete_reason = "Revit did not allow deleting consolidated source workset"
        except Exception as ex:
            delete_reason = safe_str(ex)
    else:
        move_settings = build_delete_move_settings(target_id)
        if move_settings and can_delete_workset_safe(source_id, move_settings):
            try:
                delete_workset_safe(source_id, move_settings)
                deleted = True
            except Exception as ex:
                delete_reason = safe_str(ex)
        else:
            delete_reason = "Source workset still contains {} element(s)".format(remaining)

    return changed, skipped, failed, deleted, delete_reason

def handle_workset1_rename_consolidation(discipline):
    result = {
        "applied": False,
        "changed": 0,
        "skipped": 0,
        "deleted": False,
        "renamed": False,
        "errors": [],
        "messages": []
    }

    target_name = get_base_model_target_for_discipline(discipline)
    if not target_name:
        return result

    ws_map = get_workset_table_by_name()
    workset1 = ws_map.get("Workset1") or ws_map.get("WORKSET1") or ws_map.get("Workset 1")
    target_ws = ws_map.get(target_name)

    if not workset1:
        result["messages"].append("Workset1 not found")
        return result

    if workset1.Name == target_name:
        result["messages"].append("Workset1 already named {}".format(target_name))
        return result

    ids_to_checkout = [workset1.Id]
    if target_ws:
        ids_to_checkout.append(target_ws.Id)

    checked_out, failed_checkout = checkout_worksets_if_possible(ids_to_checkout)
    if failed_checkout:
        for _, reason in failed_checkout:
            result["errors"].append("Checkout failed: {}".format(reason))

    t = DB.Transaction(doc, "pyMENVIC | Rename / Consolidate Workset1")
    t.Start()
    try:
        ws_map = get_workset_table_by_name()
        workset1 = ws_map.get("Workset1") or ws_map.get("WORKSET1") or ws_map.get("Workset 1")
        target_ws = ws_map.get(target_name)

        if not workset1:
            result["errors"].append("Workset1 not found during transaction")
            t.RollBack()
            return result

        if target_ws:
            ch, sk, failures, deleted, delete_reason = consolidate_workset_into_existing(target_name, workset1.Name)
            result["changed"] += ch
            result["skipped"] += sk

            if failures:
                for elid, reason in failures[:50]:
                    result["errors"].append("Move to Workset1 failed for {}: {}".format(elid, reason))

            if deleted:
                result["deleted"] = True
                result["messages"].append("{} consolidated into Workset1".format(target_name))
            elif delete_reason:
                result["messages"].append("Consolidation note: {}".format(delete_reason))

            ws_map = get_workset_table_by_name()
            workset1 = ws_map.get("Workset1") or ws_map.get("WORKSET1") or ws_map.get("Workset 1")

        ok, rename_reason = rename_workset_safe(workset1, target_name)
        if ok:
            result["renamed"] = True
            result["applied"] = True
            result["messages"].append("Workset1 renamed to {}".format(target_name))
        else:
            result["errors"].append("Rename failed: {}".format(rename_reason))

        t.Commit()
    except Exception as ex:
        t.RollBack()
        result["errors"].append("Rename/consolidate transaction failed: {}".format(safe_str(ex)))

    return result

def api_supports_delete_workset():
    return hasattr(DB, "WorksetTable") and hasattr(DB.WorksetTable, "DeleteWorkset")

def can_delete_workset_safe(workset_id, delete_settings=None):
    try:
        if not hasattr(DB.WorksetTable, "CanDeleteWorkset"):
            return True

        if delete_settings is not None:
            try:
                return DB.WorksetTable.CanDeleteWorkset(doc, workset_id, delete_settings)
            except Exception:
                pass

        try:
            return DB.WorksetTable.CanDeleteWorkset(doc, workset_id)
        except Exception:
            pass

        if hasattr(DB, "DeleteWorksetSettings"):
            try:
                settings = DB.DeleteWorksetSettings()
                return DB.WorksetTable.CanDeleteWorkset(doc, workset_id, settings)
            except Exception:
                pass

        return True
    except Exception:
        return False

def delete_workset_safe(workset_id, delete_settings=None):
    if not api_supports_delete_workset():
        raise Exception("DeleteWorkset API not available")

    if delete_settings is not None:
        try:
            return DB.WorksetTable.DeleteWorkset(doc, workset_id, delete_settings)
        except Exception:
            pass

    try:
        return DB.WorksetTable.DeleteWorkset(doc, workset_id)
    except Exception:
        pass

    if hasattr(DB, "DeleteWorksetSettings"):
        try:
            settings = DB.DeleteWorksetSettings()
            return DB.WorksetTable.DeleteWorkset(doc, workset_id, settings)
        except Exception:
            pass

    raise Exception("No compatible DeleteWorkset overload found")

def build_delete_move_settings(target_workset_id):
    if not hasattr(DB, "DeleteWorksetSettings"):
        return None
    if not hasattr(DB, "DeleteWorksetOption"):
        return None

    try:
        return DB.DeleteWorksetSettings(
            DB.DeleteWorksetOption.MoveElementsToWorkset,
            target_workset_id
        )
    except Exception:
        return None

# ==================================================
# CLASSIFICATION
# ==================================================

def detect_type(name):
    n = normalize_name(name)

    if n in ["WORKSET1", "WORKSET 1"]:
        return "BASE_MODEL"
    if n == "SHARED LEVELS AND GRIDS":
        return "BASE_LEVELS_GRIDS"

    if "MEP" in n:
        if "LINK" in n or "LINKED FILES" in n or "LINKED REVIT FILES" in n:
            return "MEP_LINK_LEGACY"
        if "LEVEL" in n or "GRID" in n:
            return "MEP_LEVELS_GRIDS_LEGACY"
        return "MEP_MODEL_LEGACY"

    if "LINK" in n or "LINKED FILES" in n or "LINKED REVIT FILES" in n:
        return "LINK"

    if "LEVEL" in n or "GRID" in n:
        prefix = detect_prefix(name)
        if prefix:
            return prefix + "_LEVELS_GRIDS"
        return "LEVELS_GRIDS_GENERIC"

    prefix = detect_prefix(name)
    if prefix:
        return prefix + "_MODEL"

    return "UNKNOWN"

def detect_prefix(name):
    n = normalize_name(name)

    if starts_with_any(n, ["ARC_", "ARQ_", "AR_", "A_"]) or "ARQ" in n or "ARQUITECT" in n:
        return "ARC"

    if starts_with_any(n, ["STR_", "ST_", "S_"]) or "STRUCT" in n or "ESTRUCT" in n:
        return "STR"

    if starts_with_any(n, ["MECH_"]) or "MECHANICAL" in n:
        return "MECH"

    if starts_with_any(n, ["ELE_", "EL_"]) or "ELECTR" in n:
        return "ELE"

    if starts_with_any(n, ["PLM_", "PL_", "P_"]) or "PLUMB" in n:
        return "PLM"

    if starts_with_any(n, ["SITE_"]) or "TOPO" in n or "SITE" in n:
        return "SITE"

    return None

def map_link_target(name):
    n = normalize_name(name)

    if "CAD" in n:
        return "LINK_CAD"
    if "MEP" in n:
        return PLACEHOLDER_TARGET
    if "AR" in n or "ARC" in n or "ARQ" in n:
        return "LINK_ARC"
    if "ST" in n or "STR" in n:
        return "LINK_STR"
    if "MECH" in n or "MECHANICAL" in n:
        return "LINK_MECH"
    if "EL" in n or "ELE" in n or "ELECT" in n:
        return "LINK_ELE"
    if "PL" in n or "PLM" in n or "PLUMB" in n:
        return "LINK_PLM"
    if "SITE" in n or "TOPO" in n:
        return "LINK_SITE"

    return "LINK_REF"

def suggest_target_and_action(name, discipline):
    n = normalize_name(name)

    if n in ALL_CORE or n in ALL_LINKS:
        return n, "KEEP", "Already standard"

    if n in ["WORKSET1", "WORKSET 1"]:
        if discipline in DISCIPLINE_PREFIX:
            target = "{}_MODEL".format(DISCIPLINE_PREFIX[discipline])
            return target, "ASSIGN", "Base model workset mapped by selected discipline"
        return PLACEHOLDER_TARGET, "REVIEW", "Base model workset is ambiguous in this mode"

    if n == "SHARED LEVELS AND GRIDS":
        if discipline in DISCIPLINE_PREFIX:
            target = "{}_LEVELS_GRIDS".format(DISCIPLINE_PREFIX[discipline])
            return target, "ASSIGN", "Base levels/grids workset mapped by selected discipline"
        return PLACEHOLDER_TARGET, "REVIEW", "Base levels/grids workset is ambiguous in this mode"

    if "MEP" in n:
        return PLACEHOLDER_TARGET, "REVIEW", "Legacy MEP naming detected; review manually"

    if "LINK" in n or "LINKED FILES" in n or "LINKED REVIT FILES" in n:
        link_target = map_link_target(name)
        if is_target_empty(link_target):
            return PLACEHOLDER_TARGET, "REVIEW", "Legacy or ambiguous link naming detected"
        return link_target, "ASSIGN", "Link-related workset detected"

    detected = detect_prefix(name)
    if detected:
        if "LEVEL" in n or "GRID" in n:
            return "{}_LEVELS_GRIDS".format(detected), "ASSIGN", "Discipline levels/grids workset detected"
        return "{}_MODEL".format(detected), "ASSIGN", "Discipline model workset detected"

    if discipline == "All Disciplines":
        return PLACEHOLDER_TARGET, "REVIEW", "No reliable automatic target in All Disciplines mode"

    if discipline == "Coordination":
        return PLACEHOLDER_TARGET, "REVIEW", "No reliable automatic target in Coordination mode"

    return PLACEHOLDER_TARGET, "REVIEW", "No reliable suggestion"

# ==================================================
# DATA ROW
# ==================================================

class MappingRow(object):
    def __init__(self, source_name, detected_type, suggested_target, action, reason,
                 element_count, target_count):
        self.SourceName = source_name
        self.DetectedType = detected_type
        self.SuggestedTarget = suggested_target if suggested_target else ""
        self.FinalTarget = suggested_target if suggested_target else PLACEHOLDER_TARGET
        self.Action = action
        self.Reason = reason
        self.ElementCount = element_count
        self.TargetElementCount = target_count
        self.Notes = ""

        self.AvailableTargets = ObservableCollection[object]()
        for item in ALL_TARGETS:
            self.AvailableTargets.Add(item)

        self.AvailableActions = ObservableCollection[object]()
        for item in ALL_ACTIONS:
            self.AvailableActions.Add(item)

# ==================================================
# WINDOW
# ==================================================

class WorksetMappingManagerWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)

        script_dir = os.path.dirname(__file__)
        logo_path = os.path.join(script_dir, "logo.png")
        try:
            if os.path.exists(logo_path):
                self.logoImage.Source = BitmapImage(Uri(logo_path, UriKind.Absolute))
        except Exception as ex:
            print("Logo load error: {}".format(safe_str(ex)))

        self._all_rows = ObservableCollection[object]()
        self._view = None
        self._workset_map = get_workset_table_by_name()
        self._count_by_wsid = build_element_count_by_workset()

        if hasattr(self, "chkDeleteEmptyAfterMap"):
            self.chkDeleteEmptyAfterMap.IsEnabled = api_supports_delete_workset()
            self.chkDeleteEmptyAfterMap.IsChecked = False

        self.setup_grid()
        self.bind_events()
        self.on_suggest_standards(None, None)
        self.update_selected_standards_count()
        self.refresh_summary()
        self.update_status("Ready. Click ANALYZE to load workset mapping.")

        # ------additional code
    def ensure_missing_target_worksets(self, rows_to_map):
        self._workset_map = get_workset_table_by_name()

        missing_targets = []
        seen = set()

        for row in rows_to_map:
            target_name = row.FinalTarget

            if is_target_empty(target_name):
                continue

            if target_name in self._workset_map:
                continue

            if target_name in seen:
                continue

            seen.add(target_name)
            missing_targets.append(target_name)

        if not missing_targets:
            return [], [], []

        created = []
        existing = []
        failed = []

        t = DB.Transaction(doc, "pyMENVIC | Auto-Create Missing Target Worksets")
        t.Start()
        try:
            for name in missing_targets:
                try:
                    if create_workset_if_missing(name):
                        created.append(name)
                    else:
                        existing.append(name)
                except Exception as ex:
                    failed.append((name, safe_str(ex)))
            t.Commit()
        except Exception as ex:
            t.RollBack()
            failed.append(("-", "Create missing targets transaction failed: {}".format(safe_str(ex))))

        self._workset_map = get_workset_table_by_name()
        self._count_by_wsid = build_element_count_by_workset()
        self.refresh_target_counts()

        return created, existing, failed

    # ---------------------------------------------
    # Setup
    # ---------------------------------------------

    def setup_grid(self):
        self.dgMappings.ItemsSource = self._all_rows
        self._view = CollectionViewSource.GetDefaultView(self.dgMappings.ItemsSource)
        self._view.Filter = self.row_filter

    def bind_events(self):
        self.btnClose.Click += self.on_close
        self.btnRefresh.Click += self.on_refresh
        self.btnSuggestStandards.Click += self.on_suggest_standards
        self.btnCreateSelectedStandards.Click += self.on_create_selected_standards
        self.btnAnalyze.Click += self.on_analyze
        self.btnPreview.Click += self.on_preview
        self.btnApply.Click += self.on_apply
        self.btnDeleteEmptyWorksets.Click += self.on_delete_empty_worksets

        self.cmbDisciplineMode.SelectionChanged += self.on_discipline_changed
        self.txtSearch.TextChanged += self.on_filter_changed
        self.cmbActionFilter.SelectionChanged += self.on_filter_changed

        for chk_name in self.get_standard_checkbox_names():
            if hasattr(self, chk_name):
                getattr(self, chk_name).Checked += self.on_standard_checkbox_changed
                getattr(self, chk_name).Unchecked += self.on_standard_checkbox_changed

    def on_discipline_changed(self, sender, args):
        self.on_suggest_standards(sender, args)
        self.update_selected_standards_count()
        self.update_status("Discipline updated. Click ANALYZE to refresh mapping.")

    # ---------------------------------------------
    # Small helpers
    # ---------------------------------------------

    def get_selected_discipline(self):
        try:
            item = self.cmbDisciplineMode.SelectedItem
            if not item:
                return "Architecture"
            visible_name = normalize_name(str(item.Content))
            return DISCIPLINE_DISPLAY_TO_INTERNAL.get(visible_name, "Architecture")
        except Exception:
            return "Architecture"

    def get_action_filter_text(self):
        try:
            item = self.cmbActionFilter.SelectedItem
            return str(item.Content) if item else "ALL"
        except Exception:
            return "ALL"

    def refresh_grid_ui(self):
        try:
            self.dgMappings.Items.Refresh()
        except Exception:
            pass

        try:
            self._view.Refresh()
        except Exception:
            pass

        self.refresh_target_counts()
        self.refresh_summary()

    def update_status(self, text):
        self.txtStatus.Text = text

    def get_standard_checkbox_names(self):
        return list(STANDARD_CHECKBOX_TO_WORKSET.keys())

    def set_checkbox(self, chk_name, value):
        if hasattr(self, chk_name):
            getattr(self, chk_name).IsChecked = value

    def get_selected_standard_names(self):
        result = []
        for chk_name in self.get_standard_checkbox_names():
            if hasattr(self, chk_name):
                chk = getattr(self, chk_name)
                try:
                    if chk.IsChecked:
                        result.append(STANDARD_CHECKBOX_TO_WORKSET[chk_name])
                except Exception:
                    pass
        return result

    def update_selected_standards_count(self):
        self.txtSelectedStandards.Text = str(len(self.get_selected_standard_names()))

    def apply_auto_defaults_after_analysis(self):
        for row in self._all_rows:
            if row.Action == "REVIEW":
                continue
            if is_target_empty(row.SuggestedTarget):
                continue
            row.FinalTarget = row.SuggestedTarget
        self.refresh_grid_ui()

    # ---------------------------------------------
    # Filters
    # ---------------------------------------------

    def row_filter(self, obj):
        row = obj

        search = self.txtSearch.Text.strip().lower() if self.txtSearch.Text else ""
        if search and search not in row.SourceName.lower():
            return False

        action_filter = self.get_action_filter_text()
        if action_filter != "ALL" and row.Action != action_filter:
            return False

        return True

    def on_filter_changed(self, sender, args):
        try:
            self._view.Refresh()
        except Exception:
            pass
        self.refresh_summary()

    # ---------------------------------------------
    # Summary
    # ---------------------------------------------

    def refresh_summary(self):
        rows = [r for r in self._all_rows]

        self.txtTotalRows.Text = str(len(rows))
        self.txtMappedRows.Text = str(len([r for r in rows if r.Action == "ASSIGN"]))
        self.txtKeepRows.Text = str(len([r for r in rows if r.Action == "KEEP"]))
        self.txtIgnoreRows.Text = str(len([r for r in rows if r.Action == "IGNORE"]))
        self.txtReviewRows.Text = str(len([r for r in rows if r.Action == "REVIEW"]))
        self.update_selected_standards_count()

    def refresh_target_counts(self):
        self._workset_map = get_workset_table_by_name()
        self._count_by_wsid = build_element_count_by_workset()

        for row in self._all_rows:
            target_name = row.FinalTarget
            if (not is_target_empty(target_name)) and target_name in self._workset_map:
                target_wsid = self._workset_map[target_name].Id.IntegerValue
                row.TargetElementCount = self._count_by_wsid.get(target_wsid, 0)
            else:
                row.TargetElementCount = 0

        try:
            self.dgMappings.Items.Refresh()
        except Exception:
            pass

    # ---------------------------------------------
    # Standards selection
    # ---------------------------------------------

    def on_standard_checkbox_changed(self, sender, args):
        self.update_selected_standards_count()

    def on_suggest_standards(self, sender, args):
        discipline = self.get_selected_discipline()

        for chk_name in self.get_standard_checkbox_names():
            self.set_checkbox(chk_name, False)

        discipline_to_core_checkboxes = {
            "Architecture": ["chkARC_MODEL", "chkARC_LEVELS_GRIDS"],
            "Structure": ["chkSTR_MODEL", "chkSTR_LEVELS_GRIDS"],
            "Mechanical": ["chkMECH_MODEL", "chkMECH_LEVELS_GRIDS"],
            "Electrical": ["chkELE_MODEL", "chkELE_LEVELS_GRIDS"],
            "Plumbing": ["chkPLM_MODEL", "chkPLM_LEVELS_GRIDS"],
            "Site": ["chkSITE_MODEL", "chkSITE_LEVELS_GRIDS"]
        }

        discipline_to_link_checkbox = {
            "Architecture": "chkLINK_ARC",
            "Structure": "chkLINK_STR",
            "Mechanical": "chkLINK_MECH",
            "Electrical": "chkLINK_ELE",
            "Plumbing": "chkLINK_PLM",
            "Site": "chkLINK_SITE"
        }

        if discipline in ["Coordination", "All Disciplines"]:
            for chk_name in self.get_standard_checkbox_names():
                self.set_checkbox(chk_name, True)
        else:
            core_chks = discipline_to_core_checkboxes.get(discipline, [])
            link_chk = discipline_to_link_checkbox.get(discipline)

            for chk_name in core_chks:
                self.set_checkbox(chk_name, True)

            if link_chk:
                self.set_checkbox(link_chk, True)
                self.set_checkbox("chkLINK_REF", True)

        self.update_selected_standards_count()
        self.update_status("Suggested standards loaded for {}.".format(discipline.upper()))

    def on_create_selected_standards(self, sender, args):
        selected = self.get_selected_standard_names()

        if not selected:
            MessageBox.Show("No standards selected.", "Workset Mapping Manager")
            return

        created = []
        existing = []

        t = DB.Transaction(doc, "pyMENVIC | Create Selected Standards")
        t.Start()
        try:
            for name in selected:
                try:
                    if create_workset_if_missing(name):
                        created.append(name)
                    else:
                        existing.append(name)
                except Exception:
                    existing.append(name)
            t.Commit()
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Could not create standards:\n{}".format(safe_str(ex)), "Workset Mapping Manager")
            return

        self._workset_map = get_workset_table_by_name()
        self._count_by_wsid = build_element_count_by_workset()
        self.refresh_target_counts()

        output.print_md("# pyMENVIC | CREATE SELECTED STANDARDS")
        output.print_md("")
        output.print_md("- **Created:** {}".format(len(created)))
        output.print_md("- **Already existing:** {}".format(len(existing)))
        output.print_md("")

        if created:
            output.print_md("## Created")
            output.print_md("")
            for name in created:
                output.print_md("- `{}`".format(name))
            output.print_md("")

        if existing:
            output.print_md("## Already existing")
            output.print_md("")
            for name in existing:
                output.print_md("- `{}`".format(name))
            output.print_md("")

        self.update_status("Standards checked. Created: {}. Existing: {}.".format(len(created), len(existing)))

    # ---------------------------------------------
    # Analyze
    # ---------------------------------------------

    def build_rows(self):
        discipline = self.get_selected_discipline()
        self._workset_map = get_workset_table_by_name()
        self._count_by_wsid = build_element_count_by_workset()

        rows = ObservableCollection[object]()
        selected_standards = self.get_selected_standard_names()

        for ws_name in sorted(self._workset_map.keys()):
            if ws_name in selected_standards:
                continue

            detected_type = detect_type(ws_name)
            suggested_target, action, reason = suggest_target_and_action(ws_name, discipline)

            source_wsid = self._workset_map[ws_name].Id.IntegerValue
            element_count = self._count_by_wsid.get(source_wsid, 0)

            target_count = 0
            if (not is_target_empty(suggested_target)) and suggested_target in self._workset_map:
                target_wsid = self._workset_map[suggested_target].Id.IntegerValue
                target_count = self._count_by_wsid.get(target_wsid, 0)

            row = MappingRow(
                ws_name,
                detected_type,
                suggested_target,
                action,
                reason,
                element_count,
                target_count
            )
            rows.Add(row)

        return rows

    def on_analyze(self, sender, args):
        self.update_status("Analyzing model... please wait.")
        self._all_rows = self.build_rows()
        self.dgMappings.ItemsSource = self._all_rows
        self._view = CollectionViewSource.GetDefaultView(self.dgMappings.ItemsSource)
        self._view.Filter = self.row_filter
        self.refresh_grid_ui()
        self.update_status("Analysis completed for {}.".format(self.get_selected_discipline()))

    # ---------------------------------------------
    # Preview
    # ---------------------------------------------

    def on_preview(self, sender, args):
        rows_to_map = [r for r in self._all_rows if r.Action == "ASSIGN"]

        rename_duplicates = False
        delete_empty = False
        try:
            rename_duplicates = bool(self.chkRenameDuplicates.IsChecked)
        except Exception:
            pass
        try:
            delete_empty = bool(self.chkDeleteEmptyAfterMap.IsChecked)
        except Exception:
            pass

        if rename_duplicates:
            filtered_rows = []
            for r in rows_to_map:
                n = normalize_name(r.SourceName)
                if n in ["WORKSET1", "WORKSET 1"]:
                    continue
                filtered_rows.append(r)
            rows_to_map = filtered_rows

        invalid = [r for r in rows_to_map if is_target_empty(r.FinalTarget)]
        if invalid:
            MessageBox.Show(
                "There are rows marked as ASSIGN without Final Target.\nPlease review them before applying.",
                "Workset Mapping Manager"
            )
            self.update_status("Preview found rows marked as ASSIGN without target.")
            return

        output.print_md("# pyMENVIC | WORKSET MAPPING MANAGER — PREVIEW")
        output.print_md("")
        output.print_md("## Summary")
        output.print_md("")
        output.print_md("- **Discipline mode:** {}".format(self.get_selected_discipline()))
        output.print_md("- **Rows marked as ASSIGN:** {}".format(len(rows_to_map)))
        output.print_md("- **Rename / consolidate duplicates:** {}".format("Yes" if rename_duplicates else "No"))
        output.print_md("- **Delete empty leftovers after apply:** {}".format("Yes" if delete_empty else "No"))
        output.print_md("")

        output.print_md("## Planned mappings")
        output.print_md("")
        for row in sorted(rows_to_map, key=lambda x: x.SourceName):
            output.print_md(
                "- `{}` → `{}` | elements: **{}** | reason: {}".format(
                    row.SourceName,
                    row.FinalTarget,
                    row.ElementCount,
                    row.Reason
                )
            )
        output.print_md("")
        self.update_status("Preview printed to pyRevit output.")

    # ---------------------------------------------
    # Delete Empty Non-Standard Worksets
    # ---------------------------------------------

    def on_delete_empty_worksets(self, sender, args):
        self._workset_map = get_workset_table_by_name()
        self._count_by_wsid = build_element_count_by_workset()

        standard_names = set(ALL_CORE + ALL_LINKS)
        candidates = []

        for ws_name in sorted(self._workset_map.keys()):
            if ws_name in standard_names:
                continue

            try:
                ws = self._workset_map[ws_name]
                count = self._count_by_wsid.get(ws.Id.IntegerValue, 0)
            except Exception:
                count = 0

            if count == 0:
                candidates.append(ws_name)

        if not candidates:
            MessageBox.Show(
                "No empty non-standard worksets were found.",
                "Workset Mapping Manager"
            )
            self.update_status("No empty non-standard worksets found.")
            return

        confirm = forms.alert(
            "Delete {} empty non-standard workset(s)?".format(len(candidates)),
            yes=True,
            no=True
        )
        if not confirm:
            self.update_status("Delete empty worksets canceled by user.")
            return

        deleted = []
        skipped = []
        failed = []

        if not api_supports_delete_workset():
            MessageBox.Show(
                "Delete workset API not available in this Revit version.",
                "Workset Mapping Manager"
            )
            self.update_status("Delete worksets failed. API not available.")
            return

        ids_to_checkout = []
        for ws_name in candidates:
            try:
                ws = self._workset_map.get(ws_name)
                if ws:
                    ids_to_checkout.append(ws.Id)
            except Exception:
                pass

        checked_out, failed_checkout = checkout_worksets_if_possible(ids_to_checkout)
        if failed_checkout:
            for _, reason in failed_checkout:
                failed.append(("-", "Checkout failed: {}".format(reason)))

        t = DB.Transaction(doc, "pyMENVIC | Delete Empty Non-Standard Worksets")
        t.Start()
        try:
            for ws_name in candidates:
                try:
                    self._workset_map = get_workset_table_by_name()
                    ws = self._workset_map.get(ws_name)

                    if not ws:
                        skipped.append((ws_name, "Workset not found"))
                        continue

                    if ws_name in standard_names:
                        skipped.append((ws_name, "Protected standard workset"))
                        continue

                    remaining = get_true_workset_element_count(ws.Id)
                    if remaining > 0:
                        skipped.append((ws_name, "Workset still contains {} element(s)".format(remaining)))
                        continue

                    if not can_delete_workset_safe(ws.Id):
                        skipped.append((ws_name, "Revit did not allow deleting this workset"))
                        continue

                    delete_workset_safe(ws.Id)
                    deleted.append(ws_name)

                except Exception as ex:
                    failed.append((ws_name, safe_str(ex)))

            t.Commit()
        except Exception as ex:
            t.RollBack()
            failed.append(("-", "Delete transaction failed: {}".format(safe_str(ex))))

        self._workset_map = get_workset_table_by_name()
        self._count_by_wsid = build_element_count_by_workset()

        to_remove = [r for r in self._all_rows if r.SourceName in deleted]
        for row in to_remove:
            try:
                self._all_rows.Remove(row)
            except Exception:
                pass

        self.refresh_grid_ui()

        output.print_md("# pyMENVIC | DELETE EMPTY NON-STANDARD WORKSETS")
        output.print_md("")
        output.print_md("- **Deleted:** {}".format(len(deleted)))
        output.print_md("- **Skipped:** {}".format(len(skipped)))
        output.print_md("- **Errors:** {}".format(len(failed)))
        output.print_md("")

        if deleted:
            output.print_md("## Deleted")
            output.print_md("")
            for name in deleted:
                output.print_md("- `{}`".format(name))
            output.print_md("")

        if skipped:
            output.print_md("## Not deleted")
            output.print_md("")
            for name, reason in skipped:
                output.print_md("- `{}` | {}".format(name, reason))
            output.print_md("")

        if failed:
            output.print_md("## Errors")
            output.print_md("")
            for name, reason in failed:
                output.print_md("- `{}` | {}".format(name, reason))
            output.print_md("")

        self.update_status(
            "Delete empty worksets completed. Deleted: {}. Skipped: {}.".format(
                len(deleted),
                len(skipped)
            )
        )

        MessageBox.Show(
            "Delete empty worksets completed.\n\nDeleted: {}\nSkipped: {}\nErrors: {}".format(
                len(deleted),
                len(skipped),
                len(failed)
            ),
            "Workset Mapping Manager"
        )

    # ---------------------------------------------
    # Cleanup
    # ---------------------------------------------

    def delete_empty_leftovers(self, rows_to_map):
        deleted = []
        skipped = []
        failed = []

        if not api_supports_delete_workset():
            skipped.append(("-", "Delete workset API not available"))
            return deleted, skipped, failed

        t = DB.Transaction(doc, "pyMENVIC | Delete Empty Leftovers")
        t.Start()
        try:
            for row in rows_to_map:
                try:
                    self._workset_map = get_workset_table_by_name()

                    source_ws = self._workset_map.get(row.SourceName)
                    target_ws = self._workset_map.get(row.FinalTarget)

                    if not source_ws:
                        skipped.append((row.SourceName, "Source workset not found"))
                        continue

                    if not target_ws:
                        skipped.append((row.SourceName, "Target workset not found"))
                        continue

                    source_id = source_ws.Id
                    target_id = target_ws.Id

                    if row.SourceName in ALWAYS_CORE:
                        skipped.append((row.SourceName, "Protected ARC base workset"))
                        continue

                    remaining = get_true_workset_element_count(source_id)

                    if remaining <= 0:
                        if not can_delete_workset_safe(source_id):
                            skipped.append((row.SourceName, "Revit did not allow deleting this workset"))
                            continue

                        delete_workset_safe(source_id)
                        deleted.append(row.SourceName)
                        continue

                    move_settings = build_delete_move_settings(target_id)
                    if move_settings is None:
                        skipped.append((row.SourceName, "Delete move settings not available in this Revit version"))
                        continue

                    if not can_delete_workset_safe(source_id, move_settings):
                        skipped.append((row.SourceName, "Revit did not allow move-and-delete for this workset"))
                        continue

                    delete_workset_safe(source_id, move_settings)
                    deleted.append(row.SourceName)

                except Exception as ex:
                    failed.append((row.SourceName, safe_str(ex)))

            t.Commit()
        except Exception as ex:
            t.RollBack()
            failed.append(("-", "Delete transaction failed: {}".format(safe_str(ex))))

        return deleted, skipped, failed

    # ---------------------------------------------
    # Apply
    # ---------------------------------------------

    def on_apply(self, sender, args):
        rows_to_map = [r for r in self._all_rows if r.Action == "ASSIGN"]

        rename_duplicates = False
        delete_empty = False
        try:
            rename_duplicates = bool(self.chkRenameDuplicates.IsChecked)
        except Exception:
            pass
        try:
            delete_empty = bool(self.chkDeleteEmptyAfterMap.IsChecked)
        except Exception:
            pass

        if rename_duplicates:
            filtered_rows = []
            for r in rows_to_map:
                n = normalize_name(r.SourceName)
                if n in ["WORKSET1", "WORKSET 1"]:
                    continue
                filtered_rows.append(r)
            rows_to_map = filtered_rows

        if not rows_to_map and not rename_duplicates:
            MessageBox.Show("No rows are marked as ASSIGN.", "Workset Mapping Manager")
            self.update_status("Nothing to apply.")
            return

        invalid = [r for r in rows_to_map if is_target_empty(r.FinalTarget)]
        if invalid:
            MessageBox.Show(
                "Some rows are marked as ASSIGN but have no Final Target.\nPlease review them before applying.",
                "Workset Mapping Manager"
            )
            self.update_status("Apply canceled. Missing target detected.")
            return

        same_target = [r for r in rows_to_map if r.SourceName == r.FinalTarget]
        if same_target:
            MessageBox.Show(
                "Some rows are mapped to the same source workset name.\nPlease review them before applying.",
                "Workset Mapping Manager"
            )
            self.update_status("Apply canceled. Same source/target detected.")
            return
# added text
        created_targets, existing_targets, failed_create_targets = self.ensure_missing_target_worksets(rows_to_map)

        if failed_create_targets:
            msg_lines = []
            for name, reason in failed_create_targets[:10]:
                msg_lines.append("{} | {}".format(name, reason))

            MessageBox.Show(
                "Some target worksets could not be created automatically:\n\n{}".format("\n".join(msg_lines)),
                "Workset Mapping Manager"
            )
            self.update_status("Apply canceled. Failed to auto-create some target worksets.")
            return

        if created_targets:
            output.print_md("## Auto-created target worksets")
            output.print_md("")
            for name in created_targets:
                output.print_md("- `{}`".format(name))
            output.print_md("")

# ended of added text

        total_apply_rows = len(rows_to_map)
        if rename_duplicates:
            total_apply_rows += 1

        confirm = forms.alert(
            "Apply mapping to {} row(s)?\n\nRename / consolidate duplicates: {}\nDelete empty leftovers after apply: {}".format(
                total_apply_rows,
                "Yes" if rename_duplicates else "No",
                "Yes" if delete_empty else "No"
            ),
            yes=True,
            no=True
        )
        if not confirm:
            self.update_status("Apply canceled by user.")
            return

        changed_total = 0
        skipped_total = 0
        error_rows = []

        ids_to_checkout = []
        self._workset_map = get_workset_table_by_name()
        for row in rows_to_map:
            try:
                source_ws = self._workset_map.get(row.SourceName)
                target_ws = self._workset_map.get(row.FinalTarget)
                if source_ws:
                    ids_to_checkout.append(source_ws.Id)
                if target_ws:
                    ids_to_checkout.append(target_ws.Id)
            except Exception:
                pass

        if rename_duplicates:
            try:
                ws_map = get_workset_table_by_name()
                workset1 = ws_map.get("Workset1") or ws_map.get("WORKSET1") or ws_map.get("Workset 1")
                if workset1:
                    ids_to_checkout.append(workset1.Id)
                    base_target = get_base_model_target_for_discipline(self.get_selected_discipline())
                    if base_target and base_target in ws_map:
                        ids_to_checkout.append(ws_map[base_target].Id)
            except Exception:
                pass

        checked_out, failed_checkout = checkout_worksets_if_possible(ids_to_checkout)
        if failed_checkout:
            for _, reason in failed_checkout:
                error_rows.append(("WORKSET CHECKOUT", "-", reason))

        if rename_duplicates:
            rename_summary = handle_workset1_rename_consolidation(self.get_selected_discipline())
            changed_total += rename_summary.get("changed", 0)
            skipped_total += rename_summary.get("skipped", 0)

            for msg in rename_summary.get("messages", []):
                output.print_md("**Rename step:** {}".format(msg))

            for err in rename_summary.get("errors", []):
                error_rows.append(("Workset1", get_base_model_target_for_discipline(self.get_selected_discipline()) or "-", err))

            if rename_summary.get("messages") or rename_summary.get("errors"):
                output.print_md("")

            self._workset_map = get_workset_table_by_name()
            self._count_by_wsid = build_element_count_by_workset()
            self.refresh_target_counts()

        t = DB.Transaction(doc, "pyMENVIC | Apply Workset Mapping")
        t.Start()
        try:
            for row in rows_to_map:
                try:
                    self._workset_map = get_workset_table_by_name()

                    source_ws = self._workset_map.get(row.SourceName)
                    target_ws = self._workset_map.get(row.FinalTarget)

                    if not source_ws or not target_ws:
                        error_rows.append((row.SourceName, row.FinalTarget, "Workset not found"))
                        continue

                    source_wsid = source_ws.Id.IntegerValue
                    target_wsid = target_ws.Id.IntegerValue

                    if source_wsid == target_wsid:
                        skipped_total += 1
                        continue

                    changed_this_row = 0
                    failed_this_row = []

                    for el in get_reassignable_elements_by_workset_id(source_wsid):
                        try:
                            p = el.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)

                            if not p:
                                failed_this_row.append("Element {} has no workset parameter".format(el.Id.IntegerValue))
                                continue

                            try:
                                if p.IsReadOnly:
                                    failed_this_row.append("Element {} workset parameter is read-only".format(el.Id.IntegerValue))
                                    continue
                            except Exception:
                                pass

                            p.Set(target_wsid)
                            changed_this_row += 1

                        except Exception as ex:
                            failed_this_row.append("Element {} failed: {}".format(el.Id.IntegerValue, safe_str(ex)))

                    changed_total += changed_this_row

                    if failed_this_row:
                        preview_failures = "; ".join(failed_this_row[:5])
                        if len(failed_this_row) > 5:
                            preview_failures += " ... (+{} more)".format(len(failed_this_row) - 5)

                        error_rows.append((
                            row.SourceName,
                            row.FinalTarget,
                            "{} moved / {} failed | {}".format(changed_this_row, len(failed_this_row), preview_failures)
                        ))

                except Exception as ex:
                    error_rows.append((row.SourceName, row.FinalTarget, safe_str(ex)))

            t.Commit()
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Transaction failed:\n{}".format(safe_str(ex)), "Workset Mapping Manager")
            self.update_status("Apply failed.")
            return

        self._workset_map = get_workset_table_by_name()
        self._count_by_wsid = build_element_count_by_workset()
        self.refresh_target_counts()

        deleted_worksets = []
        skipped_delete = []
        failed_delete = []

        if delete_empty:
            deleted_worksets, skipped_delete, failed_delete = self.delete_empty_leftovers(rows_to_map)
            self._workset_map = get_workset_table_by_name()
            self._count_by_wsid = build_element_count_by_workset()
            self.refresh_target_counts()

        if deleted_worksets:
            to_remove = [r for r in self._all_rows if r.SourceName in deleted_worksets]
            for row in to_remove:
                try:
                    self._all_rows.Remove(row)
                except Exception:
                    pass
            self.refresh_grid_ui()

        output.print_md("# pyMENVIC | WORKSET MAPPING MANAGER — CLEANUP SUMMARY")
        output.print_md("")
        output.print_md("## Mapping")
        output.print_md("")
        output.print_md("- **Discipline mode:** {}".format(self.get_selected_discipline()))
        output.print_md("- **Rows assigned:** {}".format(total_apply_rows))
        output.print_md("- **Elements reassigned:** {}".format(changed_total))
        output.print_md("- **Rows skipped:** {}".format(skipped_total))
        output.print_md("- **Rows with mapping errors:** {}".format(len(error_rows)))
        output.print_md("- **Rename / consolidate duplicates:** {}".format("Yes" if rename_duplicates else "No"))
        output.print_md("- **Delete empty leftovers:** {}".format("Yes" if delete_empty else "No"))
        output.print_md("")

        if error_rows:
            output.print_md("## Mapping errors")
            output.print_md("")
            for src, dst, err in error_rows:
                output.print_md("- `{}` → `{}` | {}".format(src, dst, err))
            output.print_md("")

        output.print_md("## Cleanup")
        output.print_md("")
        output.print_md("- **Worksets deleted:** {}".format(len(deleted_worksets)))
        output.print_md("- **Delete skipped:** {}".format(len(skipped_delete)))
        output.print_md("- **Delete errors:** {}".format(len(failed_delete)))
        output.print_md("")

        if deleted_worksets:
            output.print_md("### Deleted worksets")
            output.print_md("")
            for name in sorted(deleted_worksets):
                output.print_md("- `{}`".format(name))
            output.print_md("")

        if skipped_delete:
            output.print_md("### Not deleted")
            output.print_md("")
            for name, reason in skipped_delete:
                output.print_md("- `{}` | {}".format(name, reason))
            output.print_md("")

        if failed_delete:
            output.print_md("### Delete errors")
            output.print_md("")
            for name, reason in failed_delete:
                output.print_md("- `{}` | {}".format(name, reason))
            output.print_md("")

        self.update_status(
            "Apply completed. {} elements reassigned. {} worksets deleted.".format(
                changed_total,
                len(deleted_worksets)
            )
        )

        MessageBox.Show(
            "Mapping completed.\n\nElements reassigned: {}\nMapping errors: {}\nWorksets deleted: {}\nDelete errors: {}".format(
                changed_total,
                len(error_rows),
                len(deleted_worksets),
                len(failed_delete)
            ),
            "Workset Mapping Manager"
        )

    # ---------------------------------------------
    # Refresh / Close
    # ---------------------------------------------

    def on_refresh(self, sender, args):
        self._workset_map = get_workset_table_by_name()
        self._count_by_wsid = build_element_count_by_workset()
        self.refresh_grid_ui()
        self.update_selected_standards_count()
        self.update_status("Refreshed.")

    def on_close(self, sender, args):
        self.Close()

# ==================================================
# MAIN
# ==================================================

if not doc.IsWorkshared:
    forms.alert("The model is not workshared.", exitscript=True)

xaml_file = script.get_bundle_file("WorksetMappingManager.xaml")
if not xaml_file:
    forms.alert("WorksetMappingManager.xaml was not found in the bundle.", exitscript=True)

window = WorksetMappingManagerWindow(xaml_file)
window.ShowDialog()
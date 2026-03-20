# -*- coding: utf-8 -*-
__title__ = "Link Workset Manager"
__author__ = "Ricardo J. Mendieta"

__doc__ = """
LINK WORKSET MANAGER
--------------------------------------------------
Shows loaded Revit link instances in a table and allows
assigning a target host workset with a dropdown.

Applies changes to:
- RevitLinkInstance
- RevitLinkType

Behavior:
- No row selection is required
- Only rows with a changed target workset are processed
- Rows with no effective change are ignored

Important:
If multiple rows share the same Link Type, they
must all end with the same target workset, otherwise
the tool will stop and warn about the conflict.
"""

from pyrevit import revit, DB, forms
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Data import CollectionViewSource
from System.Windows.Controls import DataGridEditingUnit
from System.Windows.Input import Keyboard
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

doc = revit.doc

DEFAULT_LINK_WORKSETS = [
    "LINK_ARC",
    "LINK_STR",
    "LINK_MECH",
    "LINK_ELE",
    "LINK_PLM",
    "LINK_SITE",
    "LINK_CAD",
    "LINK_REF"
]


# ==================================================
# HELPERS
# ==================================================

def safe_str(value):
    try:
        if value is None:
            return ""
        return str(value)
    except Exception:
        return ""


def get_elem_name(elem):
    try:
        return DB.Element.Name.GetValue(elem)
    except Exception:
        try:
            p = elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if p:
                return p.AsString()
        except Exception:
            pass
        try:
            return safe_str(elem.Id)
        except Exception:
            return "<Unnamed>"


def ensure_default_link_worksets():
    created = []

    try:
        if not doc.IsWorkshared:
            return created
    except Exception:
        return created

    existing_names = set()
    try:
        collector = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
        for ws in collector:
            existing_names.add(safe_str(ws.Name).strip())
    except Exception:
        return created

    missing_names = [name for name in DEFAULT_LINK_WORKSETS if name not in existing_names]
    if not missing_names:
        return created

    t = DB.Transaction(doc, "pyMenvic | Create Default Link Worksets")
    t.Start()

    try:
        for ws_name in missing_names:
            DB.Workset.Create(doc, ws_name)
            created.append(ws_name)
        t.Commit()
    except Exception:
        t.RollBack()
        raise

    return created


def get_user_worksets():
    worksets = []
    try:
        collector = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
        for ws in collector:
            worksets.append(ws)
    except Exception:
        pass
    return worksets


def get_workset_name_by_id(workset_id):
    try:
        ws_table = doc.GetWorksetTable()
        ws = ws_table.GetWorkset(workset_id)
        if ws:
            return safe_str(ws.Name)
    except Exception:
        pass
    return ""


def get_element_workset_name(elem):
    try:
        p = elem.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
        if p and p.HasValue:
            return get_workset_name_by_id(DB.WorksetId(p.AsInteger()))
    except Exception:
        pass
    return ""


def set_element_workset(elem, target_workset_id):
    p = elem.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
    if not p or p.IsReadOnly:
        return False, "Workset parameter missing or read-only."
    try:
        p.Set(target_workset_id.IntegerValue)
        return True, ""
    except Exception as ex:
        return False, safe_str(ex)


def get_link_status(inst):
    try:
        link_doc = inst.GetLinkDocument()
        return "Loaded" if link_doc else "Not Loaded"
    except Exception:
        return "Unknown"


def get_loaded_link_instances():
    rows = []
    collector = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)

    for inst in collector:
        try:
            type_elem = doc.GetElement(inst.GetTypeId())
            if not type_elem:
                continue

            inst_name = safe_str(get_elem_name(inst)).split(":")[0].strip()
            inst_name = inst_name.replace(".rvt", "")

            type_name = safe_str(get_elem_name(type_elem)).replace(".rvt", "")

            row = {
                "instance_id": inst.Id,
                "type_id": type_elem.Id,
                "link_name": safe_str(inst_name),
                "type_name": safe_str(type_name),
                "current_instance_ws": get_element_workset_name(inst),
                "current_type_ws": get_element_workset_name(type_elem),
                "status": get_link_status(inst),
                "pinned": "Yes" if inst.Pinned else "No"
            }

            rows.append(row)

        except Exception as ex:
            print("Error reading link instance: {}".format(safe_str(ex)))

    rows = sorted(rows, key=lambda x: x["link_name"].lower())
    return rows


def commit_grid_edits(grid):
    try:
        grid.CommitEdit(DataGridEditingUnit.Cell, True)
    except Exception:
        pass

    try:
        grid.CommitEdit(DataGridEditingUnit.Row, True)
    except Exception:
        pass

    try:
        grid.CommitEdit()
    except Exception:
        pass

    try:
        Keyboard.ClearFocus()
    except Exception:
        pass

    try:
        grid.UpdateLayout()
    except Exception:
        pass

    try:
        grid.Items.Refresh()
    except Exception:
        pass



def infer_link_workset(name):
    try:
        n = safe_str(name).upper()

        for ch in [".", "-", "_", "(", ")", "[", "]"]:
            n = n.replace(ch, " ")

        tokens = n.split()

        if any(t in tokens for t in ["STR", "SRT", "ST", "STRUCT", "STRUCTURE", "STEEL"]):
            return "LINK_STR", "STR", False

        if any(t in tokens for t in ["MEC", "ME", "MECH", "MECHANICAL", "HVAC"]):
            return "LINK_MECH", "MECH", False

        if any(t in tokens for t in ["ELE", "EL", "ELECT", "ELECTRICAL", "POWER", "LIGHT"]):
            return "LINK_ELE", "ELE", False

        if any(t in tokens for t in ["PLM", "PLUMB", "PLUMBING", "SANIT", "WATER", "DRAIN", "PIPE"]):
            return "LINK_PLM", "PLM", False

        if any(t in tokens for t in ["ARC", "AR", "ARCH", "ARCHITECTURE"]):
            return "LINK_ARC", "ARC", False

        if any(t in tokens for t in ["SITE", "TOPO", "LAND", "CIVIL"]):
            return "LINK_SITE", "SITE", False

        if any(t in tokens for t in ["CAD", "DWG", "DXF"]):
            return "LINK_CAD", "CAD", False

        if tokens:
            t0 = tokens[0]

            if t0.startswith("S"):
                return "LINK_STR", "STR", False
            if t0.startswith("M"):
                return "LINK_MECH", "MECH", False
            if t0.startswith("E"):
                return "LINK_ELE", "ELE", False
            if t0.startswith("P"):
                return "LINK_PLM", "PLM", False
            if t0.startswith("A"):
                return "LINK_ARC", "ARC", False
            if t0.startswith("C"):
                return "LINK_CAD", "CAD", False
            if t0.startswith("L") or t0.startswith("T"):
                return "LINK_SITE", "SITE", False

        return "LINK_REF", "REF", True

    except Exception:
        return "LINK_REF", "REF", True


def normalize_target_workset_name(name):
    return safe_str(name).strip().upper()


def get_safe_default_workset_name(preferred_name, workset_names):
    preferred_name = normalize_target_workset_name(preferred_name)

    normalized_existing = {normalize_target_workset_name(x): x for x in workset_names}

    if preferred_name and preferred_name in normalized_existing:
        return normalized_existing[preferred_name]

    if "LINK_REF" in normalized_existing:
        return normalized_existing["LINK_REF"]

    if workset_names:
        return workset_names[0]

    return preferred_name


def get_or_create_workset_by_name(target_name):
    target_name = normalize_target_workset_name(target_name)
    if not target_name:
        return None, False, "Empty target workset name."

    for ws in get_user_worksets():
        if normalize_target_workset_name(ws.Name) == target_name:
            return ws, False, ""

    if not doc.IsWorkshared:
        return None, False, "The model is not workshared."

    if hasattr(DB.WorksetTable, "IsWorksetNameUnique"):
        try:
            if not DB.WorksetTable.IsWorksetNameUnique(doc, target_name):
                for ws in get_user_worksets():
                    if normalize_target_workset_name(ws.Name) == target_name:
                        return ws, False, ""
        except Exception:
            pass

    try:
        ws = DB.Workset.Create(doc, target_name)
        return ws, True, ""
    except Exception as ex:
        return None, False, safe_str(ex)


def row_has_pending_change(row):
    try:
        target_name = normalize_target_workset_name(row.SelectedTargetWorksetName)
        current_inst = safe_str(row.CurrentInstanceWorkset).strip()
        current_type = safe_str(row.CurrentTypeWorkset).strip()

        if not target_name:
            return False

        return target_name != current_inst or target_name != current_type
    except Exception:
        return False


# ==================================================
# DATA ROW
# ==================================================

class LinkRow(INotifyPropertyChanged):
    def __init__(self, data, workset_names, status_callback=None):
        self._propertyChangedHandlers = []
        self._status_callback = status_callback

        self.InstanceId = data["instance_id"]
        self.TypeId = data["type_id"]

        self.LinkName = data["link_name"]
        self.TypeName = data["type_name"]
        self.CurrentInstanceWorkset = data["current_instance_ws"]
        self.CurrentTypeWorkset = data["current_type_ws"]
        self.Status = data["status"]
        self.PinnedText = data["pinned"]

        self.WorksetOptions = ObservableCollection[str]()
        for ws_name in workset_names:
            self.WorksetOptions.Add(ws_name)

        inferred_ws, inferred_rule, is_unrecognized = infer_link_workset(self.LinkName)
        if is_unrecognized:
            inferred_ws, inferred_rule, is_unrecognized = infer_link_workset(self.TypeName)

        self.AutoDetectedWorkset = safe_str(inferred_ws)
        self.DisciplineText = safe_str(inferred_rule)
        self.WarningText = ""
        self._stateText = ""
        self._isUnrecognized = bool(is_unrecognized)

        if self._isUnrecognized:
            self.WarningText = "Unrecognized link naming. Defaulted to LINK_REF."
        elif self.AutoDetectedWorkset not in list(workset_names):
            self.WarningText = "Detected '{}' but that workset was not found. Defaulted safely to LINK_REF.".format(self.AutoDetectedWorkset)
            self._isUnrecognized = True

        self._selectedTargetWorksetName = get_safe_default_workset_name(self.AutoDetectedWorkset, workset_names)

        self._isPendingChange = row_has_pending_change(self)
        self._isNotLoaded = safe_str(self.Status).strip().lower() == "not loaded"
        self._stateText = ""
        self.refresh_flags()

    def add_PropertyChanged(self, handler):
        self._propertyChangedHandlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._propertyChangedHandlers:
            self._propertyChangedHandlers.remove(handler)

    def _raise_property_changed(self, prop_name):
        args = PropertyChangedEventArgs(prop_name)
        for handler in list(self._propertyChangedHandlers):
            try:
                handler(self, args)
            except Exception:
                pass

    @property
    def SelectedTargetWorksetName(self):
        return self._selectedTargetWorksetName

    @SelectedTargetWorksetName.setter
    def SelectedTargetWorksetName(self, value):
        value = normalize_target_workset_name(value)

        if value == self._selectedTargetWorksetName:
            return

        self._selectedTargetWorksetName = value
        self._raise_property_changed("SelectedTargetWorksetName")
        self.refresh_flags()

        if self._status_callback:
            try:
                self._status_callback()
            except Exception:
                pass

    @property
    def IsPendingChange(self):
        return self._isPendingChange

    @IsPendingChange.setter
    def IsPendingChange(self, value):
        if value == self._isPendingChange:
            return
        self._isPendingChange = value
        self._raise_property_changed("IsPendingChange")

    @property
    def IsNotLoaded(self):
        return self._isNotLoaded

    @IsNotLoaded.setter
    def IsNotLoaded(self, value):
        if value == self._isNotLoaded:
            return
        self._isNotLoaded = value
        self._raise_property_changed("IsNotLoaded")


    @property
    def IsUnrecognized(self):
        return self._isUnrecognized

    @IsUnrecognized.setter
    def IsUnrecognized(self, value):
        if value == self._isUnrecognized:
            return
        self._isUnrecognized = value
        self._raise_property_changed("IsUnrecognized")

    @property
    def StateText(self):
        return self._stateText

    @StateText.setter
    def StateText(self, value):
        value = safe_str(value)
        if value == self._stateText:
            return
        self._stateText = value
        self._raise_property_changed("StateText")

    def refresh_flags(self):
        self.IsPendingChange = row_has_pending_change(self)
        self.IsNotLoaded = safe_str(self.Status).strip().lower() == "not loaded"

        if self.IsUnrecognized:
            self.StateText = "REVIEW"
        elif self.IsPendingChange:
            self.StateText = "CHANGE"
        else:
            self.StateText = "OK"

    def refresh_all_after_apply(self, new_inst_ws, new_type_ws):
        self.CurrentInstanceWorkset = safe_str(new_inst_ws)
        self.CurrentTypeWorkset = safe_str(new_type_ws)

        self._raise_property_changed("CurrentInstanceWorkset")
        self._raise_property_changed("CurrentTypeWorkset")

        self.refresh_flags()

        if self._status_callback:
            try:
                self._status_callback()
            except Exception:
                pass


# ==================================================
# WINDOW
# ==================================================

class LinkWorksetManagerWindow(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        import os
        from System.Windows.Media.Imaging import BitmapImage
        from System import Uri

        script_dir = os.path.dirname(__file__)
        logo_path = os.path.join(script_dir, "resources", "logo.png")

        try:
            if os.path.exists(logo_path):
                self.logoImage.Source = BitmapImage(Uri(logo_path))
        except Exception as ex:
            print("Logo load error: {}".format(ex))

        self.rows_all = []
        self.rows_view = None
        self.worksets = []
        self.workset_by_name = {}

        self.load_data()
        self.bind_events()

    def bind_events(self):
        self.btnRefresh.Click += self.on_refresh
        self.btnApply.Click += self.on_apply
        self.btnCancel.Click += self.on_cancel
        self.txtFilter.TextChanged += self.on_filter_changed
        self.dgLinks.CurrentCellChanged += self.on_grid_cell_changed

    def load_data(self):
        self.refresh_workset_cache()
        if not self.worksets:
            forms.alert("No user worksets were found in the host model.", exitscript=True)

        workset_names = sorted(self.workset_by_name.keys())

        raw_rows = get_loaded_link_instances()
        if not raw_rows:
            forms.alert("No Revit link instances were found in the model.", exitscript=True)

        obs = ObservableCollection[object]()
        self.rows_all = []

        for raw in raw_rows:
            row = LinkRow(raw, workset_names, self.on_row_value_changed)
            self.rows_all.append(row)
            obs.Add(row)

        self.dgLinks.ItemsSource = obs
        self.rows_view = CollectionViewSource.GetDefaultView(self.dgLinks.ItemsSource)
        self.update_status()


    def on_row_value_changed(self):
        self.update_status()


        try:
            self.dgLinks.Items.Refresh()
        except Exception:
            pass

    def refresh_workset_cache(self):
        self.worksets = sorted(get_user_worksets(), key=lambda x: safe_str(x.Name).lower())
        self.workset_by_name = {}

        for ws in self.worksets:
            ws_name = normalize_target_workset_name(ws.Name)
            self.workset_by_name[ws_name] = ws

        available_names = sorted(self.workset_by_name.keys())

        for row in self.rows_all:
            try:
                existing = set([safe_str(x) for x in row.WorksetOptions])
                for ws_name in available_names:
                    if ws_name not in existing:
                        row.WorksetOptions.Add(ws_name)
            except Exception:
                pass

    def update_row_flags(self):
        for row in self.rows_all:
            try:
                row.refresh_flags()
            except Exception:
                pass

    def update_status(self):
        total = len(self.rows_all)
        changed = 0
        unknown = 0
        ok_count = 0

        for row in self.rows_all:
            try:
                if row.IsPendingChange:
                    changed += 1
            except Exception:
                pass

            try:
                if row.IsUnrecognized:
                    unknown += 1
            except Exception:
                pass

            try:
                if (not row.IsPendingChange) and (not row.IsUnrecognized):
                    ok_count += 1
            except Exception:
                pass

        self.txtStatus.Text = "Links: {} | OK: {} | Pending: {} | Review: {}".format(total, ok_count, changed, unknown)

    def refresh_grid(self):
        self.update_row_flags()

        try:
            self.dgLinks.Items.Refresh()
        except Exception:
            pass

        self.update_status()


    def on_refresh(self, sender, args):
        self.Close()
        LinkWorksetManagerWindow("ui.xaml").ShowDialog()

    def on_cancel(self, sender, args):
        self.Close()

    def on_filter_changed(self, sender, args):
        if self.rows_view:
            self.rows_view.Filter = self._filter_delegate
            self.rows_view.Refresh()
        self.update_status()


    def on_grid_cell_changed(self, sender, args):
        commit_grid_edits(self.dgLinks)
        self.refresh_grid()

    def _filter_delegate(self, item):
        return self._passes_filter(item)

    def _passes_filter(self, row):
        try:
            txt = safe_str(self.txtFilter.Text).strip().lower()
            if not txt:
                return True

            haystack = " | ".join([
                safe_str(row.LinkName),
                safe_str(row.TypeName),
                safe_str(row.CurrentInstanceWorkset),
                safe_str(row.CurrentTypeWorkset),
                safe_str(row.Status),
                safe_str(row.PinnedText),
                safe_str(row.SelectedTargetWorksetName),
                safe_str(row.DisciplineText),
                safe_str(row.WarningText)
            ]).lower()

            return txt in haystack
        except Exception:
            return True

    def on_apply(self, sender, args):
        commit_grid_edits(self.dgLinks)

        try:
            self.dgLinks.SelectedItem = None
        except Exception:
            pass

        try:
            self.dgLinks.CurrentCell = None
        except Exception:
            pass

        try:
            self.dgLinks.UpdateLayout()
        except Exception:
            pass

        self.update_status()


        changed_rows = []
        for row in self.rows_all:
            if row_has_pending_change(row):
                changed_rows.append(row)

        if not changed_rows:
            forms.alert(
                "No changes detected. All links are already assigned to their target workset.",
                warn_icon=False
            )
            return

        conflict_messages = []
        type_map = {}

        for row in changed_rows:
            type_id_int = row.TypeId.IntegerValue
            target_name = normalize_target_workset_name(row.SelectedTargetWorksetName)

            if type_id_int not in type_map:
                type_map[type_id_int] = []
            type_map[type_id_int].append((row.LinkName, target_name))

        for type_id_int, entries in type_map.items():
            unique_targets = list(set([x[1] for x in entries]))
            if len(unique_targets) > 1:
                lines = []
                for link_name, target_name in entries:
                    lines.append("- {} -> {}".format(link_name, target_name))
                conflict_messages.append("Shared Link Type conflict:\n{}".format("\n".join(lines)))

        if conflict_messages:
            forms.alert(
                "Conflicting target worksets were found for rows that share the same Link Type.\n\n"
                + "\n\n".join(conflict_messages),
                warn_icon=True
            )
            return

        confirm_lines = []
        for row in changed_rows:
            confirm_lines.append("- {} -> {}".format(row.LinkName, row.SelectedTargetWorksetName))

        proceed = forms.alert(
            "Apply workset changes to modified links?\n\n{}".format("\n".join(confirm_lines)),
            yes=True,
            no=True,
            warn_icon=False
        )
        if not proceed:
            return

        updated_instances = 0
        updated_types = 0
        errors = []
        created_now = []
        processed_type_ids = set()

        t = DB.Transaction(doc, "pyMenvic | Link Workset Manager")
        t.Start()

        try:
            for row in changed_rows:
                target_name = normalize_target_workset_name(row.SelectedTargetWorksetName)
                target_ws = self.workset_by_name.get(target_name, None)

                if not target_ws:
                    target_ws, was_created, create_msg = get_or_create_workset_by_name(target_name)
                    if target_ws:
                        self.workset_by_name[target_name] = target_ws
                        if was_created:
                            created_now.append(target_name)
                    else:
                        errors.append("Target workset not found for '{}': {}".format(row.LinkName, create_msg or "Unknown error."))
                        continue

                target_workset_id = target_ws.Id

                try:
                    inst = doc.GetElement(row.InstanceId)
                    if inst:
                        current_inst_ws = get_element_workset_name(inst)
                        if current_inst_ws != target_name:
                            ok, msg = set_element_workset(inst, target_workset_id)
                            if ok:
                                updated_instances += 1
                            else:
                                errors.append("Instance '{}': {}".format(row.LinkName, msg))
                except Exception as ex:
                    errors.append("Instance '{}': {}".format(row.LinkName, safe_str(ex)))

                try:
                    type_id_int = row.TypeId.IntegerValue
                    if type_id_int not in processed_type_ids:
                        typ = doc.GetElement(row.TypeId)
                        if typ:
                            current_type_ws = get_element_workset_name(typ)
                            if current_type_ws != target_name:
                                ok, msg = set_element_workset(typ, target_workset_id)
                                if ok:
                                    updated_types += 1
                                else:
                                    errors.append("Type for '{}': {}".format(row.LinkName, msg))
                        processed_type_ids.add(type_id_int)
                except Exception as ex:
                    errors.append("Type for '{}': {}".format(row.LinkName, safe_str(ex)))

            t.Commit()

        except Exception as ex:
            t.RollBack()
            forms.alert("Transaction failed:\n\n{}".format(safe_str(ex)), warn_icon=True)
            return

        self.refresh_workset_cache()

        try:
            for row in self.rows_all:
                inst = doc.GetElement(row.InstanceId)
                typ = doc.GetElement(row.TypeId)

                new_inst_ws = row.CurrentInstanceWorkset
                new_type_ws = row.CurrentTypeWorkset

                if inst:
                    new_inst_ws = get_element_workset_name(inst)
                if typ:
                    new_type_ws = get_element_workset_name(typ)

                row.refresh_all_after_apply(new_inst_ws, new_type_ws)
        except Exception:
            pass

        self.refresh_grid()

        msg = []
        msg.append("MENVIC | LINK WORKSET MANAGER — APPLY")
        msg.append("")
        msg.append("Summary")
        msg.append("")
        msg.append("• Modified rows detected: {}".format(len(changed_rows)))
        msg.append("• Updated instances: {}".format(updated_instances))
        msg.append("• Updated types: {}".format(updated_types))
        msg.append("• Errors: {}".format(len(errors)))

        if created_now:
            msg.append("")
            msg.append("• Created worksets: {}".format(len(created_now)))
            for name in sorted(set(created_now)):
                msg.append("  - {}".format(name))

        if updated_instances == 0 and updated_types == 0 and len(errors) == 0:
            msg.append("")
            msg.append("No actual model changes were needed.")

        if errors:
            msg.append("")
            msg.append("First errors:")
            for err in errors[:10]:
                msg.append("- {}".format(err.splitlines()[0]))

        forms.alert("\n".join(msg), title="Link Workset Manager", warn_icon=False)
        self.Close()


# ==================================================
# RUN
# ==================================================

LinkWorksetManagerWindow("ui.xaml").ShowDialog()

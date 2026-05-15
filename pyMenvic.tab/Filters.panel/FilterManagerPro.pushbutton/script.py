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

from Autodesk.Revit.DB import View, ElementId, Transaction, Category
from pyrevit import forms, revit, script

doc = revit.doc
XAML_FILE = script.get_bundle_file("filter_manager_pro.xaml")
LOGO_FILE = get_filters_logo_path()


class FilterOption(object):
    def __init__(self, element_id, name):
        self.ElementId = element_id
        self.Name = name


class AuditRow(object):
    def __init__(self, filter_id, original_name, name, categories, vc, tc, duplicate_label):
        self.FilterId = filter_id
        self.OriginalName = original_name
        self.FilterName = name
        self.Categories = categories
        self.ViewCount = vc
        self.TemplateCount = tc
        self.TotalCount = vc + tc
        self.Status = "Used" if self.TotalCount > 0 else "Unused"
        self.Duplicate = duplicate_label or "-"


class RenameRow(object):
    def __init__(self, filter_id, current, proposed):
        self.FilterId = filter_id
        self.CurrentName = current
        self.ProposedName = proposed
        self.Apply = False
        self.Status = "No change"


class ReplaceRow(object):
    def __init__(self, view_id, view_name, kind, templ, hs, ht, se, sv, te, tv):
        self.ViewId = view_id; self.ViewName = view_name; self.ViewKind = kind; self.IsTemplate = templ
        self.HasSource = hs; self.HasTarget = ht; self.SourceEnabled = se; self.SourceVisible = sv; self.TargetEnabled = te; self.TargetVisible = tv
        self.Apply = hs; self.Status = "Ready" if hs else "No source"


class FilterManagerProWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self._load_header_logo(); self.filters = self._collect_filters(); self._rebuild_maps()
        self.audit_rows = ObservableCollection[object](); self.rename_rows = ObservableCollection[object](); self.replace_rows = ObservableCollection[object]()
        self.all_rename_rows = []; self.all_replace_rows = []
        self.AuditGrid.ItemsSource = self.audit_rows; self.RenameGrid.ItemsSource = self.rename_rows; self.ReplaceGrid.ItemsSource = self.replace_rows
        self.SourceComboBox.ItemsSource = self.filter_names; self.TargetComboBox.ItemsSource = self.filter_names
        if self.filter_names: self.SourceComboBox.SelectedIndex = 0; self.TargetComboBox.SelectedIndex = min(1, len(self.filter_names)-1)
        self._load_audit(); self._load_rename_rows(); self._set_reports_status("Ready to export from current tab data.")
        self._set_rename_status("Configure rename options and click Preview."); self._set_replace_status("Select Source and Target, then Preview Usage.")
        self._refresh_active_tab_summary()

    def _rebuild_maps(self):
        self.filter_name_to_option = {f.Name: f for f in self.filters}
        self.filter_id_to_option = {element_id_value(f.ElementId): f for f in self.filters}
        self.filter_names = sorted(self.filter_name_to_option.keys(), key=lambda x: x.lower())

    def _collect_filters(self):
        return [FilterOption(f.Id, element_name(f)) for f in collect_parameter_filters(doc, key_selector=lambda i: element_name(i).lower())]

    def _views(self):
        from Autodesk.Revit.DB import FilteredElementCollector
        out = []
        for v in FilteredElementCollector(doc).OfClass(View):
            try: v.GetFilters(); out.append(v)
            except Exception: pass
        return out

    def _set_text(self, name, v):
        try: getattr(self, name).Text = str(v)
        except Exception: pass

    def _set_header_cards(self, cards):
        cards = cards[:4]
        for i in range(4):
            d = cards[i] if i < len(cards) else None
            self._set_text("HeaderCardLabel{}".format(i+1), d[0] if d else "")
            self._set_text("HeaderCardValue{}".format(i+1), d[1] if d else "")
            try: getattr(self, "HeaderCardBorder{}".format(i+1)).Visibility = Visibility.Visible if d else Visibility.Collapsed
            except Exception: pass

    def _card(self, l, v): return (l, str(v))

    def _refresh_active_tab_summary(self):
        h = "Audit"
        try: h = str(self.MainTabControl.SelectedItem.Header)
        except Exception: pass
        if "Audit" in h:
            used = len([r for r in self.audit_rows if r.TotalCount > 0])
            self._set_header_cards([self._card("FILTERS", len(self.audit_rows)), self._card("USED", used), self._card("UNUSED", len(self.audit_rows) - used), self._card("DUPLICATES", len([r for r in self.audit_rows if r.Duplicate != "-"]))])
        elif "Rename" in h:
            ready = len([r for r in self.rename_rows if r.Apply])
            self._set_header_cards([self._card("ROWS", len(self.rename_rows)), self._card("READY", ready)])
        elif "Replace" in h:
            self._set_header_cards([self._card("PREVIEW", len(self.replace_rows)), self._card("APPLY", len([r for r in self.replace_rows if r.Apply]))])
        else:
            self._set_header_cards([self._card("REPORTS", "CSV")])

    def _load_header_logo(self):
        s = None
        try:
            s = FileStream(LOGO_FILE, FileMode.Open, FileAccess.Read); i = BitmapImage(); i.BeginInit(); i.StreamSource = s; i.CacheOption = BitmapCacheOption.OnLoad; i.EndInit(); i.Freeze(); self.HeaderLogoImage.Source = i
        except Exception: pass
        finally:
            if s:
                try: s.Close()
                except Exception: pass

    def _safe_class_name(self, obj):
        try: return obj.GetType().Name
        except Exception:
            try: return obj.__class__.__name__
            except Exception: return str(type(obj))

    def _category_ids(self, filter_el):
        try: return list(filter_el.GetCategories())
        except Exception:
            try:
                raw = getattr(filter_el, "Categories", None)
                if raw: return list(raw)
            except Exception: pass
        return []

    def _category_name_from_id(self, cid):
        try:
            cat = Category.GetCategory(doc, cid)
            if cat and cat.Name: return cat.Name
        except Exception: pass
        try:
            cat = doc.Settings.Categories.get_Item(cid)
            if cat and cat.Name: return cat.Name
        except Exception: pass
        try:
            cid_value = element_id_value(cid)
            for cat in doc.Settings.Categories:
                try:
                    if element_id_value(cat.Id) == cid_value:
                        return cat.Name
                except Exception: pass
        except Exception: pass
        try: return "CategoryId {}".format(element_id_value(cid))
        except Exception: return str(cid)

    def _get_categories(self, filter_el):
        names = []
        for cid in self._category_ids(filter_el):
            names.append(self._category_name_from_id(cid))
        unique = sorted(set([n for n in names if n]))
        if not unique: return "N/A"
        if len(unique) <= 4: return ", ".join(unique)
        return "{} categories".format(len(unique))

    def _category_signature(self, filter_el):
        values = []
        for cid in self._category_ids(filter_el):
            try: values.append(str(element_id_value(cid)))
            except Exception: values.append(str(cid))
        return "|".join(sorted(values))

    def _extract_rule_parameter_id(self, rule):
        for method_name in ("GetRuleParameter",):
            try: return getattr(rule, method_name)()
            except Exception: pass
        for property_name in ("RuleParameter", "ParameterId"):
            try: return getattr(rule, property_name)
            except Exception: pass
        return None

    def _extract_rule_value(self, rule):
        for method_name in ("GetStringValue", "GetValue", "GetIntegerValue", "GetDoubleValue", "GetElementIdValue"):
            try: return getattr(rule, method_name)()
            except Exception: pass
        return None

    def _rule_signature(self, rule):
        parameter_id = self._extract_rule_parameter_id(rule)
        try: parameter_value = element_id_value(parameter_id)
        except Exception: parameter_value = str(parameter_id)
        value = self._extract_rule_value(rule)
        try: value = element_id_value(value)
        except Exception: pass
        evaluator_name = ""
        try: evaluator_name = self._safe_class_name(rule.Evaluator)
        except Exception: pass
        return "{}|{}|{}|{}".format(self._safe_class_name(rule), parameter_value, value, evaluator_name)

    def _element_filter_signature(self, element_filter):
        if element_filter is None: return ""
        class_name = self._safe_class_name(element_filter)
        if class_name in ("LogicalAndFilter", "LogicalOrFilter"):
            parts = []
            try:
                for child_filter in list(element_filter.GetFilters()):
                    parts.append(self._element_filter_signature(child_filter))
            except Exception: pass
            return "{}({})".format(class_name, ";".join(sorted(parts)))
        if class_name == "ElementParameterFilter":
            rule_parts = []
            try:
                for rule in list(element_filter.GetRules()):
                    rule_parts.append(self._rule_signature(rule))
            except Exception: pass
            return "{}({})".format(class_name, ";".join(sorted(rule_parts)))
        return "{}:{}".format(class_name, str(element_filter))

    def _filter_content_signature(self, filter_el):
        if filter_el is None: return ""
        cat_sig = self._category_signature(filter_el)
        filter_sig = ""
        try: filter_sig = self._element_filter_signature(filter_el.GetElementFilter())
        except Exception: pass
        if not filter_sig:
            rule_parts = []
            try:
                for rule in list(filter_el.GetRules()):
                    rule_parts.append(self._rule_signature(rule))
            except Exception: pass
            filter_sig = "RULES({})".format(";".join(sorted(rule_parts)))
        if not cat_sig and not filter_sig: return ""
        return "CATS[{}] FILTER[{}]".format(cat_sig, filter_sig)

    def _load_audit(self):
        self.audit_rows.Clear(); views = self._views(); rows = []; sig_groups = {}
        for f in self.filters:
            fid = element_id_value(f.ElementId); filter_el = doc.GetElement(f.ElementId)
            cats = self._get_categories(filter_el); sig = self._filter_content_signature(filter_el); vc = 0; tc = 0
            for v in views:
                try: ids = [element_id_value(x) for x in v.GetFilters()]
                except Exception: continue
                if fid in ids: tc += 1 if v.IsTemplate else 0; vc += 0 if v.IsTemplate else 1
            rows.append((fid, f.Name, cats, vc, tc, sig))
            if sig: sig_groups.setdefault(sig, []).append(fid)
        for row in rows:
            dup = "Exact" if row[5] and len(sig_groups.get(row[5], [])) > 1 else "-"
            self.audit_rows.Add(AuditRow(row[0], row[1], row[1], row[2], row[3], row[4], dup))
        self._refresh_active_tab_summary()

    def ApplyAuditChangesButton_Click(self, s, a):
        forms.alert("Audit is read-only. Use Rename / Standardize to rename filters.", title="Filter Manager Pro", exitscript=False)

    def _load_rename_rows(self):
        self.all_rename_rows = [RenameRow(element_id_value(f.ElementId), f.Name, f.Name) for f in self.filters]
        self._filter_rename_rows()

    def _filter_rename_rows(self):
        term = (self.RenameSearchTextBox.Text or "").strip().lower(); self.rename_rows.Clear()
        for r in self.all_rename_rows:
            if term and term not in r.CurrentName.lower() and term not in r.ProposedName.lower(): continue
            self.rename_rows.Add(r)
        self._refresh_active_tab_summary()

    def PreviewRenameButton_Click(self, s, a):
        f = (self.FindTextBox.Text or ""); rep = (self.ReplaceTextBox.Text or ""); pre = (self.RenamePrefixTextBox.Text or ""); suf = (self.RenameSuffixTextBox.Text or ""); up = self.UppercaseCheckBox.IsChecked
        for r in self.all_rename_rows:
            p = r.CurrentName
            if f: p = p.replace(f, rep)
            p = "{}{}{}".format(pre, p, suf)
            if up: p = p.upper()
            r.ProposedName = p; r.Apply = (r.CurrentName != r.ProposedName); r.Status = "Ready" if r.Apply else "No change"
        self._filter_rename_rows(); self._set_rename_status("Preview ready.")

    def ResetRenameRowButton_Click(self, s, a):
        r = self.RenameGrid.SelectedItem
        if r: r.ProposedName = r.CurrentName; r.Apply = False; r.Status = "Reset"; self.RenameGrid.Items.Refresh(); self._refresh_active_tab_summary()

    def RenameSearchTextBox_TextChanged(self, s, a): self._filter_rename_rows()
    def ReplaceSearchTextBox_TextChanged(self, s, a): self._filter_replace_rows()
    def RefreshAuditButton_Click(self, s, a): self._load_audit()
    def MainTabControl_SelectionChanged(self, s, a): self._refresh_active_tab_summary()

    # existing apply/replace/export methods preserved
    def ApplyRenameButton_Click(self,s,a):
        rows=[r for r in self.all_rename_rows if r.Apply]
        if not rows: self._set_rename_status("Nothing selected to rename."); return
        names=[(r.ProposedName or "").strip() for r in rows]
        if "" in names: self._set_rename_status("Validation failed: empty proposed names."); return
        if len(set([n.lower() for n in names])) != len(names): self._set_rename_status("Validation failed: duplicate proposed names."); return
        existing=set([f.Name.lower() for f in self.filters if element_id_value(f.ElementId) not in [r.FilterId for r in rows]])
        for n in names:
            if n.lower() in existing: self._set_rename_status("Validation failed: conflicts with existing filters."); return
        ok=0; fail=0; tx=Transaction(doc,"Filter Manager Pro - Rename"); tx.Start()
        try:
            for r in rows:
                try: doc.GetElement(ElementId(r.FilterId)).Name=r.ProposedName; ok+=1
                except Exception: fail+=1
            tx.Commit()
        except Exception:
            try: tx.RollBack()
            except Exception: pass
            fail=len(rows)
        self.filters=self._collect_filters(); self._rebuild_maps(); self.SourceComboBox.ItemsSource=self.filter_names; self.TargetComboBox.ItemsSource=self.filter_names
        self._load_audit(); self._load_rename_rows(); self._set_rename_status("Apply Rename complete. Renamed: {} | Failed: {}".format(ok,fail))

    def PreviewReplaceButton_Click(self,s,a):
        self.replace_rows.Clear(); self.all_replace_rows=[]
        src=self.filter_name_to_option.get(self.SourceComboBox.SelectedItem); tgt=self.filter_name_to_option.get(self.TargetComboBox.SelectedItem)
        if not src or not tgt or element_id_value(src.ElementId)==element_id_value(tgt.ElementId): self._set_replace_status("Select different source and target filters."); return
        svid=element_id_value(src.ElementId); tvid=element_id_value(tgt.ElementId)
        inc_views=self.IncludeViewsCheckBox.IsChecked; inc_t=self.IncludeTemplatesCheckBox.IsChecked
        for v in self._views():
            if (v.IsTemplate and not inc_t) or ((not v.IsTemplate) and not inc_views): continue
            try: ids=list(v.GetFilters()); vals=[element_id_value(x) for x in ids]
            except Exception: continue
            hs=svid in vals; ht=tvid in vals
            if not hs and not ht: continue
            ge=getattr(v,"GetIsFilterEnabled",None); se=te=None
            if callable(ge):
                try: se=ge(src.ElementId)
                except Exception: pass
                try: te=ge(tgt.ElementId)
                except Exception: pass
            sv=tv=None
            try: sv=v.GetFilterVisibility(src.ElementId)
            except Exception: pass
            try: tv=v.GetFilterVisibility(tgt.ElementId)
            except Exception: pass
            row=ReplaceRow(element_id_value(v.Id),element_name(v),str(v.ViewType),v.IsTemplate,hs,ht,se,sv,te,tv)
            self.all_replace_rows.append(row)
        self._filter_replace_rows(); self._set_replace_status("Preview ready: {} rows.".format(len(self.all_replace_rows)))

    def _filter_replace_rows(self):
        term=(self.ReplaceSearchTextBox.Text or "").strip().lower(); self.replace_rows.Clear()
        for r in self.all_replace_rows:
            if term and term not in r.ViewName.lower() and term not in r.ViewKind.lower(): continue
            self.replace_rows.Add(r)
        self._refresh_active_tab_summary()

    def ApplyReplaceButton_Click(self,s,a):
        src=self.filter_name_to_option.get(self.SourceComboBox.SelectedItem); tgt=self.filter_name_to_option.get(self.TargetComboBox.SelectedItem)
        if not src or not tgt: self._set_replace_status("Missing source/target."); return
        rows=[r for r in self.all_replace_rows if r.Apply]
        if not rows: self._set_replace_status("No rows selected."); return
        merge=self.MergeExistingCheckBox.IsChecked; copyv=self.CopyVisibilityCheckBox.IsChecked; copye=self.CopyEnabledCheckBox.IsChecked
        ok=0;sk=0;fl=0; tx=Transaction(doc,"Filter Manager Pro - Replace"); tx.Start()
        try:
            for r in rows:
                v=doc.GetElement(ElementId(r.ViewId))
                if not v: sk+=1; continue
                try: vals=[element_id_value(x) for x in list(v.GetFilters())]
                except Exception: sk+=1; continue
                hs=element_id_value(src.ElementId) in vals; ht=element_id_value(tgt.ElementId) in vals
                if not hs: sk+=1; continue
                try:
                    o=v.GetFilterOverrides(src.ElementId)
                    if (not ht): v.AddFilter(tgt.ElementId)
                    elif not merge: sk+=1; continue
                    v.SetFilterOverrides(tgt.ElementId,o)
                    if copyv: v.SetFilterVisibility(tgt.ElementId, v.GetFilterVisibility(src.ElementId))
                    if copye:
                        ge=getattr(v,"GetIsFilterEnabled",None); se=getattr(v,"SetIsFilterEnabled",None)
                        if callable(ge) and callable(se): se(tgt.ElementId, ge(src.ElementId))
                    v.RemoveFilter(src.ElementId); ok+=1
                except Exception: fl+=1
            tx.Commit()
        except Exception:
            try: tx.RollBack()
            except Exception: pass
            fl+=1
        self._load_audit(); self.PreviewReplaceButton_Click(None,None); self._set_replace_status("Apply Replace complete. Updated: {} | Skipped: {} | Failed: {}".format(ok,sk,fl))

    def ExportAuditCsvButton_Click(self,s,a): self._export_csv("audit_summary.csv", ["Filter","Categories","Views","Templates","Total","Status","Duplicate"], self.audit_rows, lambda r:[r.FilterName,r.Categories,r.ViewCount,r.TemplateCount,r.TotalCount,r.Status,r.Duplicate])
    def ExportUnusedCsvButton_Click(self,s,a):
        rows=[r for r in self.audit_rows if r.TotalCount==0]; self._export_csv("unused_filters.csv", ["Filter","Categories"], rows, lambda r:[r.FilterName,r.Categories])
    def ExportReplaceCsvButton_Click(self,s,a): self._export_csv("replace_preview.csv", ["View","Type","Template","Source","Target","Apply"], self.all_replace_rows, lambda r:[r.ViewName,r.ViewKind,r.IsTemplate,r.HasSource,r.HasTarget,r.Apply])
    def _export_csv(self, filename, header, rows, rowf):
        path=forms.save_file(file_ext='csv', default_name=filename)
        if not path: return
        import csv
        try:
            with open(path,'wb') as f:
                w=csv.writer(f); w.writerow(header)
                for r in rows: w.writerow([str(x) for x in rowf(r)])
            self._set_reports_status("Exported: {}".format(path))
        except Exception as ex: self._set_reports_status("Export failed: {}".format(ex))

    def _set_rename_status(self,t): self.RenameStatusTextBlock.Text=t
    def _set_replace_status(self,t): self.ReplaceStatusTextBlock.Text=t
    def _set_reports_status(self,t): self.ReportsStatusTextBlock.Text=t

FilterManagerProWindow().ShowDialog()

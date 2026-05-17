# -*- coding: utf-8 -*-

from System.Windows import Visibility
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption


class FilterManagerUIHelpers(object):
    def _set_text(self, name, v):
        try:
            getattr(self, name).Text = str(v)
        except Exception:
            pass

    def _set_header_cards(self, cards):
        cards = cards[:4]
        for i in range(4):
            d = cards[i] if i < len(cards) else None
            self._set_text("HeaderCardLabel{}".format(i + 1), d[0] if d else "")
            self._set_text("HeaderCardValue{}".format(i + 1), d[1] if d else "")
            try:
                getattr(self, "HeaderCardBorder{}".format(i + 1)).Visibility = Visibility.Visible if d else Visibility.Collapsed
            except Exception:
                pass

    def _card(self, l, v):
        return (l, str(v))

    def _refresh_active_tab_summary(self):
        h = "Audit"
        try:
            h = str(self.MainTabControl.SelectedItem.Header)
        except Exception:
            pass
        if "Audit" in h:
            used = len([r for r in self.all_audit_rows if r.TotalCount > 0])
            self._set_header_cards([
                self._card("FILTERS", len(self.all_audit_rows)),
                self._card("VISIBLE", len(self.audit_rows)),
                self._card("UNUSED", len(self.all_audit_rows) - used),
                self._card("DUP. SETS", self._duplicate_group_count(self.all_audit_rows))
            ])
        elif "Rename" in h:
            ready = len([r for r in self.rename_rows if r.Apply and r.Status == "Ready to Rename"])
            self._set_header_cards([self._card("ROWS", len(self.rename_rows)), self._card("READY", ready)])
        elif "Replace" in h:
            self._set_header_cards([self._card("PREVIEW", len(self.replace_rows)), self._card("APPLY", len([r for r in self.replace_rows if r.Apply and r.Status == "Ready to Replace"]))])
        else:
            self._set_header_cards([self._card("REPORTS", "CSV")])

    def _load_header_logo(self):
        s = None
        try:
            s = FileStream(self._logo_file, FileMode.Open, FileAccess.Read)
            i = BitmapImage()
            i.BeginInit()
            i.StreamSource = s
            i.CacheOption = BitmapCacheOption.OnLoad
            i.EndInit()
            i.Freeze()
            self.HeaderLogoImage.Source = i
        except Exception:
            pass
        finally:
            if s:
                try:
                    s.Close()
                except Exception:
                    pass

    def _set_audit_status(self, msg):
        self._set_text("AuditStatusText", msg)

    def _set_rename_status(self, msg):
        self._set_text("RenameStatusText", msg)

    def _set_replace_status(self, msg):
        self._set_text("ReplaceStatusText", msg)

    def _set_reports_status(self, msg):
        self._set_text("ReportsStatusText", msg)

    def _set_audit_details_columns(self, cols):
        parts = []
        for p, c in cols:
            parts.append("{} ({})".format(p, c))
        t = ", ".join(parts) if parts else "None"
        self._set_text("AuditDetailsColumnsText", t)

    def _set_audit_details(self, row):
        self._set_text("AuditDetailsFilterText", row.FilterName)
        self._set_text("AuditDetailsUsageText", "Views: {} | Templates: {} | Total: {}".format(row.ViewCount, row.TemplateCount, row.TotalCount))
        self._set_text("AuditDetailsCategoriesText", row.CategorySummary)
        self._set_text("AuditDetailsRuleCountText", str(row.RuleCount))
        self._set_text("AuditDetailsRuleTypesText", row.RuleTypeSummary)
        self._set_text("AuditDetailsDuplicateText", row.DuplicateType)
        self._set_text("AuditDetailsDuplicateGroupText", row.DuplicateGroup)
        self._set_text("AuditDetailsSourceText", row.SourceType)
        self._set_text("AuditDetailsRulesText", row.RuleSummary)
        self._set_audit_details_columns(row.ColumnMentions)

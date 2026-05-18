# -*- coding: utf-8 -*-


class FilterOption(object):
    def __init__(self, element_id, name):
        self.ElementId = element_id
        self.Name = name


class AuditRow(object):
    def __init__(self, filter_id, original_name, name, categories, category_names, vc, tc, duplicate_type, duplicate_group):
        self.FilterId = filter_id
        self.OriginalName = original_name
        self.FilterName = name
        self.Categories = categories
        self.CategoryNames = category_names or []
        self.ViewCount = vc
        self.TemplateCount = tc
        self.TotalCount = vc + tc
        self.Status = "Used" if self.TotalCount > 0 else "Unused"
        self.DuplicateType = duplicate_type or "Not duplicate"
        self.DuplicateGroup = duplicate_group or "-"
        self.Duplicate = self.DuplicateType
        self.Purge = False


class UsageViewRow(object):
    def __init__(self, view_id, name):
        self.ViewId = view_id
        self.Name = name


class RenameRow(object):
    def __init__(self, filter_id, current, proposed):
        self.FilterId = filter_id
        self.CurrentName = current
        self.ProposedName = proposed
        self.Apply = False
        self.Status = "No change"


class ReplaceRow(object):
    def __init__(self, view_id, view_name, kind, templ, hs, ht, se, sv, te, tv):
        self.ViewId = view_id
        self.ViewName = view_name
        self.ViewKind = kind
        self.IsTemplate = templ
        self.HasSource = hs
        self.HasTarget = ht
        self.SourceEnabled = se
        self.SourceVisible = sv
        self.TargetEnabled = te
        self.TargetVisible = tv
        self.Apply = hs
        if hs:
            self.Status = "Ready to Replace"
        elif ht:
            self.Status = "Already Has Target"
        else:
            self.Status = "No Source"

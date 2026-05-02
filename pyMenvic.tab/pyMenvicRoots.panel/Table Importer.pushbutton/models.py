# -*- coding: utf-8 -*-

class TableEntry(object):
    def __init__(
        self,
        selected=False,
        status="Not Created",
        view_name="",
        auto_sync=False,
        last_modified="",
        worksheet="",
        region="Used Range",
        view_type="Drafting View",
        view_scale="1",
        file_path="",
        revit_view_id=None
    ):
        self.Selected = selected
        self.Status = status
        self.ViewName = view_name
        self.AutoSync = auto_sync
        self.LastModified = last_modified
        self.Worksheet = worksheet
        self.Region = region
        self.ViewType = view_type
        self.ViewScale = view_scale
        self.FilePath = file_path
        self.RevitViewId = revit_view_id

    def to_dict(self):
        return {
            "selected": self.Selected,
            "status": self.Status,
            "view_name": self.ViewName,
            "auto_sync": self.AutoSync,
            "last_modified": self.LastModified,
            "worksheet": self.Worksheet,
            "region": self.Region,
            "view_type": self.ViewType,
            "view_scale": self.ViewScale,
            "file_path": self.FilePath,
            "revit_view_id": self.RevitViewId,
        }

    @staticmethod
    def from_dict(data):
        return TableEntry(
            selected=data.get("selected", False),
            status=data.get("status", "Not Created"),
            view_name=data.get("view_name", ""),
            auto_sync=data.get("auto_sync", False),
            last_modified=data.get("last_modified", ""),
            worksheet=data.get("worksheet", ""),
            region=data.get("region", "Used Range"),
            view_type=data.get("view_type", "Drafting View"),
            view_scale=data.get("view_scale", "1"),
            file_path=data.get("file_path", ""),
            revit_view_id=data.get("revit_view_id", None)
        )
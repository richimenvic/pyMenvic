# -*- coding: utf-8 -*-

import clr
clr.AddReference("WindowsBase")

from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs


class TableEntry(INotifyPropertyChanged):
    def __init__(
        self,
        selected=False,
        status="Not Created",
        view_name="",
        auto_sync=False,
        last_modified="",
        worksheet="",
        region="Used Range",
        region_options=None,
        view_type="Drafting View",
        view_scale="1",
        file_path="",
        revit_view_id=None
    ):
        self._selected = selected
        self._status = status
        self._view_name = view_name
        self._auto_sync = auto_sync
        self._last_modified = last_modified
        self._worksheet = worksheet
        self._region = region
        self._region_options = region_options if region_options else []
        self._view_type = view_type
        self._view_scale = view_scale
        self._file_path = file_path
        self._revit_view_id = revit_view_id
        self._handlers = []

    def add_PropertyChanged(self, handler):
        self._handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._handlers:
            self._handlers.remove(handler)

    def _notify(self, prop_name):
        args = PropertyChangedEventArgs(prop_name)
        for h in self._handlers:
            try:
                h(self, args)
            except Exception:
                pass

    @property
    def Selected(self):
        return self._selected

    @Selected.setter
    def Selected(self, value):
        if self._selected != value:
            self._selected = value
            self._notify("Selected")

    @property
    def Status(self):
        return self._status

    @Status.setter
    def Status(self, value):
        if self._status != value:
            self._status = value
            self._notify("Status")

    @property
    def ViewName(self):
        return self._view_name

    @ViewName.setter
    def ViewName(self, value):
        if self._view_name != value:
            self._view_name = value
            self._notify("ViewName")

    @property
    def AutoSync(self):
        return self._auto_sync

    @AutoSync.setter
    def AutoSync(self, value):
        if self._auto_sync != value:
            self._auto_sync = value
            self._notify("AutoSync")

    @property
    def LastModified(self):
        return self._last_modified

    @LastModified.setter
    def LastModified(self, value):
        if self._last_modified != value:
            self._last_modified = value
            self._notify("LastModified")

    @property
    def Worksheet(self):
        return self._worksheet

    @Worksheet.setter
    def Worksheet(self, value):
        if self._worksheet != value:
            self._worksheet = value
            self._notify("Worksheet")

    @property
    def Region(self):
        return self._region

    @Region.setter
    def Region(self, value):
        if self._region != value:
            self._region = value
            self._notify("Region")


    @property
    def RegionOptions(self):
        return self._region_options

    @RegionOptions.setter
    def RegionOptions(self, value):
        if value is None:
            value = []
        self._region_options = value
        self._notify("RegionOptions")

    @property
    def ViewType(self):
        return self._view_type

    @ViewType.setter
    def ViewType(self, value):
        if self._view_type != value:
            self._view_type = value
            self._notify("ViewType")

    @property
    def ViewScale(self):
        return self._view_scale

    @ViewScale.setter
    def ViewScale(self, value):
        if self._view_scale != value:
            self._view_scale = value
            self._notify("ViewScale")

    @property
    def FilePath(self):
        return self._file_path

    @FilePath.setter
    def FilePath(self, value):
        if self._file_path != value:
            self._file_path = value
            self._notify("FilePath")

    @property
    def RevitViewId(self):
        return self._revit_view_id

    @RevitViewId.setter
    def RevitViewId(self, value):
        if self._revit_view_id != value:
            self._revit_view_id = value
            self._notify("RevitViewId")

    def to_dict(self):
        return {
            "selected": self._selected,
            "status": self._status,
            "view_name": self._view_name,
            "auto_sync": self._auto_sync,
            "last_modified": self._last_modified,
            "worksheet": self._worksheet,
            "region": self._region,
            "region_options": list(self._region_options) if self._region_options else [],
            "view_type": self._view_type,
            "view_scale": self._view_scale,
            "file_path": self._file_path,
            "revit_view_id": self._revit_view_id,
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
            region_options=data.get("region_options", None),
            view_type=data.get("view_type", "Drafting View"),
            view_scale=data.get("view_scale", "1"),
            file_path=data.get("file_path", ""),
            revit_view_id=data.get("revit_view_id", None)
        )

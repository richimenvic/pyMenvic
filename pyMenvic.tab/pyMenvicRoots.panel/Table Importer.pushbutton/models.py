# -*- coding: utf-8 -*-

import clr
clr.AddReference("WindowsBase")

from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs


def _get(data, key, default=None):
    try:
        if key in data:
            return data.get(key, default)
    except Exception:
        pass
    return default


class TableEntry(INotifyPropertyChanged):
    def __init__(
        self,
        selected=False,
        status="Ready to Update",
        source="",
        import_type="Excel Link",
        view_name="",
        dpi="150",
        auto_sync=False,
        black_and_white=True,
        last_modified="",
        worksheet="",
        region="Full Worksheet Used Range",
        region_options=None,
        view_type="Drafting View",
        view_scale="1",
        file_path="",
        path_mode="Absolute",
        revit_view_id=None,
        table_entry_uid=None,
        created_element_ids=None
    ):
        self._selected = selected
        self._status = status
        self._source = source
        self._import_type = import_type
        self._view_name = view_name
        self._dpi = dpi
        self._auto_sync = auto_sync
        self._black_and_white = black_and_white
        self._last_modified = last_modified
        self._worksheet = worksheet
        self._region = region
        self._region_options = region_options if region_options else []
        self._view_type = view_type
        self._view_scale = view_scale
        self._file_path = file_path
        self._path_mode = path_mode
        self._revit_view_id = revit_view_id
        self._table_entry_uid = table_entry_uid
        self._created_element_ids = list(created_element_ids or [])
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

    def _set(self, private_name, prop_name, value):
        if getattr(self, private_name) != value:
            setattr(self, private_name, value)
            self._notify(prop_name)

    @property
    def Selected(self):
        return self._selected

    @Selected.setter
    def Selected(self, value):
        self._set("_selected", "Selected", value)

    @property
    def Status(self):
        return self._status

    @Status.setter
    def Status(self, value):
        self._set("_status", "Status", value)

    @property
    def Source(self):
        return self._source

    @Source.setter
    def Source(self, value):
        self._set("_source", "Source", value)

    @property
    def ImportType(self):
        return self._import_type

    @ImportType.setter
    def ImportType(self, value):
        self._set("_import_type", "ImportType", value)

    @property
    def ViewName(self):
        return self._view_name

    @ViewName.setter
    def ViewName(self, value):
        self._set("_view_name", "ViewName", value)

    @property
    def DPI(self):
        return self._dpi

    @DPI.setter
    def DPI(self, value):
        self._set("_dpi", "DPI", value)

    @property
    def AutoSync(self):
        return self._auto_sync

    @AutoSync.setter
    def AutoSync(self, value):
        self._set("_auto_sync", "AutoSync", value)

    @property
    def BlackAndWhite(self):
        return self._black_and_white

    @BlackAndWhite.setter
    def BlackAndWhite(self, value):
        self._set("_black_and_white", "BlackAndWhite", value)

    @property
    def LastModified(self):
        return self._last_modified

    @LastModified.setter
    def LastModified(self, value):
        self._set("_last_modified", "LastModified", value)

    @property
    def Worksheet(self):
        return self._worksheet

    @Worksheet.setter
    def Worksheet(self, value):
        self._set("_worksheet", "Worksheet", value)

    @property
    def Region(self):
        return self._region

    @Region.setter
    def Region(self, value):
        self._set("_region", "Region", value)

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
        self._set("_view_type", "ViewType", value)

    @property
    def ViewScale(self):
        return self._view_scale

    @ViewScale.setter
    def ViewScale(self, value):
        self._set("_view_scale", "ViewScale", value)

    @property
    def FilePath(self):
        return self._file_path

    @FilePath.setter
    def FilePath(self, value):
        self._set("_file_path", "FilePath", value)

    @property
    def PathMode(self):
        return self._path_mode

    @PathMode.setter
    def PathMode(self, value):
        self._set("_path_mode", "PathMode", value)

    @property
    def RevitViewId(self):
        return self._revit_view_id

    @RevitViewId.setter
    def RevitViewId(self, value):
        self._set("_revit_view_id", "RevitViewId", value)

    @property
    def TableEntryUid(self):
        return self._table_entry_uid

    @TableEntryUid.setter
    def TableEntryUid(self, value):
        self._set("_table_entry_uid", "TableEntryUid", value)

    @property
    def CreatedElementIds(self):
        return self._created_element_ids

    @CreatedElementIds.setter
    def CreatedElementIds(self, value):
        if value is None:
            value = []
        self._created_element_ids = list(value)
        self._notify("CreatedElementIds")

    def to_dict(self):
        return {
            "selected": self._selected,
            "status": self._status,
            "source": self._source,
            "import_type": self._import_type,
            "view_name": self._view_name,
            "dpi": self._dpi,
            "auto_sync": self._auto_sync,
            "black_and_white": self._black_and_white,
            "last_modified": self._last_modified,
            "worksheet": self._worksheet,
            "region": self._region,
            "region_options": list(self._region_options) if self._region_options else [],
            "view_type": self._view_type,
            "view_scale": self._view_scale,
            "file_path": self._file_path,
            "path_mode": self._path_mode,
            "revit_view_id": self._revit_view_id,
            "table_entry_uid": self._table_entry_uid,
            "created_element_ids": list(self._created_element_ids) if self._created_element_ids else [],
        }

    @staticmethod
    def from_dict(data):
        file_path = _get(data, "file_path", "")
        source = _get(data, "source", "")
        if not source:
            try:
                import os
                source = os.path.basename(file_path)
            except Exception:
                source = ""
        return TableEntry(
            selected=_get(data, "selected", False),
            status=_get(data, "status", "Ready to Update"),
            source=source,
            import_type=_get(data, "import_type", _get(data, "type", "Excel Link")),
            view_name=_get(data, "view_name", ""),
            dpi=_get(data, "dpi", _get(data, "DPI", "150")),
            auto_sync=_get(data, "auto_sync", False),
            black_and_white=_get(data, "black_and_white", _get(data, "black_white", True)),
            last_modified=_get(data, "last_modified", ""),
            worksheet=_get(data, "worksheet", ""),
            region=_get(data, "region", "Full Worksheet Used Range"),
            region_options=_get(data, "region_options", None),
            view_type=_get(data, "view_type", "Drafting View"),
            view_scale=_get(data, "view_scale", "1"),
            file_path=file_path,
            path_mode=_get(data, "path_mode", "Absolute"),
            revit_view_id=_get(data, "revit_view_id", None),
            table_entry_uid=_get(data, "table_entry_uid", None),
            created_element_ids=_get(data, "created_element_ids", None)
        )

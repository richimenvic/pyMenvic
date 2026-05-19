# -*- coding: utf-8 -*-

import os
import json
import codecs
from models import TableEntry


def safe_unicode(value):
    if value is None:
        return u""
    try:
        if isinstance(value, unicode):
            return value
    except Exception:
        pass
    try:
        if isinstance(value, str):
            for enc in ("utf-8", "cp1252", "latin-1"):
                try:
                    return value.decode(enc)
                except Exception:
                    pass
            return value.decode("utf-8", "replace")
    except Exception:
        pass
    try:
        return unicode(value)
    except Exception:
        try:
            return unicode(value.ToString())
        except Exception:
            return u""


def sanitize_json_value(value):
    if isinstance(value, dict):
        clean = {}
        for k, v in value.items():
            clean[safe_unicode(k)] = sanitize_json_value(v)
        return clean
    if isinstance(value, (list, tuple)):
        return [sanitize_json_value(x) for x in value]
    if isinstance(value, bool) or value is None:
        return value
    try:
        if isinstance(value, int) or isinstance(value, long):
            return value
    except Exception:
        pass
    return safe_unicode(value)


def get_storage_path():
    appdata = os.getenv("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(appdata, "pyMenvic", "TableImporter")
    if not os.path.exists(folder):
        os.makedirs(folder)
    return os.path.join(folder, "table_importer_data.json")


def _read_json(path):
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with codecs.open(path, "r", enc) as f:
                return json.load(f)
        except Exception:
            pass
    return []


def load_entries():
    path = get_storage_path()
    if not os.path.exists(path):
        return []
    try:
        data = _read_json(path)
        if isinstance(data, dict):
            data = data.get("entries", [])
        result = []
        for item in data:
            try:
                result.append(TableEntry.from_dict(sanitize_json_value(item)))
            except Exception:
                pass
        return result
    except Exception:
        return []


def save_entries(entries):
    path = get_storage_path()
    data = []
    for entry in entries:
        data.append(sanitize_json_value(entry.to_dict()))
    with codecs.open(path, "w", "utf-8") as f:
        f.write(json.dumps(data, indent=4, ensure_ascii=False))

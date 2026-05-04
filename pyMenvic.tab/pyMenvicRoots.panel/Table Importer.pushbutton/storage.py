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


def safe_ascii_text(value):
    text = safe_unicode(value)
    if not text:
        return u""
    try:
        import unicodedata
        text = unicodedata.normalize('NFKD', text)
        text = u''.join([c for c in text if not unicodedata.combining(c)])
    except Exception:
        pass
    replacements = {u"Ó":u"O", u"ó":u"o", u"Ñ":u"N", u"ñ":u"n", u"Á":u"A", u"á":u"a", u"É":u"E", u"é":u"e", u"Í":u"I", u"í":u"i", u"Ú":u"U", u"ú":u"u"}
    for a,b in replacements.items():
        text = text.replace(a,b)
    cleaned=[]
    for ch in text:
        try:
            if 32 <= ord(ch) <= 126:
                cleaned.append(ch)
        except Exception:
            pass
    return u"".join(cleaned).strip()


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
        # Keep integers as integers for RevitViewId if needed.
        if isinstance(value, int) or isinstance(value, long):
            return value
    except Exception:
        pass
    return safe_unicode(value)


def get_storage_path():
    appdata = os.getenv("APPDATA")
    folder = os.path.join(appdata, "pyMenvic", "TableImporter")

    if not os.path.exists(folder):
        os.makedirs(folder)

    return os.path.join(folder, "table_importer_data.json")


def load_entries():
    path = get_storage_path()

    if not os.path.exists(path):
        return []

    try:
        try:
            with codecs.open(path, "r", "utf-8") as f:
                data = json.load(f)
        except Exception:
            # Compatibility with older files saved using Windows ANSI encoding.
            with codecs.open(path, "r", "cp1252") as f:
                data = json.load(f)

        return [TableEntry.from_dict(sanitize_json_value(item)) for item in data]

    except Exception:
        return []


def save_entries(entries):
    path = get_storage_path()

    data = []
    for entry in entries:
        data.append(sanitize_json_value(entry.to_dict()))

    # ensure_ascii=True keeps the JSON file ASCII-safe for IronPython/Windows code pages.
    with codecs.open(path, "w", "utf-8") as f:
        f.write(json.dumps(data, indent=4, ensure_ascii=True))

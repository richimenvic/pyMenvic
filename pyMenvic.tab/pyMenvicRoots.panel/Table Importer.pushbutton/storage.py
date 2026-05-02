# -*- coding: utf-8 -*-

import os
import json
from models import TableEntry


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
        with open(path, "r") as f:
            data = json.load(f)

        return [TableEntry.from_dict(item) for item in data]

    except Exception:
        return []


def save_entries(entries):
    path = get_storage_path()

    data = []
    for entry in entries:
        data.append(entry.to_dict())

    with open(path, "w") as f:
        json.dump(data, f, indent=4)
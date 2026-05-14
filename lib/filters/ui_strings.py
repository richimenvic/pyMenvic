# -*- coding: utf-8 -*-


def load_ui_strings(strings_file):
    data = {}
    try:
        execfile(strings_file, data)
        language = data.get("LANGUAGE", "en")
        return data.get("STRINGS", {}).get(language, {})
    except Exception:
        return {}


def ui_text(text, strings_map):
    return strings_map.get(text, text)

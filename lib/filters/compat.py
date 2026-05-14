# -*- coding: utf-8 -*-

def element_id_value(element_id):
    """Return a stable integer value for ElementId across Revit versions."""
    if element_id is None:
        return None
    for attr_name in ("Value", "IntegerValue"):
        try:
            return int(getattr(element_id, attr_name))
        except Exception:
            pass
    try:
        return int(element_id)
    except Exception:
        return str(element_id)

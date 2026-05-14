# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import Element


def element_name(element):
    try:
        return Element.Name.GetValue(element)
    except Exception:
        try:
            return element.Name
        except Exception:
            return ""

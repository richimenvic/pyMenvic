# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import FilteredElementCollector, ParameterFilterElement


def collect_parameter_filters(doc, key_selector):
    filters = list(FilteredElementCollector(doc).OfClass(ParameterFilterElement))
    return sorted(filters, key=key_selector)

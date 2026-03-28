# -*- coding: utf-8 -*-

__title__ = {
    "en_us": "Drafting Audit View",
    "es_es": "Auditoría de Dibujo"
}

__tip__ = {
    "en_us": "Create or open the audit drafting view.",
    "es_es": "Crea o abre la vista de diseño de auditoría."
}

__author__ = "Ricardo J. Mendieta"

__doc__ = """
DRAFTING AUDIT VIEW / AUDITORÍA DE DIBUJO
_____________________________________________________________________

(EN)
Creates or opens the drafting view "AUDIT_DRAFTING_VIEW".
This view is intended as a workspace for audit graphics, 
reports, or diagnostic elements used by pyMenvic tools.

If the view already exists it will simply be opened.
If it does not exist, the tool will create it automatically.

(ES)
Crea o abre la vista de diseño "AUDIT_DRAFTING_VIEW".
Esta vista está pensada como un espacio de trabajo para gráficos 
de auditoría, informes o elementos de diagnóstico de pyMenvic.

Si la vista ya existe, simplemente se abrirá.
Si no existe, la herramienta la creará automáticamente.
_____________________________________________________________________

Author: Ricardo J. Mendieta
Version: 1.0
"""

from pyrevit import revit, DB, forms, script

from pyrevit import revit, DB, forms, script

doc = revit.doc
uidoc = revit.uidoc

nombre_vista = "AUDIT_DRAFTING_VIEW"


# ==================================================
# Buscar si la vista ya existe
# ==================================================

vista_audit = None

vistas_dibujo = DB.FilteredElementCollector(doc).OfClass(DB.ViewDrafting)

for v in vistas_dibujo:
    if v.Name == nombre_vista:
        vista_audit = v
        break


# ==================================================
# Crear vista si no existe
# ==================================================

if not vista_audit:

    tipos_vista = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
    id_tipo_drafting = None

    for tv in tipos_vista:
        if tv.ViewFamily == DB.ViewFamily.Drafting:
            id_tipo_drafting = tv.Id
            break

    if id_tipo_drafting:

        with revit.Transaction("Create Audit View"):

            vista_audit = DB.ViewDrafting.Create(doc, id_tipo_drafting)
            vista_audit.Name = nombre_vista
            vista_audit.Scale = 1

            param_disc = vista_audit.get_Parameter(DB.BuiltInParameter.VIEW_DISCIPLINE)

            if param_disc and not param_disc.IsReadOnly:
                param_disc.Set(4095)

    else:
        print("Error: Drafting View type not found.")


# ==================================================
# Abrir vista
# ==================================================

if vista_audit:
    uidoc.ActiveView = vista_audit
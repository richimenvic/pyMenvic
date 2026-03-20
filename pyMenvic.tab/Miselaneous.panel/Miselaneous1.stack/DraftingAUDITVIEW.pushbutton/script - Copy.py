# -*- coding: utf-8 -*-
from pyrevit import revit, DB

doc = revit.doc
uidoc = revit.uidoc
nombre_vista = "AUDIT_MENVIC"

# 1. Buscar si la vista ya existe
vistas_dibujo = DB.FilteredElementCollector(doc).OfClass(DB.ViewDrafting)
vista_audit = None

for v in vistas_dibujo:
    if v.Name == nombre_vista:
        vista_audit = v
        break

# 2. Si no existe, buscar el ID de tipo 'Drafting' y crearla
if not vista_audit:
    # Usamos un filtro manual que funciona en CUALQUIER versión de Revit
    tipos_vista = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
    id_tipo_drafting = None
    
    for tv in tipos_vista:
        if tv.ViewFamily == DB.ViewFamily.Drafting:
            id_tipo_drafting = tv.Id
            break

    if id_tipo_drafting:
        with revit.Transaction("Crear Vista Audit Menvic"):
            # Creamos la vista
            vista_audit = DB.ViewDrafting.Create(doc, id_tipo_drafting)
            vista_audit.Name = nombre_vista
            vista_audit.Scale = 1
            
            # Disciplina a Coordinación (4095)
            param_disc = vista_audit.get_Parameter(DB.BuiltInParameter.VIEW_DISCIPLINE)
            if param_disc:
                param_disc.Set(4095)
    else:
        print("Error: No se encontró el tipo de vista 'Drafting' en el proyecto.")

# 3. Abrir la vista sí o sí (Sin mensajes en consola)
if vista_audit:
    uidoc.ActiveView = vista_audit
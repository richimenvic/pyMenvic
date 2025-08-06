# -*- coding: utf-8 -*-
from Autodesk.Revit.UI import TaskDialog

autor = "Ricardo Javier Mendieta Cárdenas"
uso = "para uso de MENVIC"
anio = "2025"

mensaje = "Este conjunto de herramientas fue desarrollado por:\n\n{0}\n{1}\n{2}".format(autor, uso, anio)

TaskDialog.Show("Acerca del Pack", mensaje)

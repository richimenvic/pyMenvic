# -*- coding: utf-8 -*-
from pyrevit import forms
from Autodesk.Revit.DB import *
import os
import shutil
import re
import time

# Configuración
extensiones_relacionadas = ['.txt', '.xml', '.jpg', '.png', '.pdf', '.cvr', '.otd']
app = __revit__.Application

# Elegir carpeta
carpeta_base = forms.pick_folder(title='Selecciona la carpeta con archivos .rfa')
if not carpeta_base:
    script.exit()

# Funciones auxiliares
def limpiar_nombre(nombre):
    return re.sub(r'[<>:"/\\|?*]', '_', nombre)

def tiene_sufijo(nombre):
    return re.search(r'_V\d{2}$', nombre) is not None

def detectar_version(filepath):
    try:
        with open(filepath, 'rb') as f:
            contenido = f.read(256).decode(errors='ignore')
            for linea in contenido.splitlines():
                if 'Autodesk Revit' in linea:
                    for token in linea.split():
                        if token.isdigit() and len(token) == 4:
                            return "_V" + token[-2:]
    except:
        pass
    return "_V??"

def detectar_categoria(ruta_rfa):
    doc = None
    try:
        doc = app.OpenDocumentFile(ruta_rfa)
        if doc and doc.Family and doc.Family.FamilyCategory:
            return doc.Family.FamilyCategory.Name
    except:
        pass
    finally:
        if doc:
            doc.Close(False)
    return "Uncategorized"

# Procesar archivos
for archivo in os.listdir(carpeta_base):
    if archivo.lower().endswith('.rfa'):
        ruta_rfa = os.path.join(carpeta_base, archivo)
        nombre_base_original, extension = os.path.splitext(archivo)
        nombre_base_limpio = limpiar_nombre(nombre_base_original)

        sufijo = detectar_version(ruta_rfa)
        if not tiene_sufijo(nombre_base_limpio):
            nuevo_nombre_base = nombre_base_limpio + sufijo
        else:
            nuevo_nombre_base = nombre_base_limpio

        nuevo_nombre_rfa = limpiar_nombre(nuevo_nombre_base + extension)
        categoria = limpiar_nombre(detectar_categoria(ruta_rfa))

        carpeta_destino = os.path.join(carpeta_base, categoria)
        if not os.path.exists(carpeta_destino):
            os.mkdir(carpeta_destino)

        ruta_destino_rfa = os.path.join(carpeta_destino, nuevo_nombre_rfa)
        try:
            if not os.path.exists(ruta_destino_rfa):
                shutil.move(ruta_rfa, ruta_destino_rfa)
        except:
            continue

        for ext in extensiones_relacionadas:
            archivo_rel = nombre_base_original + ext
            ruta_rel = os.path.join(carpeta_base, archivo_rel)
            if os.path.exists(ruta_rel):
                if not tiene_sufijo(nombre_base_limpio):
                    nuevo_nombre_rel = limpiar_nombre(nuevo_nombre_base + ext)
                else:
                    nuevo_nombre_rel = limpiar_nombre(nombre_base_limpio + ext)

                destino_rel = os.path.join(carpeta_destino, nuevo_nombre_rel)
                try:
                    if not os.path.exists(destino_rel):
                        shutil.move(ruta_rel, destino_rel)
                except:
                    pass

        # 🕐 Pausa para liberar memoria
        time.sleep(0.2)

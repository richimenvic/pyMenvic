# -*- coding: utf-8 -*-

import os
import sys

try:
    from lib.core.tab_sorter import sort_tabs_by_document
except ImportError:
    p = os.path.dirname(os.path.abspath(__file__))
    while True:
        if os.path.basename(p).lower() == "pymenvic.extension":
            lib_dir = os.path.join(p, "lib")
            if lib_dir not in sys.path:
                sys.path.append(lib_dir)
            break
        q = os.path.dirname(p)
        if q == p:
            break
        p = q
    from core.tab_sorter import sort_tabs_by_document

STATE_FILE = os.path.join(os.environ.get("LOCALAPPDATA", os.environ.get("TEMP", os.getcwd())), "Temp", "pyMenvic", "tab_sort_once.txt")
if not os.environ.get("LOCALAPPDATA", ""):
    STATE_FILE = os.path.join(os.environ.get("TEMP", os.getcwd()), "pyMenvic", "tab_sort_once.txt")

try:
    value = ""
    if os.path.exists(STATE_FILE):
        f = open(STATE_FILE, "r")
        value = f.read().strip()
        f.close()
    if value == "1":
        folder = os.path.dirname(STATE_FILE)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        f = open(STATE_FILE, "w")
        f.write("0")
        f.close()
        sort_tabs_by_document()
except:
    pass

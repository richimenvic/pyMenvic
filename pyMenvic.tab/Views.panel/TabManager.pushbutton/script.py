# -*- coding: utf-8 -*-

__title__ = "Sort Tabs"
__author__ = "Ricardo J. Mendieta"

import os

STATE_FILE = os.path.join(os.environ.get("LOCALAPPDATA", os.environ.get("TEMP", os.getcwd())), "Temp", "pyMenvic", "tab_sort_once.txt")
if not os.environ.get("LOCALAPPDATA", ""):
    STATE_FILE = os.path.join(os.environ.get("TEMP", os.getcwd()), "pyMenvic", "tab_sort_once.txt")


def main():
    try:
        folder = os.path.dirname(STATE_FILE)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        f = open(STATE_FILE, "w")
        f.write("1")
        f.close()
    except:
        pass


if __name__ == "__main__":
    main()

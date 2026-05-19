# -*- coding: utf-8 -*-

__title__ = "Sort Tabs"
__author__ = "Ricardo J. Mendieta"

import os
import sys

try:
    from lib.core.tab_sorter import sort_tabs_by_document
except ImportError:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    while True:
        if os.path.basename(current_dir).lower() == "pymenvic.extension":
            lib_dir = os.path.join(current_dir, "lib")
            if lib_dir not in sys.path:
                sys.path.append(lib_dir)
            break
        parent_dir = os.path.dirname(current_dir)
        if parent_dir == current_dir:
            break
        current_dir = parent_dir
    from core.tab_sorter import sort_tabs_by_document


def main():
    try:
        sort_tabs_by_document()
    except:
        pass


if __name__ == "__main__":
    main()

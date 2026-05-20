# -*- coding: utf-8 -*-
"""Local branding fallback for Workset Seeder.

This avoids relying on pyRevit sys.path to find the extension-level lib package.
"""

import os


def _find_extension_root(start_dir):
    current = os.path.abspath(start_dir)
    while True:
        if os.path.basename(current).lower() == "pymenvic.extension":
            return current
        parent = os.path.dirname(current)
        if parent == current:
            return os.path.abspath(start_dir)
        current = parent


def get_logo_path(dark=False):
    _ = dark
    extension_root = _find_extension_root(os.path.dirname(__file__))
    return os.path.join(extension_root, "_resources", "logos", "menvic_logo.png")

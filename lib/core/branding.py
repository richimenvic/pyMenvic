# -*- coding: utf-8 -*-
"""Branding helpers for shared UI assets."""

import os


def get_logo_path(dark=False):
    """Return absolute path to shared pyMenvic logo.

    Args:
        dark (bool): Reserved for future dark-theme variants.
    """
    _ = dark
    core_dir = os.path.dirname(os.path.abspath(__file__))
    repo_root = os.path.dirname(os.path.dirname(core_dir))
    return os.path.join(repo_root, "_resources", "logos", "menvic_logo.png")

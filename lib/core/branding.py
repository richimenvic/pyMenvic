# -*- coding: utf-8 -*-
"""Branding helpers for shared UI assets."""

import os


def _get_logo_file_path(filename):
    """Return absolute path to a shared pyMenvic logo asset by filename."""
    core_dir = os.path.dirname(os.path.abspath(__file__))
    repo_root = os.path.dirname(os.path.dirname(core_dir))
    return os.path.join(repo_root, "_resources", "logos", filename)


def get_logo_path(dark=False):
    """Return absolute path to shared pyMenvic logo.

    Args:
        dark (bool): Reserved for future dark-theme variants.
    """
    _ = dark
    return _get_logo_file_path("menvic_logo.png")


def get_about_logo_path(dark=False):
    """Return absolute path to About dialog pyMenvic logo.

    Args:
        dark (bool): Reserved for future dark-theme variants.
    """
    _ = dark
    return _get_logo_file_path("menvic_logo_about.png")

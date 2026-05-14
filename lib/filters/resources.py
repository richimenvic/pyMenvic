# -*- coding: utf-8 -*-

import os

try:
    from lib.core.branding import get_logo_path as _get_logo_path_from_branding
except ImportError:
    _get_logo_path_from_branding = None


def _resolve_logo_path_from_extension(relative_logo_path):
    """Resolve logo path by walking up from this file to pyMenvic.extension."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    while True:
        if os.path.basename(current_dir) == "pyMenvic.extension":
            return os.path.join(current_dir, relative_logo_path)

        parent_dir = os.path.dirname(current_dir)
        if parent_dir == current_dir:
            break
        current_dir = parent_dir

    return os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        relative_logo_path,
    )


def get_filters_logo_path():
    """Shared filters header logo path resolved from pyMenvic.extension."""
    relative_logo_path = "_resources/logos/menvic_logo.png"

    if _get_logo_path_from_branding is not None:
        try:
            return _get_logo_path_from_branding(relative_logo_path)
        except Exception:
            pass

    return _resolve_logo_path_from_extension(relative_logo_path)

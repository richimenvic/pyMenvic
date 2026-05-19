# -*- coding: utf-8 -*-

import os
import sys
import time

from pyrevit import HOST_APP

try:
    from lib.core.tab_sorter import sort_tabs_by_document
    from lib.core.tab_sort_request import consume_sort_request
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
    from core.tab_sort_request import consume_sort_request


SUBSCRIBED_ENVVAR = "PYMENVIC_TABS_STARTUP_SUBSCRIBED"
LAST_RUN_ENVVAR = "PYMENVIC_TABS_STARTUP_LAST_RUN"
MIN_INTERVAL_SECONDS = 0.30


def _safe_float(value, default_value):
    try:
        return float(value)
    except:
        return default_value


def _on_idling(sender, args):
    try:
        if not consume_sort_request():
            return

        now = time.time()
        last_run = _safe_float(os.environ.get(LAST_RUN_ENVVAR, "0"), 0.0)
        if now - last_run < MIN_INTERVAL_SECONDS:
            return

        os.environ[LAST_RUN_ENVVAR] = str(now)
        sort_tabs_by_document()
    except:
        pass


def _subscribe_events():
    try:
        if os.environ.get(SUBSCRIBED_ENVVAR, "") == "1":
            return
        uiapp = HOST_APP.uiapp
        if uiapp is None:
            return
        uiapp.Idling += _on_idling
        os.environ[SUBSCRIBED_ENVVAR] = "1"
    except:
        pass


_subscribe_events()

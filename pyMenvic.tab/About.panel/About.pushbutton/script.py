# -*- coding: utf-8 -*-
__title__ = "About pyMenvic"

import os
import sys
import time
import webbrowser
import subprocess

try:
    import urllib2
except ImportError:
    import urllib.request as urllib2

try:
    from lib.core.branding import get_about_logo_path
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
    from core.branding import get_about_logo_path

from pyrevit import forms, script, versionmgr
from System import Uri
from System.Windows import Clipboard
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption


CONTACT_EMAIL = "contact@menvic.com"
WEBSITE_URL = "https://github.com/richimenvic/pyMenvic"
REMOTE_VERSION_URL = "https://raw.githubusercontent.com/richimenvic/pyMenvic/main/version.txt"
GIT_PULL_CMD = ["git", "pull", "--ff-only", "origin", "main"]


def _find_extension_root(start_dir):
    current_dir = os.path.abspath(start_dir)
    while True:
        if os.path.basename(current_dir).lower() == "pymenvic.extension":
            return current_dir
        if os.path.exists(os.path.join(current_dir, "pyMenvic.tab")):
            return current_dir
        parent_dir = os.path.dirname(current_dir)
        if parent_dir == current_dir:
            return None
        current_dir = parent_dir


def read_local_version():
    bundle_version_file = script.get_bundle_file('version.txt')
    extension_root = _find_extension_root(os.path.dirname(os.path.abspath(__file__)))

    candidate_files = []
    if extension_root:
        candidate_files.append(os.path.join(extension_root, "version.txt"))
    if bundle_version_file:
        candidate_files.append(bundle_version_file)

    for vfile in candidate_files:
        if vfile and os.path.exists(vfile):
            try:
                with open(vfile, 'r') as f:
                    return f.read().strip()
            except:
                continue

    return "0.0.0"


def parse_version(version_text):
    cleaned = (version_text or "").strip()
    if not cleaned:
        return [0]

    parts = cleaned.split(".")
    numbers = []
    for part in parts:
        token = (part or "").strip()
        if not token:
            numbers.append(0)
            continue

        digits = []
        for ch in token:
            if ch.isdigit():
                digits.append(ch)
            else:
                break

        if digits:
            numbers.append(int("".join(digits)))
        else:
            numbers.append(0)

    while len(numbers) > 1 and numbers[-1] == 0:
        numbers.pop()

    return numbers


def compare_versions(local_version, remote_version):
    local_numbers = parse_version(local_version)
    remote_numbers = parse_version(remote_version)

    max_len = max(len(local_numbers), len(remote_numbers))
    for i in range(max_len):
        local_part = local_numbers[i] if i < len(local_numbers) else 0
        remote_part = remote_numbers[i] if i < len(remote_numbers) else 0
        if remote_part > local_part:
            return 1
        if remote_part < local_part:
            return -1

    return 0


def read_remote_version():
    request = urllib2.Request(REMOTE_VERSION_URL)
    response = urllib2.urlopen(request, timeout=8)
    try:
        raw_content = response.read()
    finally:
        try:
            response.close()
        except:
            pass

    try:
        return raw_content.decode("utf-8").strip()
    except:
        return str(raw_content).strip()


def _run_process(command, cwd, timeout_seconds):
    proc = subprocess.Popen(
        command,
        cwd=cwd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )

    start_time = time.time()
    while proc.poll() is None:
        if timeout_seconds and (time.time() - start_time) > timeout_seconds:
            try:
                proc.kill()
            except:
                pass
            return -1, "", "Process timed out."
        time.sleep(0.1)

    stdout_data, stderr_data = proc.communicate()

    try:
        stdout_text = stdout_data.decode("utf-8", "ignore")
    except:
        stdout_text = str(stdout_data)

    try:
        stderr_text = stderr_data.decode("utf-8", "ignore")
    except:
        stderr_text = str(stderr_data)

    return proc.returncode, stdout_text.strip(), stderr_text.strip()


def update_extension_from_git():
    extension_root = _find_extension_root(os.path.dirname(os.path.abspath(__file__)))
    if not extension_root:
        return False, "Could not locate the pyMenvic extension folder."

    git_folder = os.path.join(extension_root, ".git")
    if not os.path.exists(git_folder):
        return False, "This pyMenvic installation is not a Git repository."

    try:
        code, _, err = _run_process(["git", "--version"], extension_root, 8)
    except OSError:
        return False, "Git is not installed or not available in PATH."

    if code != 0:
        return False, err or "Git is not installed or not available in PATH."

    code, _, err = _run_process(["git", "rev-parse", "--is-inside-work-tree"], extension_root, 8)
    if code != 0:
        return False, err or "This pyMenvic installation is not a Git repository."

    code, out, err = _run_process(GIT_PULL_CMD, extension_root, 45)
    if code != 0:
        reason = err or out or "Unknown error."
        return False, reason

    return True, ""


def load_bitmap(path):
    bmp = BitmapImage()
    bmp.BeginInit()
    bmp.UriSource = Uri(path)
    bmp.CacheOption = BitmapCacheOption.OnLoad
    bmp.EndInit()
    return bmp


class AboutWindow(forms.WPFWindow):

    def __init__(self):
        xamlfile = script.get_bundle_file('AboutWindow.xaml')
        forms.WPFWindow.__init__(self, xamlfile)

        # VERSION
        version = read_local_version()
        self.titleText.Text = "pyMenvic - Version {}".format(version)

        # ENVIRONMENT INFO
        try:
            pyrevit_ver = versionmgr.get_pyrevit_version().get_formatted()
        except:
            pyrevit_ver = "Unknown"

        try:
            revit_ver = __revit__.Application.VersionName
        except:
            revit_ver = "Revit"

        self.envText.Text = "Running on {} | pyRevit {}".format(revit_ver, pyrevit_ver)

        # EMAIL
        self.emailText.Text = CONTACT_EMAIL

        # LOGO
        icon_path = get_about_logo_path()
        if icon_path and os.path.exists(icon_path):
            bmp = load_bitmap(icon_path)
            self.logo.Source = bmp
            self.Icon = bmp

        # BUTTONS
        self.copyEmailBtn.Click += self.copy_email
        self.webBtn.Click += self.open_website
        self.checkUpdatesBtn.Click += self.check_for_updates

    def copy_email(self, sender, args):
        Clipboard.SetText(CONTACT_EMAIL)
        forms.alert("Email copied to clipboard.", title="pyMenvic")

    def open_website(self, sender, args):
        webbrowser.open(WEBSITE_URL)

    def check_for_updates(self, sender, args):
        try:
            local_version = read_local_version()
            remote_version = read_remote_version()
            comparison = compare_versions(local_version, remote_version)

            if comparison == 1:
                update_message = (
                    "New update for pyMenvic available.\n\n"
                    "Installed version: {}\n"
                    "Latest version: {}\n\n"
                    "Do you want to update now?"
                ).format(local_version, remote_version)
                update_now = forms.alert(
                    update_message,
                    title="pyMenvic",
                    yes=True,
                    no=True,
                    ok=False
                )
                if update_now:
                    success, reason = update_extension_from_git()
                    if success:
                        forms.alert(
                            "pyMenvic was updated successfully.\n"
                            "Please reload pyRevit or restart Revit to apply the changes.",
                            title="pyMenvic"
                        )
                    else:
                        forms.alert(
                            "Could not update pyMenvic automatically.\n"
                            "Reason: {}\n\n"
                            "Please update manually from GitHub or run:\n"
                            "git pull --ff-only origin main".format(reason),
                            title="pyMenvic"
                        )
            else:
                forms.alert("pyMenvic is up to date.", title="pyMenvic")
        except:
            forms.alert("Could not check for updates. Please verify your internet connection or GitHub access.", title="pyMenvic")


AboutWindow().ShowDialog()

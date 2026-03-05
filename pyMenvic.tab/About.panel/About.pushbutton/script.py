# -*- coding: utf-8 -*-

import os
import webbrowser

from pyrevit import forms, script, versionmgr
from System import Uri
from System.Windows import Clipboard
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption


CONTACT_EMAIL = "contact@menvic.com"
WEBSITE_URL = "https://menvic.com"


def read_version():
    vfile = script.get_bundle_file('version.txt')
    if vfile and os.path.exists(vfile):
        try:
            with open(vfile, 'r') as f:
                return f.read().strip()
        except:
            pass
    return "1.0"


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
        version = read_version()
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
        icon_path = script.get_bundle_file('logo.png')
        if icon_path and os.path.exists(icon_path):
            bmp = load_bitmap(icon_path)
            self.logo.Source = bmp
            self.Icon = bmp

        # BUTTONS
        self.closeBtn.Click += self.close_window
        self.copyEmailBtn.Click += self.copy_email
        self.webBtn.Click += self.open_website

    def close_window(self, sender, args):
        self.Close()

    def copy_email(self, sender, args):
        Clipboard.SetText(CONTACT_EMAIL)
        forms.alert("Email copied to clipboard.", title="pyMenvic")

    def open_website(self, sender, args):
        webbrowser.open(WEBSITE_URL)


AboutWindow().ShowDialog()
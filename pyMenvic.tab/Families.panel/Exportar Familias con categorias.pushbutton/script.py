# -*- coding: utf-8 -*-
"""Saves selected families into folders named by their category (in English)."""

import os
import os.path as op
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script

logger = script.get_logger()
output = script.get_output()

family_dict = {}
for family in revit.query.get_families(revit.doc, only_editable=True):
    if family.FamilyCategory:
        key = "%s: %s" % (family.FamilyCategory.Name, family.Name)
        family_dict[key] = family

if family_dict:
    selected_families = forms.SelectFromList.show(
        sorted(family_dict.keys()),
        title="Select Families to Save",
        multiselect=True,
    )

    if selected_families:
        dest_folder = forms.pick_folder()
        if dest_folder:
            selected_option = forms.CommandSwitchWindow.show(
                ["Skip Existing Families", "Overwrite Existing Families"],
                message="Choose what to do with existing families:",
            )
            overwrite_exst = selected_option == "Overwrite Existing Families"
            save_opts = DB.SaveAsOptions()
            save_opts.OverwriteExistingFile = overwrite_exst
            total_work = len(selected_families)

            for idx, family in enumerate([family_dict[x] for x in selected_families]):
                # Categoría en inglés (según Revit)
                cat_name = family.FamilyCategory.Name if family.FamilyCategory else "Uncategorized"
                subfolder = op.join(dest_folder, cat_name)
                if not os.path.exists(subfolder):
                    os.makedirs(subfolder)

                family_filepath = op.join(subfolder, family.Name + ".rfa")

                if not overwrite_exst and os.path.exists(family_filepath):
                    logger.info("Skipping existing family %s ...", family_filepath)
                else:
                    logger.info(
                        "%s %s ...",
                        "Updating" if os.path.exists(family_filepath) else "Saving",
                        family_filepath,
                    )
                    try:
                        family_doc = revit.doc.EditFamily(family)
                        family_doc.SaveAs(family_filepath, save_opts)
                        family_doc.Close(False)
                    except Exception as ex:
                        logger.error("Error saving family %s | %s", family_filepath, ex)

                output.update_progress(idx + 1, total_work)
else:
    forms.alert("No editable families found in the project.")

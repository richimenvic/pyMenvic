# Branding Centralization Report

- **Branch name:** `cleanup/centralize-branding-logo`
- **New shared logo path:** `resources/branding/pyMenvic_logo.png`

## Tools updated to shared internal UI logo

1. `pyMenvic.tab/WorksetsStandards.panel/WorkSet1.stack/WorksetStandardizer.pushbutton`
   - old: `.../WorksetStandardizer.pushbutton/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`
2. `pyMenvic.tab/Filters.panel/Filters1.stack/ReplaceFiltersInViews.pushbutton`
   - old: `.../ReplaceFiltersInViews.pushbutton/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`
3. `pyMenvic.tab/Filters.panel/Filters1.stack/ManageFilters.pushbutton`
   - old: `.../ManageFilters.pushbutton/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`
4. `pyMenvic.tab/Filters.panel/Filters1.stack/RenameFilters.pushbutton`
   - old: `.../RenameFilters.pushbutton/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`
5. `pyMenvic.tab/About.panel/About.pushbutton`
   - old: `.../About.pushbutton/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`
6. `pyMenvic.tab/pyMenvicRoots.panel/ScheduleBrowser.pushbutton`
   - old: `.../ScheduleBrowser.pushbutton/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`
7. `pyMenvic.tab/WorksetsStandards.panel/WorkSet3.stack/WorksetSeeder.pushbutton`
   - old: `.../WorksetSeeder.pushbutton/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`
8. `pyMenvic.tab/WorksetsStandards.panel/WorkSet1.stack/WorksetMappingManager.pushbutton`
   - old: `.../WorksetMappingManager.pushbutton/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`
9. `pyMenvic.tab/ViewTypeAudit.panel/ViewTypeManager.pushbutton`
   - old: `.../ViewTypeManager.pushbutton/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`
10. `pyMenvic.tab/WorksetsStandards.panel/WorkSet1.stack/LinkWorksetMananger.pushbutton`
   - old: `.../LinkWorksetMananger.pushbutton/resources/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`
11. `pyMenvic.tab/WorksetsStandards.panel/WorkSet2.stack/SyncLinkWorksets.pushbutton`
   - old: `.../SyncLinkWorksets.pushbutton/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`
12. `pyMenvic.tab/WorksetsStandards.panel/WorkSet2.stack/LevelsGrids2Workset.pushbutton`
   - old: `.../LevelsGrids2Workset.pushbutton/logo.png`
   - new: `resources/branding/pyMenvic_logo.png`

## Duplicated logo files removed

- `pyMenvic.tab/WorksetsStandards.panel/WorkSet1.stack/WorksetStandardizer.pushbutton/logo.png`
- `pyMenvic.tab/Filters.panel/Filters1.stack/ReplaceFiltersInViews.pushbutton/logo.png`
- `pyMenvic.tab/Filters.panel/Filters1.stack/ManageFilters.pushbutton/logo.png`
- `pyMenvic.tab/Filters.panel/Filters1.stack/RenameFilters.pushbutton/logo.png`
- `pyMenvic.tab/About.panel/About.pushbutton/logo.png`
- `pyMenvic.tab/pyMenvicRoots.panel/ScheduleBrowser.pushbutton/logo.png`
- `pyMenvic.tab/WorksetsStandards.panel/WorkSet3.stack/WorksetSeeder.pushbutton/logo.png`
- `pyMenvic.tab/WorksetsStandards.panel/WorkSet1.stack/WorksetMappingManager.pushbutton/logo.png`
- `pyMenvic.tab/ViewTypeAudit.panel/ViewTypeManager.pushbutton/logo.png`
- `pyMenvic.tab/WorksetsStandards.panel/WorkSet1.stack/LinkWorksetMananger.pushbutton/resources/logo.png`
- `pyMenvic.tab/WorksetsStandards.panel/WorkSet2.stack/SyncLinkWorksets.pushbutton/logo.png`
- `pyMenvic.tab/WorksetsStandards.panel/WorkSet2.stack/LevelsGrids2Workset.pushbutton/logo.png`

## Duplicated logo files intentionally kept

- None.

## `icon.png` files intentionally not changed

`icon.png` files were not modified, because pyRevit ribbon buttons rely on those assets and they are separate from internal UI branding logos.

## Manual tests needed in Revit/pyRevit

1. Open each updated tool and verify its window renders the centralized logo.
2. Confirm each updated tool launches and runs without import/path errors.
3. Confirm all ribbon icons (`icon.png`) remain unchanged.
4. Confirm no tool still depends on a local `logo.png` file.

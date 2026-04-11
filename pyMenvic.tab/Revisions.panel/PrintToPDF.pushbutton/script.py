# -*- coding: utf-8 -*-
__title__ = "PRINT TO PDF"
"""
==========================================================
pyMENVIC | PRINT SHEETS IN ORDER
Revit + pyRevit

Descripción
-----------
Herramienta para gestionar la impresión y exportación
ordenada de hojas del proyecto desde Revit.

Permite controlar la secuencia de salida, generar archivos
PDF individuales o combinados, exportar formatos CAD
complementarios y aplicar reglas automáticas de nombrado.

Capacidades
-----------
- impresión de hojas en orden personalizado
- exportación PDF individual o combinada
- exportación adicional en DWG o DXF
- ordenamiento por disciplina, numeración o ajuste manual
- generación automática de nombres de archivo
- soporte para modelos activos y vinculados
- guardado de selección como Sheet Set
- persistencia del orden mediante PRINT_ORDER

Funciones principales
---------------------
PRINT / EXPORT
    Ejecuta la impresión o exportación de hojas.

COMBINE INTO ONE FILE
    Genera un único PDF con la selección activa.

ADD SHEETS
    Añade hojas manualmente a la lista de trabajo.

REMOVE SELECTED
    Elimina hojas seleccionadas de la lista activa.

SAVE CURRENT SET
    Guarda la selección actual como Sheet Set.

SAVE PRINT ORDER
    Registra el orden actual en PRINT_ORDER.

RENAME FILES
    Aplica nombres automáticos mediante plantillas.

Criterios de orden
------------------
- disciplina
- numeración natural
- secuencia manual
- inversión de orden
- restauración al estándar

Disciplinas:
G, S, A, M, E, T, P

Reglas importantes
------------------
- En versiones anteriores a ciertos flujos nativos, el script
  puede usar caracteres no imprimibles para forzar el orden.
- Shift-Click limpia esos caracteres del número de hoja.
- Si PRINT_ORDER no existe, puede crearse automáticamente.
- El nombre de salida admite prefijos y plantillas dinámicas.
- En versiones compatibles, se utiliza exportación PDF nativa.

Variables de nombre
-------------------
{index}, {number}, {name}, {name_dash}, {name_underline},
{current_date}, {issue_date}, {rev_number}, {rev_desc},
{rev_date}, {proj_name}, {proj_number}, {proj_building_name},
{proj_issue_date}, {proj_org_name}, {proj_status},
{username}, {revit_version},
{sheet_param:PARAM_NAME}, {tblock_param:PARAM_NAME},
{proj_param:PARAM_NAME}, {glob_param:PARAM_NAME}

Autor
-----
Ricardo J. Mendieta
pyMENVIC – Ayudas para MENVIC ARQ
==========================================================
"""

# pylint: disable=unused-argument,too-many-lines
# pylint: disable=missing-function-docstring,missing-class-docstring
"""Print sheets in order from a sheet index.

Note:
When using the `Combine into one file` option
in Revit 2022 and earlier,
the tool adds non-printable character u'\u200e'
(Left-To-Right Mark) at the start of the sheet names
to push Revit's internal printing engine to sort
the sheets correctly per the drawing index order.

Make sure your drawings indices consider this
when filtering for sheet numbers.

Shift-Click:
Shift-Clicking the tool will remove all
non-printable characters from the sheet numbers,
in case an error in the tool causes these characters
to remain.
"""
# pylint: disable=import-error,invalid-name,broad-except,superfluous-parens


# =============================================================================
# IMPORTS
# =============================================================================

import re
import os.path as op
import codecs
import os
import datetime
import locale
import tempfile
from collections import namedtuple

from pyrevit import HOST_APP
from pyrevit import framework
from pyrevit.framework import Windows, Drawing, ObjectModel, Forms, List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import revit, DB
from pyrevit import script
from pyrevit.compat import get_elementid_value_func
from pyrevit import forms


# =============================================================================
# GLOBALS & CONSTANTS
# =============================================================================

get_elementid_value = get_elementid_value_func()

logger = script.get_logger()
config = script.get_config()

# Non Printable Char
NPC = u'\u200e'
INDEX_FORMAT = '{{:0{digits}}}'

EXPORT_ENCODING = 'utf_16_le'
if HOST_APP.is_newer_than(2020):
    EXPORT_ENCODING = 'utf_8'

IS_REVIT_2022_OR_NEWER = HOST_APP.is_newer_than(2021)


# =============================================================================
# SORTING UTILITIES
# =============================================================================

def natural_sort_key(value):
    value = value or ''
    return [int(part) if part.isdigit() else part.lower()
            for part in re.split(r'([0-9]+)', value)]

DISCIPLINE_ORDER = ['G', 'S', 'A', 'M', 'E', 'T', 'P']


def extract_sheet_code(sheet_number):
    value = (sheet_number or '').strip().upper()

    match = re.match(r'^[A-Z]+-([A-Z]+)', value)
    if match:
        return match.group(1)

    match = re.match(r'^([A-Z]+)', value)
    if match:
        return match.group(1)

    return ''


def get_discipline_rank(sheet_number):
    code = extract_sheet_code(sheet_number)
    try:
        return DISCIPLINE_ORDER.index(code)
    except ValueError:
        return 999


def sheet_number_sort_key(value):
    value = (value or '').strip()
    parts = re.split(r'([0-9]+)', value)
    natural_parts = [int(part) if part.isdigit() else part.lower()
                     for part in parts]
    return natural_parts


def sheet_index_sort_key(name):
    name = (name or '').strip()
    normalized = re.sub(r'\s+', ' ', name).lower()
    # Natural alphabetic sort so "Revision 4" comes before "Revision 10"
    return natural_sort_key(normalized)


# =============================================================================
# FILE NAME UTILITIES
# =============================================================================

def ensure_pdf_extension(filename):
    filename = (filename or '').strip()
    if not filename:
        return ''
    if not filename.lower().endswith('.pdf'):
        filename += '.pdf'
    return filename


# =============================================================================
# DATA STRUCTURES (NamedTuples)
# =============================================================================

AvailableDoc = namedtuple('AvailableDoc', ['name', 'hash', 'linked'])

NamingFormatter = namedtuple('NamingFormatter', ['template', 'desc'])

SheetRevision = namedtuple('SheetRevision', ['number', 'desc', 'date', 'is_set'])
UNSET_REVISION = SheetRevision(number=None, desc=None, date=None, is_set=False)

TitleBlockPrintSettings = \
    namedtuple('TitleBlockPrintSettings', ['psettings', 'set_by_param'])


DEFAULT_PDF_EXPORT_OVERRIDES = {
    'enabled': False,
    'paper_placement': 'Center',
    'origin_offset_x_mm': '0',
    'origin_offset_y_mm': '0',
    'output_processing': 'Vector',
    'raster_quality': 'High',
    'view_links_in_blue': True,
    'hide_reference_plane': True,
    'hide_scope_boxes': True,
    'hide_crop_boundaries': True,
    'hide_unreferenced_view_tags': True,
    'replace_halftone_with_thin_lines': False,
    'mask_coincident_lines': True,
}


# =============================================================================
# PRINT UTILITIES CLASS
# =============================================================================

class PrintUtils:
    """Utility functions for printing and exporting sheets."""

    @staticmethod
    def get_doc():
        return revit.doc

    @staticmethod
    def get_dir():
        return os.path.join(os.path.expanduser("~"), "Desktop", "pyRevit Print Folder")

    @staticmethod
    def get_folder(task="_PDF"):
        dateStamp = datetime.datetime.today().strftime("%y%m%d")
        timeStamp = datetime.datetime.today().strftime("%H%M%S")
        return dateStamp + "_" + timeStamp + task

    @staticmethod
    def ensure_dir(dp):
        if not os.path.exists(dp):
            os.makedirs(dp)
        return dp

    @staticmethod
    def open_dir(dp):
        try:
            os.startfile(dp)
        except Exception:
            pass
        return dp

    @staticmethod
    def pdf_opts(doc=None, hcb=True, hsb=True, hrp=True, hvt=True, mcl=True):
        if doc:
            try:
                active_pdf_setup = DB.ExportPDFSettings.GetActivePredefinedSettings(doc)
                if active_pdf_setup:
                    opts = active_pdf_setup.GetOptions()
                    return PrintUtils.apply_pdf_overrides(opts)
            except Exception as pdf_setup_err:
                logger.warning(
                    'Failed to load active PDF export setup. Falling back to script defaults. | %s',
                    pdf_setup_err
                )

        opts = DB.PDFExportOptions()
        opts.HideCropBoundaries = hcb
        opts.HideScopeBoxes = hsb
        opts.HideReferencePlane = hrp
        opts.HideUnreferencedViewTags = hvt
        opts.MaskCoincidentLines = mcl
        opts.PaperFormat = DB.ExportPaperFormat.Default
        return PrintUtils.apply_pdf_overrides(opts)

    @staticmethod
    def get_pdf_overrides():
        overrides = dict(DEFAULT_PDF_EXPORT_OVERRIDES)
        overrides.update(config.get_option('pdf_export_overrides', {}))
        return overrides

    @staticmethod
    def set_pdf_overrides(overrides):
        merged = dict(DEFAULT_PDF_EXPORT_OVERRIDES)
        merged.update(overrides or {})
        config.pdf_export_overrides = merged
        script.save_config()

    @staticmethod
    def _parse_mm_to_feet(value):
        try:
            return float(str(value).replace(',', '.')) / 304.8
        except Exception:
            return 0.0

    @staticmethod
    def apply_pdf_overrides(opts):
        overrides = PrintUtils.get_pdf_overrides()
        if not overrides.get('enabled'):
            return opts

        try:
            placement_name = overrides.get('paper_placement', 'Center')
            if hasattr(DB, 'PaperPlacementType'):
                center_value = getattr(DB.PaperPlacementType, 'Center', None)
                offset_value = \
                    getattr(DB.PaperPlacementType, 'LowerLeft', None) \
                    or getattr(DB.PaperPlacementType, 'OffsetFromCorner', None)
                if placement_name == 'LowerLeft':
                    if offset_value is not None:
                        opts.PaperPlacement = offset_value
                elif center_value is not None:
                    opts.PaperPlacement = center_value

            if placement_name == 'LowerLeft':
                opts.OriginOffsetX = PrintUtils._parse_mm_to_feet(
                    overrides.get('origin_offset_x_mm', '0')
                )
                opts.OriginOffsetY = PrintUtils._parse_mm_to_feet(
                    overrides.get('origin_offset_y_mm', '0')
                )

            output_processing = overrides.get('output_processing', 'Automatic')
            if output_processing == 'Raster':
                opts.AlwaysUseRaster = True

            raster_quality = overrides.get('raster_quality', 'High')
            if hasattr(DB, 'RasterQualityType') and hasattr(opts, 'RasterQuality'):
                raster_quality_map = {
                    'Low': getattr(DB.RasterQualityType, 'Low', None),
                    'Medium': getattr(DB.RasterQualityType, 'Medium', None),
                    'High': getattr(DB.RasterQualityType, 'High', None),
                    'Presentation': getattr(DB.RasterQualityType, 'Presentation', None),
                }
                raster_quality_value = raster_quality_map.get(raster_quality)
                if raster_quality_value is not None:
                    opts.RasterQuality = raster_quality_value

            opts.ViewLinksInBlue = bool(overrides.get('view_links_in_blue'))
            opts.HideReferencePlane = bool(overrides.get('hide_reference_plane'))
            opts.HideScopeBoxes = bool(overrides.get('hide_scope_boxes'))
            opts.HideCropBoundaries = bool(overrides.get('hide_crop_boundaries'))
            opts.HideUnreferencedViewTags = bool(overrides.get('hide_unreferenced_view_tags'))
            opts.ReplaceHalftoneWithThinLines = bool(
                overrides.get('replace_halftone_with_thin_lines')
            )
            opts.MaskCoincidentLines = bool(overrides.get('mask_coincident_lines'))
        except Exception as apply_err:
            logger.warning(
                'Failed to apply PDF export overrides. Using active PDF setup values. | %s',
                apply_err
            )

        return opts

    @staticmethod
    def dwg_opts(sc=False, mv=True):
        opts = DB.DWGExportOptions()
        opts.SharedCoords = sc
        opts.MergedViews = mv
        return opts

    @staticmethod
    def export_sheet_pdf(dir_path, sheet, opt, doc, filename):
        pdf_doc_name = op.splitext(filename)[0]
        opt.FileName = pdf_doc_name
        export_sheet = List[DB.ElementId]()
        export_sheet.Add(sheet.Id)
        doc.Export(dir_path, export_sheet, opt)
        return True

    @staticmethod
    def export_sheet_cad(dir_path, sheet, opt, doc, filename, extension):
        base_name = op.splitext(filename)[0]
        cad_doc_name = base_name + extension
        export_sheet = List[DB.ElementId]()
        export_sheet.Add(sheet.Id)
        doc.Export(dir_path, cad_doc_name, export_sheet, opt)
        return True

    @staticmethod
    def export_sheet_dwg(dir_path, sheet, opt, doc, filename):
        return PrintUtils.export_sheet_cad(dir_path, sheet, opt, doc, filename, '.dwg')



# =============================================================================
# NAMING FORMAT CLASSES
# =============================================================================

class NamingFormat(forms.Reactive):
    """Print File Naming Format"""
    def __init__(self, name, template, builtin=False, auto_name=False):
        self.builtin = builtin
        self._manual_name = not auto_name
        self._syncing_name = False
        self._template = self.verify_template(template)
        if auto_name or not name or name == '<unnamed>':
            self._name = self.generate_name_from_template(self._template)
        else:
            self._name = name

    @staticmethod
    def verify_template(value):
        """Verify template is valid"""
        value = value or ''
        if not value.lower().endswith('.pdf'):
            value += '.pdf'
        return value

    @staticmethod
    def generate_name_from_template(template):
        template = template or ''
        preview = template
        replacements = {
            '{index}': '0001',
            '{number}': 'A1.00',
            '{name}': '1ST FLOOR PLAN',
            '{name_dash}': '1ST-FLOOR-PLAN',
            '{name_underline}': '1ST_FLOOR_PLAN',
            '{current_date}': 'YYYY-MM-DD',
            '{issue_date}': 'YYYY-MM-DD',
            '{rev_number}': '01',
            '{rev_desc}': 'REVISION',
            '{rev_date}': 'YYYY-MM-DD',
            '{proj_name}': 'PROJECT',
            '{proj_number}': 'P001',
            '{proj_building_name}': 'BUILDING',
            '{proj_issue_date}': 'YYYY-MM-DD',
            '{proj_org_name}': 'COMPANY',
            '{proj_status}': 'STATUS',
            '{username}': 'USER',
            '{revit_version}': '2025',
        }

        for token, sample in replacements.items():
            preview = preview.replace(token, sample)

        preview = re.sub(r'\{sheet_param:.*?\}', 'SHEET_PARAM', preview)
        preview = re.sub(r'\{tblock_param:.*?\}', 'TITLEBLOCK_PARAM', preview)
        preview = re.sub(r'\{proj_param:.*?\}', 'PROJECT_PARAM', preview)
        preview = re.sub(r'\{glob_param:.*?\}', 'GLOBAL_PARAM', preview)

        preview = re.sub(r'\s+', ' ', preview).strip()
        if not preview:
            preview = 'Custom Naming Format.pdf'
        return preview

    def sync_name_from_template(self):
        self._syncing_name = True
        try:
            self.name = self.generate_name_from_template(self._template)
        finally:
            self._syncing_name = False
            self._manual_name = False

    @forms.reactive
    def name(self):
        """Format name"""
        return self._name

    @name.setter
    def name(self, value):
        self._name = value
        if not self._syncing_name:
            self._manual_name = True

    @forms.reactive
    def template(self):
        """Format template string"""
        return self._template

    @template.setter
    def template(self, value):
        self._template = self.verify_template(value)
        if not self._manual_name:
            self.sync_name_from_template()


# =============================================================================
# SHEET LIST ITEM CLASS
# =============================================================================

class ViewSheetListItem(forms.Reactive):
    """Revit Sheet show in Print Window"""

    def __init__(self, view_sheet, view_tblock,
                 print_settings=None, rev_settings=None):
        self._sheet = view_sheet
        self._tblock = view_tblock
        if self._tblock:
            self._tblock_type = \
                view_sheet.Document.GetElement(view_tblock.GetTypeId())
        else:
            self._tblock_type = None
        self.name = self._sheet.Name
        self.number = self._sheet.SheetNumber if hasattr(self._sheet, 'SheetNumber') else ''
        self.issue_date = \
            self._sheet.Parameter[
                DB.BuiltInParameter.SHEET_ISSUE_DATE].AsString() if self._sheet.Parameter[
                DB.BuiltInParameter.SHEET_ISSUE_DATE] else ''
        self.printable = self._sheet.CanBePrinted
        self.revision_date_sortable = ""
        self._print_index = 0
        self._print_filename = ''

        self._tblock_psettings = print_settings
        self._print_settings = self._tblock_psettings.psettings
        self.all_print_settings = self._tblock_psettings.psettings
        if self.all_print_settings:
            self._print_settings = self.all_print_settings[0]
        self.read_only = self._tblock_psettings.set_by_param

        per_sheet_revisions = \
            rev_settings.RevisionNumbering == DB.RevisionNumbering.PerSheet \
            if rev_settings else False
        cur_rev = revit.query.get_current_sheet_revision(self._sheet) if hasattr(self._sheet, 'GetCurrentRevision') else ''
        self.revision = UNSET_REVISION
        if cur_rev:
            on_sheet = self._sheet if per_sheet_revisions else None
            self.revision = SheetRevision(
                number=revit.query.get_rev_number(cur_rev, sheet=on_sheet),
                desc=cur_rev.Description,
                date=cur_rev.RevisionDate,
                is_set=True
            )

    @property
    def revit_sheet(self):
        """Revit sheet instance"""
        return self._sheet

    @property
    def revit_tblock(self):
        """Revit titleblock instance"""
        return self._tblock

    @property
    def revit_tblock_type(self):
        """Revit titleblock type"""
        return self._tblock_type

    @forms.reactive
    def print_settings(self):
        """Sheet pring settings"""
        return self._print_settings

    @print_settings.setter
    def print_settings(self, value):
        self._print_settings = value

    @forms.reactive
    def print_index(self):
        """Sheet print index"""
        return self._print_index

    @print_index.setter
    def print_index(self, value):
        self._print_index = value

    @forms.reactive
    def print_filename(self):
        """Sheet print output filename"""
        return self._print_filename

    @print_filename.setter
    def print_filename(self, value):
        self._print_filename = \
            coreutils.cleanup_filename(value, windows_safe=True)


# =============================================================================
# PRINT SETTING LIST ITEM CLASSES
# =============================================================================

class PrintSettingListItem(forms.TemplateListItem):
    """Print Setting shown in Print Window"""

    def __init__(self, print_settings=None):
        super(PrintSettingListItem, self).__init__(print_settings)
        self.is_compatible = isinstance(self.item, DB.InSessionPrintSetting)

    @property
    def name(self):
        if isinstance(self.item, DB.InSessionPrintSetting):
            return "<In Session>"
        else:
            return self.item.Name

    @property
    def print_settings(self):
        return self.item

    @property
    def print_params(self):
        if self.print_settings:
            return self.print_settings.PrintParameters

    @property
    def paper_size(self):
        try:
            if self.print_params:
                return self.print_params.PaperSize
        except Exception:
            pass

    @property
    def allows_variable_paper(self):
        return False

    @property
    def is_user_defined(self):
        return not self.name.startswith('<')


class VariablePaperPrintSettingListItem(PrintSettingListItem):
    def __init__(self):
        PrintSettingListItem.__init__(self, None)
        # always compatible
        self.is_compatible = True

    @property
    def name(self):
        return "<Variable Paper Size>"

    @property
    def allows_variable_paper(self):
        return True


# =============================================================================
# EDIT NAMING FORMATS WINDOW
# =============================================================================

class EditNamingFormatsWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, start_with=None):
        forms.WPFWindow.__init__(self, xaml_file_name)

        self._drop_pos = 0
        self._starting_item = start_with
        self._saved = False

        self.reset_naming_formats()
        self.reset_formatters()

    # -------------------------------------------------------------------------
    # Default Data
    # -------------------------------------------------------------------------

    @staticmethod
    def get_default_formatters():
        return [
            NamingFormatter(
                template='{index}',
                desc='Print Index Number e.g. "0001"'
            ),
            NamingFormatter(
                template='{number}',
                desc='Sheet Number e.g. "A1.00"'
            ),
            NamingFormatter(
                template='{name}',
                desc='Sheet Name e.g. "1ST FLOOR PLAN"'
            ),
            NamingFormatter(
                template='{name_dash}',
                desc='Sheet Name (with - for space) e.g. "1ST-FLOOR-PLAN"'
            ),
            NamingFormatter(
                template='{name_underline}',
                desc='Sheet Name (with _ for space) e.g. "1ST_FLOOR_PLAN"'
            ),
            NamingFormatter(
                template='{current_date}',
                desc='Today''s Date e.g. "2019-10-12"'
            ),
            NamingFormatter(
                template='{issue_date}',
                desc='Sheet Issue Date e.g. "2019-10-12"'
            ),
            NamingFormatter(
                template='{rev_number}',
                desc='Revision Number e.g. "01"'
            ),
            NamingFormatter(
                template='{rev_desc}',
                desc='Revision Description e.g. "ASI01"'
            ),
            NamingFormatter(
                template='{rev_date}',
                desc='Revision Date e.g. "2019-10-12"'
            ),
            NamingFormatter(
                template='{proj_name}',
                desc='Project Name e.g. "MY_PROJECT"'
            ),
            NamingFormatter(
                template='{proj_number}',
                desc='Project Number e.g. "PR2019.12"'
            ),
            NamingFormatter(
                template='{proj_building_name}',
                desc='Project Building Name e.g. "BLDG01"'
            ),
            NamingFormatter(
                template='{proj_issue_date}',
                desc='Project Issue Date e.g. "2019-10-12"'
            ),
            NamingFormatter(
                template='{proj_org_name}',
                desc='Project Organization Name e.g. "MYCOMP"'
            ),
            NamingFormatter(
                template='{proj_status}',
                desc='Project Status e.g. "CD100"'
            ),
            NamingFormatter(
                template='{username}',
                desc='Active User e.g. "eirannejad"'
            ),
            NamingFormatter(
                template='{revit_version}',
                desc='Active Revit Version e.g. "2019"'
            ),
            NamingFormatter(
                template='{sheet_param:PARAM_NAME}',
                desc='Value of Given Sheet Parameter e.g. '
                     'Replace PARAM_NAME with target parameter name'
            ),
            NamingFormatter(
                template='{tblock_param:PARAM_NAME}',
                desc='Value of Given TitleBlock Parameter e.g. '
                     'Replace PARAM_NAME with target parameter name'
            ),
            NamingFormatter(
                template='{proj_param:PARAM_NAME}',
                desc='Value of Given Project Information Parameter e.g. '
                     'Replace PARAM_NAME with target parameter name'
            ),
            NamingFormatter(
                template='{glob_param:PARAM_NAME}',
                desc='Value of Given Global Parameter. '
                     'Replace PARAM_NAME with target parameter name'
            ),
        ]

    @staticmethod
    def get_default_naming_formats():
        return [
            NamingFormat(
                name='0001 A1.00 1ST FLOOR PLAN.pdf',
                template='{index} {number} {name}.pdf',
                builtin=True
            ),
            NamingFormat(
                name='0001_A1.00_1ST FLOOR PLAN.pdf',
                template='{index}_{number}_{name}.pdf',
                builtin=True
            ),
            NamingFormat(
                name='0001-A1.00-1ST FLOOR PLAN.pdf',
                template='{index}-{number}-{name}.pdf',
                builtin=True
            ),
        ]

    # -------------------------------------------------------------------------
    # Config Read / Write
    # -------------------------------------------------------------------------

    @staticmethod
    def get_naming_formats():
        naming_formats = EditNamingFormatsWindow.get_default_naming_formats()
        naming_formats_dict = config.get_option('namingformats', {})
        for name, template in naming_formats_dict.items():
            naming_formats.append(NamingFormat(name=name, template=template))
        return naming_formats

    @staticmethod
    def set_naming_formats(naming_formats):
        naming_formats_dict = {
            x.name: x.template for x in naming_formats if not x.builtin
        }
        config.namingformats = naming_formats_dict
        script.save_config()

    # -------------------------------------------------------------------------
    # Properties
    # -------------------------------------------------------------------------

    @property
    def naming_formats(self):
        return self.formats_lb.ItemsSource

    @property
    def selected_naming_format(self):
        return self.formats_lb.SelectedItem

    @selected_naming_format.setter
    def selected_naming_format(self, value):
        self.formats_lb.SelectedItem = value
        self.namingformat_edit.DataContext = value

    # -------------------------------------------------------------------------
    # UI Setup
    # -------------------------------------------------------------------------

    def reset_formatters(self):
        self.formatters_wp.ItemsSource = \
            EditNamingFormatsWindow.get_default_formatters()

    def reset_naming_formats(self):
        self.formats_lb.ItemsSource = \
                ObjectModel.ObservableCollection[object](
                    EditNamingFormatsWindow.get_naming_formats()
                )
        if isinstance(self._starting_item, NamingFormat):
            for item in self.formats_lb.ItemsSource:
                if item.name == self._starting_item.name:
                    self.selected_naming_format = item
                    break

    # -------------------------------------------------------------------------
    # Drag & Drop Handlers
    # -------------------------------------------------------------------------

    # https://www.wpftutorial.net/DragAndDrop.html
    def start_drag(self, sender, args):
        name_formatter = args.OriginalSource.DataContext
        Windows.DragDrop.DoDragDrop(
            self.formatters_wp,
            Windows.DataObject("name_formatter", name_formatter),
            Windows.DragDropEffects.Copy
            )

    # https://social.msdn.microsoft.com/Forums/vstudio/en-US/941f6bf2-a321-459e-85c9-501ec1e13204/how-do-you-get-a-drag-and-drop-event-for-a-wpf-textbox-hosted-in-a-windows-form
    def preview_drag(self, sender, args):
        mouse_pos = Forms.Cursor.Position
        mouse_po_pt = Windows.Point(mouse_pos.X, mouse_pos.Y)
        self._drop_pos = \
            self.template_tb.GetCharacterIndexFromPoint(
                point=self.template_tb.PointFromScreen(mouse_po_pt),
                snapToText=True
                )
        self.template_tb.SelectionStart = self._drop_pos
        self.template_tb.SelectionLength = 0
        self.template_tb.Focus()
        args.Effects = Windows.DragDropEffects.Copy
        args.Handled = True

    def stop_drag(self, sender, args):
        name_formatter = args.Data.GetData("name_formatter")
        if name_formatter:
            new_template = \
                str(self.template_tb.Text)[:self._drop_pos] \
                + name_formatter.template \
                + str(self.template_tb.Text)[self._drop_pos:]
            self.template_tb.Text = new_template
            self.template_tb.Focus()

    # -------------------------------------------------------------------------
    # Event Handlers
    # -------------------------------------------------------------------------

    def namingformat_changed(self, sender, args):
        naming_format = self.selected_naming_format
        self.namingformat_edit.DataContext = naming_format

    def duplicate_namingformat(self, sender, args):
        naming_format = self.selected_naming_format
        new_naming_format = NamingFormat(
            name='<unnamed>',
            template=naming_format.template,
            auto_name=True
            )
        self.naming_formats.Add(new_naming_format)
        self.selected_naming_format = new_naming_format

    def reset_name_from_template(self, sender, args):
        naming_format = self.selected_naming_format
        if naming_format:
            naming_format.sync_name_from_template()
            self.namingformat_edit.DataContext = None
            self.namingformat_edit.DataContext = naming_format

    def delete_namingformat(self, sender, args):
        naming_format = self.selected_naming_format
        if naming_format.builtin:
            return
        item_index = self.naming_formats.IndexOf(naming_format)
        self.naming_formats.Remove(naming_format)
        next_index = min([item_index, self.naming_formats.Count-1])
        self.selected_naming_format = self.naming_formats[next_index]

    def save_formats(self, sender, args):
        EditNamingFormatsWindow.set_naming_formats(self.naming_formats)
        self._saved = True
        self.Close()

    def cancelled(self, sender, args):
        if not self._saved:
            self.reset_naming_formats()

    def show_dialog(self):
        self.ShowDialog()


# =============================================================================
# PDF EXPORT OPTIONS WINDOW
# =============================================================================

class PdfExportOptionsWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self._saved = False
        self._load_settings()
        self.update_offset_state()

    def _load_settings(self):
        settings = PrintUtils.get_pdf_overrides()
        self.enableoverrides_cb.IsChecked = settings.get('enabled', False)
        use_offset = settings.get('paper_placement') == 'LowerLeft'
        self.paperplacement_center_rb.IsChecked = not use_offset
        self.paperplacement_offset_rb.IsChecked = use_offset
        self.offsetx_tb.Text = str(settings.get('origin_offset_x_mm', '0'))
        self.offsety_tb.Text = str(settings.get('origin_offset_y_mm', '0'))
        output_processing = settings.get('output_processing', 'Vector')
        self.usevector_rb.IsChecked = output_processing != 'Raster'
        self.useraster_rb.IsChecked = output_processing == 'Raster'
        raster_quality = settings.get('raster_quality', 'High')
        raster_quality_items = ['Low', 'Medium', 'High', 'Presentation']
        self.rasterquality_cb.SelectedIndex = \
            raster_quality_items.index(raster_quality) \
            if raster_quality in raster_quality_items else 2
        self.viewlinksinblue_cb.IsChecked = settings.get('view_links_in_blue', True)
        self.hiderefplanes_cb.IsChecked = settings.get('hide_reference_plane', True)
        self.hidescopeboxes_cb.IsChecked = settings.get('hide_scope_boxes', True)
        self.hidecropboundaries_cb.IsChecked = settings.get('hide_crop_boundaries', True)
        self.hideunrefviewtags_cb.IsChecked = settings.get('hide_unreferenced_view_tags', True)
        self.replacehalftone_cb.IsChecked = settings.get(
            'replace_halftone_with_thin_lines',
            False
        )
        self.maskcoincident_cb.IsChecked = settings.get('mask_coincident_lines', True)

    def update_offset_state(self, sender=None, args=None):
        enabled = bool(self.enableoverrides_cb.IsChecked)
        use_offsets = enabled and bool(self.paperplacement_offset_rb.IsChecked)
        self.options_panel.IsEnabled = enabled
        self.offsets_lbl.IsEnabled = use_offsets
        self.offsetx_tb.IsEnabled = use_offsets
        self.offsety_tb.IsEnabled = use_offsets

    def save(self, sender, args):
        settings = {
            'enabled': bool(self.enableoverrides_cb.IsChecked),
            'paper_placement': 'LowerLeft' if self.paperplacement_offset_rb.IsChecked else 'Center',
            'origin_offset_x_mm': self.offsetx_tb.Text.strip() or '0',
            'origin_offset_y_mm': self.offsety_tb.Text.strip() or '0',
            'output_processing': 'Raster' if self.useraster_rb.IsChecked else 'Vector',
            'raster_quality': (
                self.rasterquality_cb.SelectedItem.Content
                if self.rasterquality_cb.SelectedItem
                else 'High'
            ),
            'view_links_in_blue': bool(self.viewlinksinblue_cb.IsChecked),
            'hide_reference_plane': bool(self.hiderefplanes_cb.IsChecked),
            'hide_scope_boxes': bool(self.hidescopeboxes_cb.IsChecked),
            'hide_crop_boundaries': bool(self.hidecropboundaries_cb.IsChecked),
            'hide_unreferenced_view_tags': bool(self.hideunrefviewtags_cb.IsChecked),
            'replace_halftone_with_thin_lines': bool(self.replacehalftone_cb.IsChecked),
            'mask_coincident_lines': bool(self.maskcoincident_cb.IsChecked),
        }
        PrintUtils.set_pdf_overrides(settings)
        self._saved = True
        self.Close()

    def cancel(self, sender, args):
        self.Close()

    def show_dialog(self):
        self.ShowDialog()
        return self._saved


# =============================================================================
# SHEET LIST SOURCE CLASSES
# =============================================================================

class SheetSetList(object):
    """List of sheets from a named Revit Sheet Set."""
    def __init__(self, view_sheetset):
        self.doc = view_sheetset.Document
        self.name = view_sheetset.Name
        self.sheetset = view_sheetset

    def get_sheets(self, doc):
        if doc == self.doc:
            return list(self.sheetset.Views)
        return []


class ScheduleSheetList(object):
    def __init__(self, view_shedule):
        self.doc = view_shedule.Document
        self.name = view_shedule.Name
        self.schedule = view_shedule

    def get_sheets(self, doc):
        return self._get_ordered_schedule_sheets(doc)

    def _get_schedule_text_data(self, view_shedule):
        schedule_data_file = \
            script.get_instance_data_file(str(get_elementid_value(view_shedule.Id)))
        vseop = DB.ViewScheduleExportOptions()
        vseop.TextQualifier = coreutils.get_enum_none(DB.ExportTextQualifier)
        view_shedule.Export(op.dirname(schedule_data_file),
                            op.basename(schedule_data_file),
                            vseop)

        sched_data = []
        try:
            with codecs.open(schedule_data_file, 'r', EXPORT_ENCODING) \
                    as sched_data_file:
                return [x.strip() for x in sched_data_file.readlines()]
        except Exception as open_err:
            logger.error('Error opening sheet index export: %s | %s',
                         schedule_data_file, open_err)
            return sched_data

    def _order_sheets_by_schedule_data(self, view_shedule, sheet_list):
        sched_data = self._get_schedule_text_data(view_shedule)

        if not sched_data:
            return sheet_list

        ordered_sheets_dict = {}
        for sheet in sheet_list:
            logger.debug('finding index for: %s', sheet.SheetNumber)
            for line_no, data_line in enumerate(sched_data):
                match_pattern = r'(^|.*\t){}(\t.*|$)'.format(sheet.SheetNumber)
                matches_sheet = re.match(match_pattern, data_line)
                logger.debug('match: %s', matches_sheet)
                try:
                    if matches_sheet:
                        ordered_sheets_dict[line_no] = sheet
                        break
                    if not sheet.CanBePrinted:
                        logger.debug('Sheet %s is not printable.',
                                     sheet.SheetNumber)
                except Exception:
                    continue

        sorted_keys = sorted(ordered_sheets_dict.keys())
        return [ordered_sheets_dict[x] for x in sorted_keys]

    def _get_ordered_schedule_sheets(self, doc):
        if doc == self.doc:
            sheets = DB.FilteredElementCollector(self.doc,
                                                 self.schedule.Id)\
                    .OfClass(framework.get_type(DB.ViewSheet))\
                    .WhereElementIsNotElementType()\
                    .ToElements()

            return self._order_sheets_by_schedule_data(
                self.schedule,
                sheets
                )
        return []


class AllSheetsList(object):
    @property
    def name(self):
        return "<All Sheets>"

    def get_sheets(self, doc):
        return DB.FilteredElementCollector(doc)\
                 .OfClass(framework.get_type(DB.ViewSheet))\
                 .WhereElementIsNotElementType()\
                 .ToElements()


class UnlistedSheetsList(object):
    @property
    def name(self):
        return "<Unlisted Sheets>"

    def get_sheets(self, doc):
        scheduled_param_id = DB.ElementId(DB.BuiltInParameter.SHEET_SCHEDULED)
        param_prov = DB.ParameterValueProvider(scheduled_param_id)
        param_equality = DB.FilterNumericEquals()
        value_rule = DB.FilterIntegerRule(param_prov, param_equality, 0)
        param_filter = DB.ElementParameterFilter(value_rule)
        return DB.FilteredElementCollector(doc)\
                 .OfClass(framework.get_type(DB.ViewSheet))\
                 .WherePasses(param_filter) \
                 .WhereElementIsNotElementType()\
                 .ToElements()


# =============================================================================
# MAIN PRINT SHEETS WINDOW
# =============================================================================

class PrintSheetsWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)

        self._init_psettings = None
        self._scheduled_sheets = []
        self._original_sheet_order = []
        self._docs_by_hash = {}
        self._doc_cache = {}
        self._suspend_ui_events = True
        self._defer_initial_sheet_load = True

        self.project_info = revit.query.get_project_info(doc=revit.doc)
        self.sheet_cat_id = \
            revit.query.get_category(DB.BuiltInCategory.OST_Sheets).Id
        self.Loaded += self.window_loaded

        self._setup_docs_list()
        self._setup_naming_formats()
        self._setup_export_formats()

        self.folder_tb.Text = PrintUtils.get_dir()
        self._suspend_ui_events = False
        self._load_selected_doc_context()

    # -------------------------------------------------------------------------
    # UI Setup Methods
    # -------------------------------------------------------------------------

    def _setup_docs_list(self):
        if not revit.doc.IsFamilyDocument:
            docs = [AvailableDoc(name=revit.doc.Title,
                                 hash=revit.doc.GetHashCode(),
                                 linked=False)]
            docs.extend([
                AvailableDoc(name=x.Title, hash=x.GetHashCode(), linked=True)
                for x in revit.query.get_all_linkeddocs(doc=revit.doc)
            ])
            self._docs_by_hash = {x.GetHashCode(): x for x in revit.docs}
            self.documents_cb.ItemsSource = docs
            self.documents_cb.SelectedIndex = 0

    def _load_selected_doc_context(self):
        self._suspend_ui_events = True
        try:
            self.project_info = revit.query.get_project_info(doc=self.selected_doc)
            self._apply_projectinfo_naming_format_default()
            self.filename_tb.Text = self._suggest_file_name_from_doc(self.selected_doc)
            self._default_custom_file_name = self.filename_tb.Text
            self._setup_printers()
            self._setup_print_settings()
            self._setup_sheet_list()
        finally:
            self._suspend_ui_events = False

        if not self._defer_initial_sheet_load and self.selected_sheetlist and self.has_print_settings:
            self.sheetlist_changed(None, None)
        else:
            self._scheduled_sheets = []
            self._original_sheet_order = []
            self.sheet_list = []
            self.options_changed(None, None)
        self._all_sheets_list = list(self.sheets_lb.ItemsSource) if self.sheets_lb.ItemsSource else []
        self._update_output_mode_ui()

    def window_loaded(self, sender, args):
        try:
            self.Loaded -= self.window_loaded
        except Exception:
            pass

        self._apply_sheet_grid_headers()

        if self._defer_initial_sheet_load:
            self._defer_initial_sheet_load = False
            if self.selected_sheetlist and self.has_print_settings:
                self.sheetlist_changed(None, None)

    def _get_ui_text(self, key, fallback):
        try:
            value = self.TryFindResource(key)
            return value if value else fallback
        except Exception:
            return fallback

    def _apply_sheet_grid_headers(self):
        try:
            columns = list(self.sheets_lb.Columns)
        except Exception:
            return

        header_map = [
            ('PrintIndex', 'Print Index'),
            ('SheetNumber', 'Sheet Number'),
            ('SheetName', 'Sheet Name'),
            ('Revision', 'Revision'),
            ('PrintSetting', 'Print Setting'),
            ('FileName', 'File Name'),
        ]

        for idx, (key, fallback) in enumerate(header_map):
            if idx < len(columns):
                columns[idx].Header = self._get_ui_text(key, fallback)

    def _get_doc_cache(self, doc):
        doc_hash = doc.GetHashCode()
        if doc_hash not in self._doc_cache:
            self._doc_cache[doc_hash] = {}
        return self._doc_cache[doc_hash]

    def _get_doc_titleblocks(self, doc):
        cache = self._get_doc_cache(doc)
        if 'titleblocks' not in cache:
            cache['titleblocks'] = revit.query.get_elements_by_categories(
                [DB.BuiltInCategory.OST_TitleBlocks],
                doc=doc
            )
        return cache['titleblocks']

    def _get_doc_print_settings(self, doc):
        cache = self._get_doc_cache(doc)
        if 'all_print_settings' not in cache:
            cache['all_print_settings'] = revit.query.get_all_print_settings(doc=doc)
        return cache['all_print_settings']

    def _get_printer_compatible_sizes(self):
        cache = self._get_doc_cache(self.selected_doc)
        printer_name = self.selected_printer or '<none>'
        compatible_sizes_by_printer = cache.setdefault('compatible_sizes_by_printer', {})

        if printer_name == "Revit Internal Printer":
            compatible_sizes_by_printer[printer_name] = None
            return None

        if printer_name not in compatible_sizes_by_printer:
            print_mgr = self._get_printmanager()
            compatible_sizes_by_printer[printer_name] = {x.Name for x in print_mgr.PaperSizes}

        return compatible_sizes_by_printer[printer_name]

    def _get_preferred_pdf_print_setting(self, psetting_items):
        if not psetting_items:
            return None

        preferred_tokens = [
            'pdf',
            'print to pdf',
            'to pdf',
            'adobe pdf',
            'bluebeam',
        ]

        for psetting_item in psetting_items:
            item_name = (psetting_item.name or '').strip().lower()
            if any(token in item_name for token in preferred_tokens):
                return psetting_item

        return None

    def _setup_naming_formats(self):
        self.namingformat_cb.ItemsSource = \
            EditNamingFormatsWindow.get_naming_formats()
        self.namingformat_cb.SelectedIndex = 0

    def _setup_export_formats(self):
        if hasattr(self, 'exportformat_cb'):
            self.exportformat_cb.ItemsSource = ['PDF', 'PDF + DWG', 'PDF + DXF']
            self.exportformat_cb.SelectedIndex = 0

    def _setup_printers(self):
        printers = list(Drawing.Printing.PrinterSettings.InstalledPrinters)

        if IS_REVIT_2022_OR_NEWER:
            printers.insert(0, "Revit Internal Printer")

        self.printers_cb.ItemsSource = printers
        if IS_REVIT_2022_OR_NEWER and "Revit Internal Printer" in printers:
            self.printers_cb.SelectedItem = "Revit Internal Printer"
        else:
            print_mgr = self._get_printmanager()
            self.printers_cb.SelectedItem = print_mgr.PrinterName

    def _setup_print_settings(self):
        psetting_items = \
            self._get_psetting_items(
                doc=self.selected_doc,
                psettings=self._get_doc_print_settings(self.selected_doc),
                include_varsettings=not self.selected_doc.IsLinked
                )
        self.printsettings_cb.ItemsSource = psetting_items
        preferred_pdf_psetting = self._get_preferred_pdf_print_setting(psetting_items)

        print_mgr = self._get_printmanager()
        if isinstance(print_mgr.PrintSetup.CurrentPrintSetting,
                      DB.InSessionPrintSetting):
            in_session = PrintSettingListItem(
                print_mgr.PrintSetup.CurrentPrintSetting
                )
            psetting_items.append(in_session)
            self.printsettings_cb.SelectedItem = preferred_pdf_psetting or in_session
        else:
            self._init_psettings = print_mgr.PrintSetup.CurrentPrintSetting
            cur_psetting_name = print_mgr.PrintSetup.CurrentPrintSetting.Name
            if preferred_pdf_psetting:
                self.printsettings_cb.SelectedItem = preferred_pdf_psetting
                cur_psetting_name = None
            for psetting_item in psetting_items:
                if cur_psetting_name and psetting_item.name == cur_psetting_name:
                    self.printsettings_cb.SelectedItem = psetting_item

        if self.selected_doc.IsLinked:
            self.disable_element(self.printsettings_cb)
        else:
            self.enable_element(self.printsettings_cb)

        self._update_combine_option()

    def _setup_sheet_list(self):
        cache = self._get_doc_cache(self.selected_doc)
        sheet_indices = list(cache.get('sheet_indices', self._get_sheet_index_list()))
        if 'sheet_indices' not in cache:
            cache['sheet_indices'] = list(sheet_indices)
        try:
            if 'sheetsets' not in cache:
                cl = DB.FilteredElementCollector(self.selected_doc)
                cache['sheetsets'] = cl.OfClass(
                    framework.get_type(DB.ViewSheetSet)
                ).WhereElementIsNotElementType().ToElements()
            sheetsets = cache['sheetsets']
            for ss in sheetsets:
                sheet_indices.append(SheetSetList(ss))
        except Exception as e:
            logger.warning("Could not load sheet sets: {}".format(e))

        special_lists = [AllSheetsList(), UnlistedSheetsList()]
        sorted_indices = sorted(sheet_indices, key=lambda x: sheet_index_sort_key(x.name))
        self.schedules_cb.ItemsSource = sorted_indices + special_lists
        self.schedules_cb.SelectedIndex = 0
        if self.schedules_cb.ItemsSource:
            self.enable_element(self.schedules_cb)
        else:
            self.disable_element(self.schedules_cb)

    def _get_output_mode_description(self):
        def res(key, fallback):
            try:
                value = self.TryFindResource(key)
                return value if value else fallback
            except Exception:
                return fallback

        overrides_enabled = PrintUtils.get_pdf_overrides().get('enabled', False)
        if IS_REVIT_2022_OR_NEWER and self.selected_printer == "Revit Internal Printer":
            if overrides_enabled:
                return res('OutputModeNativeOverrides', "Using Native PDF Export + Overrides")
            return res('OutputModeNative', "Using Native PDF Export")
        return res('OutputModePrinter', "Using Printer / PrintManager")

    def _update_output_mode_ui(self):
        if hasattr(self, 'outputmode_tb'):
            self.outputmode_tb.Text = self._get_output_mode_description()
        if hasattr(self, 'pdfoptions_b'):
            self.pdfoptions_b.IsEnabled = bool(
                IS_REVIT_2022_OR_NEWER
                and self.selected_printer == "Revit Internal Printer"
            )

    # -------------------------------------------------------------------------
    # Properties
    # -------------------------------------------------------------------------

    def copy_naming_format(self, sender, args):
        try:
            naming_format = sender.DataContext
            if not naming_format:
                return

            script.clipboard_copy(naming_format.name)

        except Exception as e:
            logger.error("Failed to copy naming format: %s", e)

    def _apply_projectinfo_naming_format_default(self):
        pi = self.selected_doc.ProjectInformation
        param = pi.LookupParameter("Naming Format") if pi else None
        param_value = param.AsString() if param else None

        selected_item = next(
            (nf for nf in self.namingformat_cb.ItemsSource if nf.name == param_value),
            None
        )

        if not selected_item and self.namingformat_cb.ItemsSource:
            selected_item = self.namingformat_cb.ItemsSource[0]

        self.namingformat_cb.SelectedItem = selected_item

    @property
    def selected_doc(self):
        selected_doc = self.documents_cb.SelectedItem
        if not selected_doc:
            return None
        return self._docs_by_hash.get(selected_doc.hash)

    @property
    def selected_sheetlist(self):
        return self.schedules_cb.SelectedItem

    @property
    def has_errors(self):
        return self.errormsg_tb.Text != ''

    @property
    def reverse_print(self):
        return self.reverse_cb.IsChecked

    @property
    def combine_print(self):
        return self.combine_cb.IsChecked

    @property
    def show_placeholders(self):
        return self.placeholder_cb.IsChecked

    @property
    def index_digits(self):
        return int(self.index_slider.Value)

    @property
    def index_start(self):
        return int(self.indexstart_tb.Text or 0)

    @property
    def include_placeholders(self):
        return self.indexspace_cb.IsChecked

    @property
    def output_folder(self):
        folder = self.folder_tb.Text.strip() if hasattr(self, 'folder_tb') and self.folder_tb.Text else ''
        return folder or PrintUtils.get_dir()

    @property
    def custom_file_name(self):
        return self.filename_tb.Text.strip() if hasattr(self, 'filename_tb') and self.filename_tb.Text else ''

    @property
    def selected_export_format(self):
        if hasattr(self, 'exportformat_cb') and self.exportformat_cb.SelectedItem:
            return self.exportformat_cb.SelectedItem
        return 'PDF'

    @property
    def export_dwg_enabled(self):
        return self.selected_export_format == 'PDF + DWG'

    @property
    def export_dxf_enabled(self):
        return self.selected_export_format == 'PDF + DXF'

    @property
    def export_extra_enabled(self):
        return self.export_dwg_enabled or self.export_dxf_enabled

    @property
    def selected_naming_format(self):
        return self.namingformat_cb.SelectedItem

    @property
    def selected_printer(self):
        return self.printers_cb.SelectedItem

    @property
    def selected_print_setting(self):
        return self.printsettings_cb.SelectedItem

    @property
    def has_print_settings(self):
        # self.selected_print_setting implements __nonzero__
        # manually check None-ness
        return self.selected_print_setting is not None

    @property
    def print_settings(self):
        return self.printsettings_cb.ItemsSource

    @property
    def sheet_list(self):
        return self.sheets_lb.ItemsSource

    @sheet_list.setter
    def sheet_list(self, value):
        self.sheets_lb.ItemsSource = value

    @property
    def selected_sheets(self):
        return self.sheets_lb.SelectedItems

    @property
    def printable_sheets(self):
        return [x for x in self.sheet_list if x.printable]

    @property
    def selected_printable_sheets(self):
        return [x for x in self.selected_sheets if x.printable]

    # -------------------------------------------------------------------------
    # Private Helper Methods
    # -------------------------------------------------------------------------

    def _is_sheet_index(self, schedule_view):
        return self.sheet_cat_id == schedule_view.Definition.CategoryId \
               and not schedule_view.IsTemplate

    def _get_sheet_index_list(self):
        schedules = DB.FilteredElementCollector(self.selected_doc)\
                      .OfClass(framework.get_type(DB.ViewSchedule))\
                      .WhereElementIsNotElementType()\
                      .ToElements()

        sheet_indices = [
            ScheduleSheetList(s) for s in schedules
            if self._is_sheet_index(s)
            ]

        result = sorted(sheet_indices, key=lambda x: (x.name or '').lower())
        return result

    def _get_printmanager(self):
        try:
            return self.selected_doc.PrintManager
        except Exception as printerr:
            logger.critical('Error getting printer manager from document. '
                            'Most probably there is not a printer defined '
                            'on your system. | %s', printerr)
            script.exit()

    def _get_psetting_items(self, doc,
                            psettings=None, include_varsettings=False):
        if include_varsettings:
            psetting_items = [VariablePaperPrintSettingListItem()]
        else:
            psetting_items = []

        psettings = psettings or revit.query.get_all_print_settings(doc=doc)
        psetting_items.extend([PrintSettingListItem(x) for x in psettings])

        compatible_sizes = self._get_printer_compatible_sizes()
        if compatible_sizes is None:
            for psetting_item in psetting_items:
                if isinstance(psetting_item, PrintSettingListItem):
                    psetting_item.is_compatible = True
        else:
            for psetting_item in psetting_items:
                if isinstance(psetting_item, PrintSettingListItem):
                    if psetting_item.paper_size \
                            and psetting_item.paper_size.Name in compatible_sizes:
                        psetting_item.is_compatible = True
        return psetting_items

    def _update_combine_option(self):
        self.enable_element(self.combine_cb)
        if self.selected_doc.IsLinked \
                or ((self.selected_sheetlist and self.has_print_settings)
                    and self.selected_print_setting.allows_variable_paper):
            self.disable_element(self.combine_cb)
            self.combine_cb.IsChecked = False

    def _strip_username_suffix(self, model_name):
        value = (model_name or '').strip()
        if not value:
            return ''

        username = ''
        try:
            username = (HOST_APP.username or '').strip()
        except Exception:
            username = ''

        if username:
            patterns = [
                r'[_\- ]{}$'.format(re.escape(username)),
                r'\.{}$'.format(re.escape(username)),
                r'_{}$'.format(re.escape(username)),
                r'-{}$'.format(re.escape(username)),
            ]
            for pattern in patterns:
                new_value = re.sub(pattern, '', value, flags=re.IGNORECASE)
                if new_value != value:
                    value = new_value
                    break

        value = re.sub(r'\s+', ' ', value).strip(' _-.')
        return value

    def _suggest_file_name_from_doc(self, doc):
        doc_title = ''
        if doc:
            try:
                doc_title = op.splitext(doc.Title)[0]
            except Exception:
                doc_title = ''

        doc_title = self._strip_username_suffix(doc_title)
        doc_title = coreutils.cleanup_filename(doc_title, windows_safe=True).strip()
        if not doc_title:
            doc_title = 'Ordered Sheet Set'
        return doc_title

    def _apply_suggested_filename(self):
        suggested_name = self._suggest_file_name_from_doc(self.selected_doc)
        current_name = self.filename_tb.Text.strip() if self.filename_tb.Text else ''
        previous_default = getattr(self, '_default_custom_file_name', '')

        if not current_name or current_name == previous_default:
            self.filename_tb.Text = suggested_name

        self._default_custom_file_name = suggested_name

    def _get_extra_export_extension(self):
        if self.export_dwg_enabled:
            return '.dwg'
        if self.export_dxf_enabled:
            return '.dxf'
        return ''

    def _get_extra_export_label(self):
        if self.export_dwg_enabled:
            return 'DWG'
        if self.export_dxf_enabled:
            return 'DXF'
        return ''

    def _get_output_dir(self):
        dir_path = self.output_folder
        if not dir_path:
            dir_path = PrintUtils.get_dir()
        PrintUtils.ensure_dir(dir_path)
        return dir_path

    def _get_combined_pdf_name(self):
        filename = ensure_pdf_extension(self.custom_file_name)
        if not filename:
            filename = ensure_pdf_extension(self._suggest_file_name_from_doc(self.selected_doc))
        return coreutils.cleanup_filename(filename, windows_safe=True)

    def _apply_filename_prefix(self, filename):
        prefix = self.custom_file_name.strip()
        default_prefix = getattr(self, '_default_custom_file_name', '').strip()
        if not prefix or prefix == default_prefix:
            return filename

        prefix = coreutils.cleanup_filename(prefix, windows_safe=True).strip()
        if not prefix:
            return filename

        base_name, ext = op.splitext(filename)
        if not ext:
            ext = '.pdf'
        return '{}_{}{}'.format(prefix, base_name, ext)

    def _resolve_output_filepath(self, sheet_name, target_filepath):
        if not op.exists(target_filepath):
            return target_filepath

        filename = op.basename(target_filepath)
        folder = op.dirname(target_filepath)
        base_name, ext = op.splitext(filename)

        decision = forms.alert(
            'The file "{}" already exists.\n\nWhat would you like to do?'.format(filename),
            options=['Overwrite', 'Rename', 'Skip']
        )

        if decision == 'Skip' or not decision:
            logger.warning(
                'Skipping sheet "%s". File already exists at %s.',
                sheet_name,
                target_filepath
            )
            return None

        if decision == 'Overwrite':
            return target_filepath

        if decision == 'Rename':
            while True:
                new_name = forms.ask_for_string(
                    default='{}_NEW{}'.format(base_name, ext),
                    prompt='Enter a new file name:',
                    title='Rename Output File'
                )

                if not new_name:
                    logger.warning(
                        'Skipping sheet "%s". Rename cancelled by user.',
                        sheet_name
                    )
                    return None

                new_name = new_name.strip()
                if not new_name:
                    continue

                if not new_name.lower().endswith(ext.lower()):
                    new_name += ext

                new_name = coreutils.cleanup_filename(new_name, windows_safe=True)
                if not new_name:
                    continue

                new_filepath = op.join(folder, new_name)

                if not op.exists(new_filepath):
                    return new_filepath

                retry = forms.alert(
                    'The new file name also already exists.\n\nDo you want to enter another name?',
                    yes=True,
                    no=True,
                    ok=False
                )
                if not retry:
                    logger.warning(
                        'Skipping sheet "%s". Alternate file name already exists.',
                        sheet_name
                    )
                    return None

        return None


    def _reset_psettings(self):
        if self._init_psettings:
            print_mgr = self._get_printmanager()
            with revit.Transaction("Revert to Original Print Settings"):
                print_mgr.PrintSetup.CurrentPrintSetting = self._init_psettings

    def _update_index_slider(self):
        index_digits = \
            coreutils.get_integer_length(
                len(self._scheduled_sheets) + self.index_start
                )
        self.index_slider.Minimum = max([index_digits, 2])
        self.index_slider.Maximum = self.index_slider.Minimum + 3

    def _update_print_indices(self, sheet_list):
        start_idx = self.index_start
        for idx, sheet in enumerate(sheet_list):
            sheet.print_index = INDEX_FORMAT\
                .format(digits=self.index_digits)\
                .format(idx + start_idx)

    def _sort_sheet_list(self, sheet_list):
        return sorted(
            sheet_list,
            key=lambda sheet: (
                get_discipline_rank(sheet.number),
                sheet_number_sort_key(sheet.number),
                (sheet.name or '').lower()
            )
        )

    def _refresh_sheet_order_ui(self, sheet_list):
        if not self.show_placeholders:
            self.indexspace_cb.IsEnabled = True
            self._update_print_indices(sheet_list)

            printable_sheets = []
            for sheet in sheet_list:
                if sheet.printable:
                    printable_sheets.append(sheet)

            if not self.include_placeholders:
                self._update_print_indices(printable_sheets)
            self.sheet_list = printable_sheets
        else:
            self.indexspace_cb.IsChecked = True
            self.indexspace_cb.IsEnabled = False
            self._update_print_indices(sheet_list)
            self.sheet_list = sheet_list

        self._update_print_filenames(sheet_list)

    # -------------------------------------------------------------------------
    # Filename Template Resolution
    # -------------------------------------------------------------------------

    def _update_filename_template(self, template, value_type, value_getter):
        finder_pattern = r'{' + value_type + r':(.*?)}'
        for param_name in re.findall(finder_pattern, template):
            param_value = value_getter(param_name)
            repl_pattern = r'{' + value_type + ':' + param_name + r'}'
            if param_value:
                template = re.sub(repl_pattern, str(param_value), template)
            template = re.sub(repl_pattern, '', template)
        return template

    def _update_print_filename(self, template, sheet):
        # resolve sheet-level custom param values
        # get titleblock param values
        template = self._update_filename_template(
            template=template,
            value_type='tblock_param',
            value_getter=lambda x: revit.query.get_param_value(
                    revit.query.get_param(sheet.revit_tblock, x)
                ) or revit.query.get_param_value(
                    revit.query.get_param(sheet.revit_tblock_type, x)
                )
        )

        # get sheet param values
        template = self._update_filename_template(
            template=template,
            value_type='sheet_param',
            value_getter=lambda x: revit.query.get_param_value(
                revit.query.get_param(sheet.revit_sheet, x)
                )
        )

        # get date for sortable list
        rev_date_str = sheet.revision.date or ""
        sortable_date = ""

        # Try to detect user's locale
        locale_tuple = locale.getdefaultlocale()
        user_locale = (locale_tuple[0] if locale_tuple and locale_tuple[0] else "en_GB")
        dayfirst = not user_locale.startswith("en_US")

        # Try several common patterns
        date_formats = ["%d.%m.%y", "%m.%d.%y", "%d/%m/%y", "%m/%d/%y"]
        if not dayfirst:
            date_formats = ["%m.%d.%y", "%m/%d/%y", "%d.%m.%y", "%d/%m/%y"]

        for fmt in date_formats:
            try:
                parsed = datetime.datetime.strptime(rev_date_str, fmt)
                sortable_date = parsed.strftime("%Y%m%d")
                break
            except (ValueError, TypeError):
                continue

        sheet.revision_date_sortable = sortable_date
        # resolved the fixed formatters
        try:
            output_fname = \
                template.format(
                    index=sheet.print_index,
                    number=sheet.number,
                    name=sheet.name,
                    name_dash=sheet.name.replace(' ', '-'),
                    name_underline=sheet.name.replace(' ', '_'),
                    current_date=coreutils.current_date(),
                    issue_date=sheet.issue_date,
                    rev_number=sheet.revision.number if sheet.revision else '',
                    rev_desc=sheet.revision.desc if sheet.revision else '',
                    rev_date=sheet.revision.date if sheet.revision else '',
                    proj_name=self.project_info.name,
                    proj_number=self.project_info.number,
                    proj_building_name=self.project_info.building_name,
                    proj_issue_date=self.project_info.issue_date,
                    proj_org_name=self.project_info.org_name,
                    proj_status=self.project_info.status,
                    username=HOST_APP.username,
                    revit_version=HOST_APP.version,
                )
        except Exception as ferr:
            output_fname = ''
            if isinstance(ferr, KeyError):
                self._set_error('Unknown key in selected naming format')
        # and set the sheet file name
        sheet.print_filename = output_fname

    def _update_print_filenames(self, sheet_list):
        doc = self.selected_doc
        naming_fmt = self.selected_naming_format
        if naming_fmt:
            template = naming_fmt.template
            # resolve project-level custom param values
            # project info param values
            template = self._update_filename_template(
                template=template,
                value_type='proj_param',
                value_getter=lambda x: revit.query.get_param_value(
                    doc.ProjectInformation.LookupParameter(x)
                    )
            )

            # global param values
            template = self._update_filename_template(
                template=template,
                value_type='glob_param',
                value_getter=lambda x: revit.query.get_param_value(
                    revit.query.get_global_parameter(x, doc=doc)
                    )
            )

            for sheet in sheet_list:
                self._update_print_filename(template, sheet)
                sheet.print_filename = self._apply_filename_prefix(sheet.print_filename)

    # -------------------------------------------------------------------------
    # Print Settings & Titleblock Analysis
    # -------------------------------------------------------------------------

    def _find_sheet_tblock(self, revit_sheet, tblocks):
        for tblock in tblocks:
            view_sheet = revit_sheet.Document.GetElement(tblock.OwnerViewId)
            if view_sheet.Id == revit_sheet.Id:
                return tblock

    def _get_sheet_printsettings(self, tblocks, psettings):
        tblock_printsettings = {}
        sheet_printsettings = {}
        for tblock in tblocks:
            tblock_psetting = None
            sheet = self.selected_doc.GetElement(tblock.OwnerViewId)
            # build a unique id for this tblock
            tblock_tform = tblock.GetTotalTransform()
            tblock_tid = get_elementid_value(tblock.GetTypeId())
            tblock_tid = tblock_tid * 100 + tblock_tform.BasisX.X * 10 + tblock_tform.BasisX.Y
            # can not use None as default. see notes below
            tblock_psetting = tblock_printsettings.get(tblock_tid, None)
            # if found a tblock print settings, assign that to sheet
            if tblock_psetting:
                sheet_printsettings[sheet.SheetNumber] = tblock_psetting
            # otherwise, analyse the tblock and determine print settings
            else:
                # try the type parameter "Print Setting"
                tblock_type = tblock.Document.GetElement(tblock.GetTypeId())
                if tblock_type:
                    psparam = tblock_type.LookupParameter("Print Setting")
                    if psparam:
                        psetting_name = psparam.AsString()
                        psparam_psetting = \
                            next(
                                (x for x in psettings
                                    if x.Name == psetting_name),
                                None
                            )
                        if psparam_psetting:
                            tblock_psetting = \
                                TitleBlockPrintSettings(
                                    psettings=[psparam_psetting],
                                    set_by_param=True
                                )
                # otherwise, try to detect applicable print settings
                # based on title block geometric properties
                if not tblock_psetting:
                    tblock_psetting = \
                        TitleBlockPrintSettings(
                            psettings=revit.query.get_titleblock_print_settings(
                                tblock,
                                self.selected_printer,
                                psettings
                                ),
                            set_by_param=False
                        )
                # the analysis result might be None
                tblock_printsettings[tblock_tid] = tblock_psetting
                sheet_printsettings[sheet.SheetNumber] = tblock_psetting
        return sheet_printsettings

    # -------------------------------------------------------------------------
    # Sheet Set Management (Add / Remove / Save)
    # -------------------------------------------------------------------------

    def _get_existing_sheet_sets(self):
        existing_sets = {}
        try:
            collector = DB.FilteredElementCollector(self.selected_doc)
            sheetsets = collector.OfClass(framework.get_type(DB.ViewSheetSet))\
                                .WhereElementIsNotElementType()\
                                .ToElements()
            for sheetset in sheetsets:
                existing_sets[sheetset.Name] = sheetset
        except Exception as read_err:
            logger.error('Failed to read existing sheet sets: %s', read_err)
        return existing_sets

    def _generate_unique_sheetset_name(self, base_name, existing_names):
        base_name = (base_name or 'Custom Sheet Set').strip()
        if base_name not in existing_names:
            return base_name

        index = 2
        while True:
            candidate = '{} ({})'.format(base_name, index)
            if candidate not in existing_names:
                return candidate
            index += 1

    def _set_selected_sheetlist_by_name(self, set_name):
        if not set_name or not self.schedules_cb.ItemsSource:
            return

        for item in self.schedules_cb.ItemsSource:
            if getattr(item, 'name', None) == set_name:
                self.schedules_cb.SelectedItem = item
                return

    def _build_viewset_from_sheet_items(self, sheet_items):
        view_set = DB.ViewSet()
        for sheet in sheet_items:
            try:
                if sheet and sheet.revit_sheet:
                    view_set.Insert(sheet.revit_sheet)
            except Exception as insert_err:
                logger.error('Failed to add sheet %s to view set: %s', getattr(sheet, 'number', '---'), insert_err)
        return view_set

    def add_sheets(self, sender=None, args=None):
        if self.selected_doc.IsLinked:
            forms.alert('Add Sheets is not available for linked models.')
            return

        current_sheet_ids = set()
        for sheet in list(self.sheet_list or []):
            try:
                current_sheet_ids.add(sheet.revit_sheet.Id.IntegerValue)
            except Exception:
                pass

        tblocks = revit.query.get_elements_by_categories(
            [DB.BuiltInCategory.OST_TitleBlocks],
            doc=self.selected_doc
        )
        rev_cfg = DB.RevisionSettings.GetRevisionSettings(revit.doc)
        current_print_setting = self.selected_print_setting.print_settings if self.has_print_settings and not self.selected_print_setting.allows_variable_paper else None

        available_sheets = []
        for sheet in revit.query.get_sheets(doc=self.selected_doc):
            try:
                if sheet.Id.IntegerValue not in current_sheet_ids:
                    available_sheets.append(sheet)
            except Exception:
                continue

        available_sheets = sorted(available_sheets, key=lambda x: (natural_sort_key(x.SheetNumber), (x.Name or '').lower()))
        selected_revit_sheets = forms.SelectFromList.show(
            available_sheets,
            name_attr='SheetNumber',
            multiselect=True,
            button_name='Add Selected Sheets',
            title='Add Sheets to Current List',
            width=500,
            height=650
        )

        if not selected_revit_sheets:
            return

        sheet_items = list(self.sheet_list or [])
        for revit_sheet in selected_revit_sheets:
            tblock = self._find_sheet_tblock(revit_sheet, tblocks)
            if self.has_print_settings and self.selected_print_setting.allows_variable_paper:
                sheet_printsettings = self._get_sheet_printsettings(
                    tblocks,
                    revit.query.get_all_print_settings(doc=self.selected_doc)
                )
                tb_printsettings = sheet_printsettings.get(
                    revit_sheet.SheetNumber,
                    TitleBlockPrintSettings(psettings=[], set_by_param=False)
                )
            else:
                tb_printsettings = TitleBlockPrintSettings(
                    psettings=[current_print_setting] if current_print_setting else [],
                    set_by_param=False
                )

            sheet_items.append(
                ViewSheetListItem(
                    view_sheet=revit_sheet,
                    view_tblock=tblock,
                    print_settings=tb_printsettings,
                    rev_settings=rev_cfg
                )
            )

        self._scheduled_sheets = list(sheet_items)
        self._original_sheet_order = list(sheet_items)
        self._refresh_sheet_order_ui(sheet_items)
        self._all_sheets_list = list(self.sheets_lb.ItemsSource) if self.sheets_lb.ItemsSource else []

    def remove_selected(self, sender=None, args=None):
        selected = list(self.selected_sheets)
        if not selected:
            return

        selected_ids = set()
        for sheet in selected:
            try:
                selected_ids.add(sheet.revit_sheet.Id.IntegerValue)
            except Exception:
                pass

        sheet_items = []
        for sheet in list(self.sheet_list or []):
            try:
                if sheet.revit_sheet.Id.IntegerValue not in selected_ids:
                    sheet_items.append(sheet)
            except Exception:
                sheet_items.append(sheet)

        self._scheduled_sheets = list(sheet_items)
        self._original_sheet_order = list(sheet_items)
        self._refresh_sheet_order_ui(sheet_items)
        self._all_sheets_list = list(self.sheets_lb.ItemsSource) if self.sheets_lb.ItemsSource else []

    def save_current_set(self, sender=None, args=None):
        if self.selected_doc.IsLinked:
            forms.alert('Save Current Set is not available for linked models.')
            return

        current_sheets = list(self.sheet_list or [])
        if not current_sheets:
            forms.alert('There are no sheets in the current list to save.')
            return

        suggested_name = None
        selected_source = self.selected_sheetlist
        if selected_source and hasattr(selected_source, 'name'):
            source_name = (selected_source.name or '').strip()
            if source_name and not source_name.startswith('<'):
                suggested_name = source_name

        if not suggested_name:
            suggested_name = 'Custom Sheet Set'

        set_name = forms.ask_for_string(
            default=suggested_name,
            prompt='Name for the Revit sheet set:',
            title='Save Current Set'
        )
        if not set_name:
            return

        set_name = set_name.strip()
        if not set_name:
            return

        view_set = self._build_viewset_from_sheet_items(current_sheets)
        existing_sets = self._get_existing_sheet_sets()
        target_name = set_name
        existing_set = existing_sets.get(set_name)

        if existing_set:
            decision = forms.alert(
                'A sheet set named "{}" already exists. What would you like to do?'.format(set_name),
                options=['Overwrite', 'Save As New', 'Cancel']
            )

            if decision == 'Cancel' or not decision:
                return
            elif decision == 'Save As New':
                target_name = self._generate_unique_sheetset_name(set_name, existing_sets.keys())
            else:
                print_mgr = self._get_printmanager()
                t = DB.Transaction(self.selected_doc, 'Overwrite Sheet Set')
                t.Start()
                try:
                    print_mgr.ViewSheetSetting.CurrentViewSheetSet = existing_set
                    print_mgr.ViewSheetSetting.Delete()
                    t.Commit()
                except Exception as delete_err:
                    if t.HasStarted():
                        t.RollBack()
                    forms.alert('Failed to overwrite existing sheet set.', expanded=str(delete_err))
                    return

        print_mgr = self._get_printmanager()
        t = DB.Transaction(self.selected_doc, 'Save Current Sheet Set')
        t.Start()
        try:
            view_sheet_setting = print_mgr.ViewSheetSetting
            view_sheet_setting.CurrentViewSheetSet.Views = view_set
            view_sheet_setting.SaveAs(target_name)
            t.Commit()
        except Exception as save_err:
            if t.HasStarted():
                t.RollBack()
            forms.alert('Failed to save current sheet set.', expanded=str(save_err))
            return

        self._setup_sheet_list()
        self._set_selected_sheetlist_by_name(target_name)
        forms.alert('Sheet set saved as "{}".'.format(target_name))

    # -------------------------------------------------------------------------
    # Sheet Order Management
    # -------------------------------------------------------------------------

    def _get_print_order_param(self, sheet):
        try:
            return sheet.revit_sheet.LookupParameter("PRINT_ORDER")
        except Exception:
            return None

    def _get_saved_print_order(self, sheet):
        param = self._get_print_order_param(sheet)
        if param and param.HasValue:
            try:
                return param.AsInteger()
            except Exception:
                return None
        return None

    def _has_saved_print_order(self, sheet):
        return self._get_saved_print_order(sheet) is not None

    def move_up(self, sender, args):
        selected = list(self.selected_sheets)
        if not selected:
            return

        sheet_list = list(self.sheet_list)
        moved = False

        for sheet in selected:
            idx = sheet_list.index(sheet)
            if idx > 0 and sheet_list[idx - 1] not in selected:
                sheet_list[idx], sheet_list[idx - 1] = sheet_list[idx - 1], sheet_list[idx]
                moved = True

        if moved:
            self._refresh_sheet_order_ui(sheet_list)
            self.sheets_lb.SelectedItems.Clear()
            for sheet in selected:
                self.sheets_lb.SelectedItems.Add(sheet)

    def move_down(self, sender, args):
        selected = list(self.selected_sheets)
        if not selected:
            return

        sheet_list = list(self.sheet_list)
        moved = False

        for sheet in reversed(selected):
            idx = sheet_list.index(sheet)
            if idx < len(sheet_list) - 1 and sheet_list[idx + 1] not in selected:
                sheet_list[idx], sheet_list[idx + 1] = sheet_list[idx + 1], sheet_list[idx]
                moved = True

        if moved:
            self._refresh_sheet_order_ui(sheet_list)
            self.sheets_lb.SelectedItems.Clear()
            for sheet in selected:
                self.sheets_lb.SelectedItems.Add(sheet)

    def move_to_top(self, sender=None, args=None):
        selected = list(self.selected_sheets)
        if not selected:
            return

        sheet_list = list(self.sheet_list)
        selected_set = set(selected)
        reordered = selected + [sheet for sheet in sheet_list if sheet not in selected_set]

        self._refresh_sheet_order_ui(reordered)

        try:
            self.sheets_lb.SelectedItems.Clear()
            for sheet in selected:
                self.sheets_lb.SelectedItems.Add(sheet)
        except Exception:
            pass

    def sort_by_sheet_number(self, sender=None, args=None):
        if not self.sheet_list:
            return

        sheet_list = sorted(
            list(self.sheet_list),
            key=lambda sheet: (
                natural_sort_key(sheet.number),
                (sheet.name or '').lower()
            )
        )

        if self.reverse_print:
            sheet_list.reverse()

        self._refresh_sheet_order_ui(sheet_list)

        try:
            self.sheets_lb.SelectedItems.Clear()
        except Exception:
            pass

    def reset_to_standard(self, sender=None, args=None):
        base_sheet_list = list(self._original_sheet_order) if self._original_sheet_order else [x for x in self._scheduled_sheets]
        sheet_list = self._sort_sheet_list(base_sheet_list)

        if self.reverse_print:
            sheet_list.reverse()

        self._refresh_sheet_order_ui(sheet_list)

        try:
            self.sheets_lb.SelectedItems.Clear()
        except Exception:
            pass

    def save_print_order(self, sender=None, args=None):
        if not self.sheet_list:
            return

        if not self._ensure_print_order_parameter():
            forms.alert('Could not create or find a writable PRINT_ORDER parameter on sheets.')
            return

        sheets_to_save = list(self.sheet_list)
        writable_params = []
        for sheet in sheets_to_save:
            param = self._get_print_order_param(sheet)
            if param and not param.IsReadOnly:
                writable_params.append(param)

        if not writable_params:
            forms.alert('PRINT_ORDER exists but is not writable on these sheets.')
            return

        with revit.Transaction('Save Print Order', doc=self.selected_doc):
            for idx, sheet in enumerate(sheets_to_save):
                param = self._get_print_order_param(sheet)
                if param and not param.IsReadOnly:
                    try:
                        param.Set(idx)
                    except Exception as set_err:
                        logger.error('Failed to set PRINT_ORDER on sheet %s: %s', sheet.number, set_err)

        forms.alert('Print order saved to parameter PRINT_ORDER.')

    def _ensure_print_order_parameter(self):
        sample_sheet = None
        try:
            sample_sheet = next(iter(revit.query.get_sheets(doc=self.selected_doc)), None)
        except Exception:
            sample_sheet = None

        if sample_sheet:
            existing_param = sample_sheet.LookupParameter("PRINT_ORDER")
            if existing_param:
                return True

        app = self.selected_doc.Application
        original_spfile = app.SharedParametersFilename
        temp_spfile = tempfile.NamedTemporaryFile(prefix='pyrevit_print_order_', suffix='.txt', delete=False)
        temp_spfile.close()

        try:
            app.SharedParametersFilename = temp_spfile.name
            with open(temp_spfile.name, 'w') as spf:
                spf.write('# This is a Revit shared parameter file.\n')
                spf.write('*META\tVERSION\tMINVERSION\n')
                spf.write('META\t2\t1\n')
                spf.write('*GROUP\tID\tNAME\n')
                spf.write('*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE\n')

            sp_file = app.OpenSharedParameterFile()
            if sp_file is None:
                return False

            group = None
            for grp in sp_file.Groups:
                if grp.Name == 'pyRevit':
                    group = grp
                    break
            if group is None:
                group = sp_file.Groups.Create('pyRevit')

            definition = None
            for dfn in group.Definitions:
                if dfn.Name == 'PRINT_ORDER':
                    definition = dfn
                    break

            if definition is None:
                try:
                    options = DB.ExternalDefinitionCreationOptions('PRINT_ORDER', DB.SpecTypeId.Int.Integer)
                except Exception:
                    options = DB.ExternalDefinitionCreationOptions('PRINT_ORDER', DB.ParameterType.Integer)
                definition = group.Definitions.Create(options)

            catset = app.Create.NewCategorySet()
            catset.Insert(self.selected_doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Sheets))
            binding = app.Create.NewInstanceBinding(catset)

            with revit.Transaction('Create PRINT_ORDER Parameter', doc=self.selected_doc):
                try:
                    success = self.selected_doc.ParameterBindings.Insert(definition, binding, DB.BuiltInParameterGroup.PG_IDENTITY_DATA)
                except Exception:
                    success = self.selected_doc.ParameterBindings.Insert(definition, binding)

            self.selected_doc.Regenerate()

            check_sheet = sample_sheet
            if check_sheet is None:
                try:
                    sheets = list(revit.query.get_sheets(doc=self.selected_doc))
                    if sheets:
                        check_sheet = sheets[0]
                except Exception:
                    check_sheet = None

            if check_sheet:
                existing_param = check_sheet.LookupParameter("PRINT_ORDER")
                if existing_param:
                    return True

            return bool(success)
        except Exception as create_err:
            logger.error('Failed to create PRINT_ORDER parameter: %s', create_err)
            return False
        finally:
            app.SharedParametersFilename = original_spfile
            try:
                os.remove(temp_spfile.name)
            except Exception:
                pass

    # -------------------------------------------------------------------------
    # Search
    # -------------------------------------------------------------------------

    def sheet_search_changed(self, sender, args):
        search_text = self.sheetsearch_tb.Text.strip().lower()
        if not self._all_sheets_list:
            return

        if not search_text:
            self.sheets_lb.ItemsSource = self._all_sheets_list
        else:
            filtered = []
            for sheet in self._all_sheets_list:
                number = sheet.number.lower() if sheet.number else ''
                name = sheet.name.lower() if sheet.name else ''
                if search_text in number or search_text in name:
                    filtered.append(sheet)
            self.sheets_lb.ItemsSource = filtered

    # -------------------------------------------------------------------------
    # Output Folder
    # -------------------------------------------------------------------------

    def pick_output_folder(self, sender=None, args=None):
        folder = forms.pick_folder()
        if folder:
            self.folder_tb.Text = folder

    # -------------------------------------------------------------------------
    # Error Handling
    # -------------------------------------------------------------------------

    def _reset_error(self):
        self.enable_element(self.print_b)
        self.hide_element(self.errormsg_block)
        self.errormsg_tb.Text = ''

    def _set_error(self, err_msg):
        if self.errormsg_tb.Text != err_msg:
            self.disable_element(self.print_b)
            self.show_element(self.errormsg_block)
            self.errormsg_tb.Text = err_msg

    # -------------------------------------------------------------------------
    # Print Execution — Combined (single file)
    # -------------------------------------------------------------------------

    def _print_combined_sheets_in_order(self, target_sheets):
        # Use Revit native PDF export for combined output only when the
        # internal Revit PDF engine is the selected output on Revit 2022+.
        if IS_REVIT_2022_OR_NEWER and self.selected_printer == "Revit Internal Printer":
            dirPath = self._get_output_dir()
            PrintUtils.open_dir(dirPath)

            export_sheets = List[DB.ElementId]()
            printable_sheets = []
            for sheet in target_sheets:
                if sheet.printable:
                    export_sheets.Add(sheet.revit_sheet.Id)
                    printable_sheets.append(sheet)

            if not printable_sheets:
                forms.alert('There are no printable sheets in the selection.')
                return

            combined_name = self._get_combined_pdf_name()
            combined_path = op.join(dirPath, combined_name)
            resolved_combined_path = self._resolve_output_filepath('Combined PDF', combined_path)
            if not resolved_combined_path:
                return

            combined_name = op.basename(resolved_combined_path)
            combined_path = resolved_combined_path

            try:
                with revit.Transaction('Reload Keynote File',
                                       doc=self.selected_doc):
                    DB.KeynoteTable.GetKeynoteTable(self.selected_doc).Reload(None)

                optspdf = PrintUtils.pdf_opts(doc=self.selected_doc)
                optspdf.Combine = True
                optspdf.FileName = op.splitext(combined_name)[0]
                self.selected_doc.Export(dirPath, export_sheets, optspdf)
            except Exception as export_err:
                forms.alert(
                    'Failed to export combined PDF with Revit Export to PDF.',
                    expanded=str(export_err)
                    )
            return

        # Legacy print path for Revit 2021 and earlier
        # make sure we can access the print config
        print_mgr = self._get_printmanager()
        if not print_mgr:
            forms.alert(
                "Error getting print manager for this document",
                exitscript=True
                )
        with revit.TransactionGroup('Print Sheets in Order',
                                    doc=self.selected_doc):
            with revit.Transaction('Set Printer Settings',
                                   doc=self.selected_doc,
                                   log_errors=False):
                try:
                    print_mgr.PrintSetup.CurrentPrintSetting =                         self.selected_print_setting.print_settings
                    print_mgr.SelectNewPrintDriver(self.selected_printer)
                    print_mgr.PrintRange = DB.PrintRange.Select
                except Exception as cpSetEx:
                    forms.alert(
                        "Print setting is incompatible with printer.",
                        expanded=str(cpSetEx)
                        )
                    return
            # The OrderedViewList property was added to the IViewSheetSet
            # interface in Revit 2023 and makes the non-printable char
            # technique unnecessary.
            supports_OrderedViewList = HOST_APP.is_newer_than(2022)
            if supports_OrderedViewList:
                sheet_list = List[DB.View]()
                for sheet in target_sheets:
                    if sheet.printable:
                        sheet_list.Add(sheet.revit_sheet)
            else:
                # add non-printable char in front of sheet Numbers
                # to push revit to sort them per user
                sheet_set = DB.ViewSet()
                original_sheetnums = []
                with revit.Transaction('Fix Sheet Numbers',
                                       doc=self.selected_doc):
                    for idx, sheet in enumerate(target_sheets):
                        rvtsheet = sheet.revit_sheet
                        # removing any NPC from previous failed prints
                        if NPC in rvtsheet.SheetNumber:
                            rvtsheet.SheetNumber =                                 rvtsheet.SheetNumber.replace(NPC, '')
                        # create a list of the existing sheet numbers
                        original_sheetnums.append(rvtsheet.SheetNumber)
                        # add a prefix (NPC) for sorting purposes
                        rvtsheet.SheetNumber =                             NPC * (idx + 1) + rvtsheet.SheetNumber
                        if sheet.printable:
                            sheet_set.Insert(rvtsheet)

            # Collect existing sheet sets
            cl = DB.FilteredElementCollector(self.selected_doc)
            viewsheetsets = cl.OfClass(framework.get_type(DB.ViewSheetSet))                              .WhereElementIsNotElementType()                              .ToElements()
            all_viewsheetsets = {vss.Name: vss for vss in viewsheetsets}

            sheetsetname = 'OrderedPrintSet'

            with revit.Transaction('Remove Previous Print Set',
                                   doc=self.selected_doc):
                # Delete existing matching sheet set
                if sheetsetname in all_viewsheetsets:
                    print_mgr.ViewSheetSetting.CurrentViewSheetSet =                         all_viewsheetsets[sheetsetname]
                    print_mgr.ViewSheetSetting.Delete()

            with revit.Transaction('Update Ordered Print Set',
                                   doc=self.selected_doc):
                try:
                    viewsheet_settings = print_mgr.ViewSheetSetting
                    if supports_OrderedViewList:
                        viewsheet_settings.CurrentViewSheetSet.IsAutomatic = False
                        viewsheet_settings.CurrentViewSheetSet.OrderedViewList =                             sheet_list
                    else:
                        viewsheet_settings.CurrentViewSheetSet.Views =                             sheet_set
                    viewsheet_settings.SaveAs(sheetsetname)
                except Exception as viewset_err:
                    sheet_report = ''
                    for sheet in sheet_set:
                        sheet_report += '{} {}\n'.format(
                            sheet.SheetNumber if isinstance(sheet,
                                                            DB.ViewSheet)
                            else '---',
                            type(sheet)
                            )
                    logger.critical(
                        'Error setting sheet set on print mechanism. '
                        'These items are included in the viewset '
                        'object:\n%s', sheet_report
                        )
                    raise viewset_err

            # set print job configurations
            print_mgr.PrintOrderReverse = self.reverse_print
            try:
                print_mgr.CombinedFile = True
            except Exception as e:
                forms.alert(str(e) +
                            '\nSet printer correctly in Print settings.')
                script.exit()
            print_filepath = op.join('C:', 'Ordered Sheet Set.pdf')
            print_mgr.PrintToFile = True
            print_mgr.PrintToFileName = print_filepath

            with revit.Transaction('Reload Keynote File',
                                   doc=self.selected_doc):
                DB.KeynoteTable.GetKeynoteTable(revit.doc).Reload(None)
            print_mgr.Apply()
            print_mgr.SubmitPrint()
            if not supports_OrderedViewList:
                # now fix the sheet names
                with revit.Transaction('Restore Sheet Numbers', doc=self.selected_doc):
                    for sheet, sheetnum in zip(target_sheets, original_sheetnums):
                        rvtsheet = sheet.revit_sheet
                        rvtsheet.SheetNumber = sheetnum

            self._reset_psettings()

    # -------------------------------------------------------------------------
    # Print Execution — Individual Sheets
    # -------------------------------------------------------------------------

    def _print_sheets_in_order(self, target_sheets):
        # make sure we can access the print config
        print_mgr = self._get_printmanager()
        print_mgr.PrintToFile = True
        per_sheet_psettings = self.selected_print_setting.allows_variable_paper

        # make sure you can print, construct print path and make directory
        dirPath = self._get_output_dir()
        doc = self.selected_doc

        if self.selected_printer == "Revit Internal Printer" or self.export_extra_enabled:
            PrintUtils.open_dir(dirPath)
        with revit.Transaction('Reload Keynote File',
                               doc=self.selected_doc):
            DB.KeynoteTable.GetKeynoteTable(self.selected_doc).Reload(None)

        with revit.DryTransaction('Set Printer Settings',
                                  doc=self.selected_doc):
            try:
                if not per_sheet_psettings:
                    print_mgr.PrintSetup.CurrentPrintSetting = \
                        self.selected_print_setting.print_settings
                if not (IS_REVIT_2022_OR_NEWER and self.selected_printer == "Revit Internal Printer"):
                    print_mgr.SelectNewPrintDriver(self.selected_printer)
                print_mgr.PrintRange = DB.PrintRange.Current
            except Exception as cpSetEx:
                forms.alert(
                    "Print setting is incompatible with printer.",
                    expanded=str(cpSetEx)
                    )
                return
            if target_sheets:
                if self.export_extra_enabled:
                    export_label = self._get_extra_export_label()
                    with forms.ProgressBar(step=1, title='Exporting PDF & {}s... '.format(export_label) + '{value} of {max_value}', cancellable=True) as pb1:
                        pbTotal1 = len(target_sheets) * 2
                        pbCount1 = 1
                        for sheet in target_sheets:
                            if pb1.cancelled:
                                break
                            else:
                                if sheet.printable:
                                    if sheet.print_filename:
                                        effective_filename = self._apply_filename_prefix(sheet.print_filename)
                                        print_filepath = op.join(dirPath, effective_filename)
                                        print_mgr.PrintToFileName = print_filepath

                                        # set the per-sheet print settings if required
                                        if per_sheet_psettings:
                                            print_mgr.PrintSetup.CurrentPrintSetting = \
                                                sheet.print_settings

                                        resolved_filepath = self._resolve_output_filepath(sheet.name, print_filepath)
                                        if resolved_filepath:
                                            print_filepath = resolved_filepath
                                            effective_filename = op.basename(print_filepath)
                                            print_mgr.PrintToFileName = print_filepath
                                            try:
                                                pb1.update_progress(pbCount1, pbTotal1)
                                                pbCount1 += 1
                                                if IS_REVIT_2022_OR_NEWER and self.selected_printer == "Revit Internal Printer":
                                                    optspdf = PrintUtils.pdf_opts(doc=doc)
                                                    PrintUtils.export_sheet_pdf(dirPath, sheet.revit_sheet, optspdf, doc, effective_filename)
                                                else:
                                                    print_mgr.SubmitPrint(sheet.revit_sheet)
                                            except Exception as e:
                                                logger.error('Failed to export PDF for sheet %s: %s', sheet.number, e)

                                            try:
                                                pb1.update_progress(pbCount1, pbTotal1)
                                                pbCount1 += 1
                                                optsdwg = PrintUtils.dwg_opts()
                                                PrintUtils.export_sheet_cad(dirPath, sheet.revit_sheet, optsdwg, doc, effective_filename, self._get_extra_export_extension())
                                            except Exception as e:
                                                logger.error('Failed to export %s for sheet %s: %s', self._get_extra_export_label(), sheet.number, e)
                                    else:
                                        pbCount1 += 2
                                        logger.debug(
                                            'Sheet %s does not have a valid file name.',
                                            sheet.number)
                                else:
                                    pbCount1 += 2
                                    logger.debug('Sheet %s is not printable. Skipping print.', sheet.number)
                else:
                    with forms.ProgressBar(step=1, title='Exporting PDFs... ' + '{value} of {max_value}', cancellable=True) as pb1:
                        pbTotal1 = len(target_sheets)
                        pbCount1 = 1
                        for sheet in target_sheets:
                            if pb1.cancelled:
                                break
                            else:
                                if sheet.printable:
                                    if sheet.print_filename:
                                        effective_filename = self._apply_filename_prefix(sheet.print_filename)
                                        print_filepath = op.join(dirPath, effective_filename)

                                        print_mgr.PrintToFileName = print_filepath

                                        if per_sheet_psettings:
                                            print_mgr.PrintSetup.CurrentPrintSetting = \
                                                sheet.print_settings

                                        resolved_filepath = self._resolve_output_filepath(sheet.name, print_filepath)
                                        if resolved_filepath:
                                            print_filepath = resolved_filepath
                                            effective_filename = op.basename(print_filepath)
                                            print_mgr.PrintToFileName = print_filepath
                                            try:
                                                pb1.update_progress(pbCount1, pbTotal1)
                                                pbCount1 += 1
                                                if IS_REVIT_2022_OR_NEWER and self.selected_printer == "Revit Internal Printer":
                                                    optspdf = PrintUtils.pdf_opts(doc=doc)
                                                    PrintUtils.export_sheet_pdf(dirPath, sheet.revit_sheet, optspdf, doc, effective_filename)
                                                else:
                                                    print_mgr.SubmitPrint(sheet.revit_sheet)
                                            except Exception as e:
                                                logger.error('Failed to export PDF for sheet %s: %s', sheet.number, e)

                                    else:
                                        pbCount1 += 1
                                        logger.debug(
                                            'Sheet %s does not have a valid file name.',
                                            sheet.number)
                                else:
                                    pbCount1 += 1
                                    logger.debug('Sheet %s is not printable. Skipping print.', sheet.number)

    # -------------------------------------------------------------------------
    # Print Execution — Linked Model Sheets
    # -------------------------------------------------------------------------

    def _print_linked_sheets_in_order(self, target_sheets, target_doc):
        # make sure we can access the print config
        print_mgr = self._get_printmanager()
        print_mgr.PrintToFile = True
        if not (IS_REVIT_2022_OR_NEWER and self.selected_printer == "Revit Internal Printer"):
            print_mgr.SelectNewPrintDriver(self.selected_printer)
            print_mgr.PrintRange = DB.PrintRange.Current

        dirPath = self._get_output_dir()
        doc = target_doc

        if self.selected_printer == "Revit Internal Printer":
            PrintUtils.open_dir(dirPath)

        if target_sheets:
            with forms.ProgressBar(step=1, title='Exporting Linked PDFs... ' + '{value} of {max_value}', cancellable=True) as pb1:
                pbTotal1 = len(target_sheets)
                pbCount1 = 1
                for sheet in target_sheets:
                    if pb1.cancelled:
                        break
                    else:
                        if sheet.printable:
                            if sheet.print_filename:
                                effective_filename = self._apply_filename_prefix(sheet.print_filename)
                                print_filepath = op.join(dirPath, effective_filename)
                                print_mgr.PrintToFileName = print_filepath

                                resolved_filepath = self._resolve_output_filepath(sheet.name, print_filepath)
                                if resolved_filepath:
                                    print_filepath = resolved_filepath
                                    effective_filename = op.basename(print_filepath)
                                    print_mgr.PrintToFileName = print_filepath
                                    try:
                                        pb1.update_progress(pbCount1, pbTotal1)
                                        pbCount1 += 1
                                        if IS_REVIT_2022_OR_NEWER and self.selected_printer == "Revit Internal Printer":
                                            optspdf = PrintUtils.pdf_opts(doc=doc)
                                            PrintUtils.export_sheet_pdf(dirPath, sheet.revit_sheet, optspdf, doc, effective_filename)
                                        else:
                                            print_mgr.SubmitPrint(sheet.revit_sheet)
                                    except Exception as e:
                                        logger.error('Failed to export PDF for sheet %s: %s', sheet.number, e)
                            else:
                                pbCount1 += 1
                                logger.debug(
                                    'Sheet %s does not have a valid file name.',
                                    sheet.number)
                        else:
                            pbCount1 += 1
                            logger.debug('Sheet %s is not printable. Skipping print.', sheet.number)

    # -------------------------------------------------------------------------
    # Event Handlers
    # -------------------------------------------------------------------------

    def doclist_changed(self, sender, args):
        if self._suspend_ui_events:
            return
        self.project_info = revit.query.get_project_info(doc=self.selected_doc)
        self._apply_suggested_filename()
        self._load_selected_doc_context()

    def sheetlist_changed(self, sender, args):
        if self._suspend_ui_events:
            return
        print_settings = None
        tblocks = self._get_doc_titleblocks(self.selected_doc)
        if self.selected_sheetlist and self.has_print_settings:
            rev_cfg = DB.RevisionSettings.GetRevisionSettings(revit.doc)
            if self.selected_print_setting.allows_variable_paper:
                sheet_printsettings = \
                    self._get_sheet_printsettings(
                        tblocks,
                        self._get_doc_print_settings(self.selected_doc)
                        )
                self.show_element(self.varsizeguide)
                self.show_element(self.psettingcol)
                self._scheduled_sheets = [
                    ViewSheetListItem(
                        view_sheet=x,
                        view_tblock=self._find_sheet_tblock(x, tblocks),
                        print_settings=sheet_printsettings.get(
                            x.SheetNumber,
                            None),
                        rev_settings=rev_cfg)
                    for x in self.selected_sheetlist.get_sheets(
                        doc=self.selected_doc
                        )
                    ]
            else:
                print_settings = self.selected_print_setting.print_settings
                self.hide_element(self.varsizeguide)
                self.hide_element(self.psettingcol)
                self._scheduled_sheets = [
                    ViewSheetListItem(
                        view_sheet=x,
                        view_tblock=self._find_sheet_tblock(x, tblocks),
                        print_settings=TitleBlockPrintSettings(
                            psettings=[print_settings],
                            set_by_param=False
                        ),
                        rev_settings=rev_cfg)
                    for x in self.selected_sheetlist.get_sheets(
                        doc=self.selected_doc
                        )
                    ]
        self._original_sheet_order = list(self._scheduled_sheets)
        self._update_combine_option()
        # self._update_index_slider()
        self.options_changed(None, None)

    def printers_changed(self, sender, args):
        if self._suspend_ui_events:
            return
        print_mgr = self._get_printmanager()

        self._update_output_mode_ui()
        if self.selected_printer == "Revit Internal Printer":
            return
        print_mgr.SelectNewPrintDriver(self.selected_printer)
        self._setup_print_settings()

    def options_changed(self, sender, args):
        if self._suspend_ui_events:
            return
        self._reset_error()
        self._update_output_mode_ui()

        # update index digit range
        self._update_index_slider()

        # reverse sheet if reverse is set
        base_sheet_list = list(self._original_sheet_order) if self._original_sheet_order else [x for x in self._scheduled_sheets]
        sheet_list = self._sort_sheet_list(base_sheet_list)
        if self.reverse_print:
            sheet_list.reverse()

        if self.combine_cb.IsChecked:
            self.hide_element(self.order_sp)
            self.hide_element(self.namingformat_dp)
            self.hide_element(self.pfilename)
            self.exportformat_cb.SelectedIndex = 0
            self.exportformat_cb.IsEnabled = False
        else:
            self.show_element(self.order_sp)
            self.show_element(self.namingformat_dp)
            self.show_element(self.pfilename)
            self.exportformat_cb.IsEnabled = True

        if self.selected_doc.IsLinked:
            self.exportformat_cb.SelectedIndex = 0
            self.exportformat_cb.IsEnabled = False

        # decide whether to show the placeholders or not
        self._refresh_sheet_order_ui(sheet_list)

    def set_sheet_printsettings(self, sender, args):
        if self.selected_printable_sheets:
            # make sure none of the sheets has readonly print setting
            if any(x.read_only for x in self.selected_printable_sheets):
                forms.alert("Print settings has been set by titleblock "
                            "for one or more sheets and can only be changed "
                            "by modifying the titleblock print setting")
                return

            all_psettings = \
                [x for x in self.print_settings if x.is_user_defined]
            sheet_psettings = \
                self.selected_printable_sheets[0].all_print_settings
            if sheet_psettings:
                options = {
                    'Matching Print Settings':
                        self._get_psetting_items(
                            doc=self.selected_doc,
                            psettings=sheet_psettings
                            ),
                    'All Print Settings':
                        all_psettings
                }
            else:
                options = all_psettings or []

            if options:
                psetting_item = forms.SelectFromList.show(
                    options,
                    name_attr='name',
                    group_selector_title='Print Settings:',
                    default_group='Matching Print Settings',
                    title='Select Print Setting',
                    item_container_template=self.Resources["printSettingsItem"],
                    width=450, height=400
                    )
                if psetting_item:
                    for sheet in self.selected_printable_sheets:
                        sheet.print_settings = psetting_item
            else:
                forms.alert('There are no print settings in this model.')

    def sheet_selection_changed(self, sender, args):
        if self.selected_printable_sheets:
            return self.enable_element(self.sheetopts_wp)
        self.disable_element(self.sheetopts_wp)

    def validate_index_start(self, sender, args):
        args.Handled = re.match(r'[^0-9]+', args.Text)

    def rest_index(self, sender, args):
        self.indexstart_tb.Text = '0'

    def edit_formats(self, sender, args):
        editfmt_wnd = \
            EditNamingFormatsWindow(
                'EditNamingFormats.xaml',
                start_with=self.selected_naming_format
                )
        editfmt_wnd.show_dialog()
        self.namingformat_cb.ItemsSource = editfmt_wnd.naming_formats
        self.namingformat_cb.SelectedItem = editfmt_wnd.selected_naming_format

    def edit_pdf_options(self, sender, args):
        try:
            xaml_path = op.join(op.dirname(__file__), 'PdfExportOptions.xaml')
            pdfopts_wnd = PdfExportOptionsWindow(xaml_path)
            if pdfopts_wnd.show_dialog():
                self._update_output_mode_ui()
        except Exception as pdf_options_err:
            logger.error('Failed to open PDF options window: %s', pdf_options_err)
            forms.alert(
                'Could not open PDF Options window.',
                expanded=str(pdf_options_err)
            )

    def copy_filenames(self, sender, args):
        if self.selected_sheets:
            filenames = [x.print_filename for x in self.selected_sheets]
            script.clipboard_copy('\n'.join(filenames))

    def print_sheets(self, sender, args):
        if self.sheet_list:
            selected_only = False
            if self.selected_sheets:
                opts = forms.alert(
                    "You have a series of sheets selected. Do you want to "
                    "print the selected sheets or all sheets?",
                    options=["Only Selected Sheets", "All Scheduled Sheets"]
                    )
                selected_only = opts == "Only Selected Sheets"

            target_sheets = \
                self.selected_sheets if selected_only else self.sheet_list

            if not self.combine_print:
                # verify all sheets have print settings
                if (self.selected_print_setting.allows_variable_paper
                        and not all(x.print_settings for x in target_sheets)):
                    forms.alert(
                        'Not all sheets have a print setting assigned to them. '
                        'Select sheets and assign print settings.')
                    return
                # confirm print if a lot of sheets are going to be printed
                printable_count = len([x for x in target_sheets if x.printable])
                if printable_count > 5:
                    # prepare warning message
                    sheet_count = len(target_sheets)
                    message = str(printable_count)
                    if printable_count != sheet_count:
                        message += ' (out of {} total)'.format(sheet_count)

                    if not forms.alert('Are you sure you want to print {} '
                                       'sheets individually? The process can '
                                       'not be cancelled.'.format(message),
                                       ok=False, yes=True, no=True):
                        return
            # close window and submit print
            self.Close()
            if self.combine_print:
                self._print_combined_sheets_in_order(target_sheets)
            else:
                if self.selected_doc.IsLinked:
                    self._print_linked_sheets_in_order(target_sheets, self.selected_doc)
                else:
                    self._print_sheets_in_order(target_sheets)


# =============================================================================
# SHIFT-CLICK UTILITY — CLEANUP NON-PRINTABLE CHARACTERS
# =============================================================================

def cleanup_sheetnumbers(doc):
    sheets = revit.query.get_sheets(doc=doc)
    with revit.Transaction('Cleanup Sheet Numbers', doc=doc):
        for sheet in sheets:
            sheet.SheetNumber = sheet.SheetNumber.replace(NPC, '')


# =============================================================================
# ENTRY POINT
# =============================================================================

# verify model is printable
forms.check_modeldoc(exitscript=True)
# ensure there is nothing selected
revit.selection.get_selection().clear()

# TODO: add copy filenames to sheet list
if __shiftclick__:  # pylint: disable=E0602
    open_docs = forms.select_open_docs(check_more_than_one=False)
    if open_docs:
        for open_doc in open_docs:
            cleanup_sheetnumbers(open_doc)
else:
    PrintSheetsWindow('PrintSheets.xaml').ShowDialog()

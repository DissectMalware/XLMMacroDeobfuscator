from XLMMacroDeobfuscator.excel_wrapper import XlApplicationInternational, RowAttribute
from zipfile import ZipFile
from glob import fnmatch
try:
    from defusedxml import ElementTree
except:
    from xml.etree import ElementTree
    print('XLMMacroDeobfuscator: defusedxml is not installed (required to securely parse XLSM files)')
from XLMMacroDeobfuscator.excel_wrapper import ExcelWrapper
from XLMMacroDeobfuscator.boundsheet import *
import untangle
from io import StringIO
import os


class XLSMWrapper(ExcelWrapper):
    def __init__(self, xlsm_doc_path):
        self.xlsm_doc_path = xlsm_doc_path
        self.xlsm_workbook_name = os.path.basename(xlsm_doc_path)
        self._content_types = None
        self._style = None
        self._theme = None
        self._types = None
        self._workbook = None
        self._workbook_rels = None
        self._workbook_relationships = None
        self._workbook_style = None
        self._defined_names = None
        self._macrosheets = None
        self._worksheets = None
        self._shared_strings = None
        self.xl_international_flags = {XlApplicationInternational.xlLeftBracket: '[',
                                       XlApplicationInternational.xlListSeparator: ',',
                                       XlApplicationInternational.xlRightBracket: ']'}

        self._types = self._get_types()
        self.color_maps = None

    def get_workbook_name(self):
        return self.xlsm_workbook_name

    def _get_types(self):
        result = {}
        if self._types is None:
            main = self.get_content_types()
            if hasattr(main, 'Types'):
                if hasattr(main.Types, 'Override'):
                    for i in main.Types.Override:
                        result[i.get_attribute('ContentType')] = i.get_attribute('PartName')

                if hasattr(main.Types, 'Default'):
                    for i in main.Types.Default:
                        result[i.get_attribute('ContentType')] = i.get_attribute('Extension')
        else:
            result = self._types

        return result

    def _get_relationships(self):
        result = {}
        if self._workbook_relationships is None:
            main = self._get_workbook_rels()
            if hasattr(main, 'Relationships'):
                if hasattr(main.Relationships, 'Relationship'):
                    for i in main.Relationships.Relationship:
                        result[i.get_attribute('Id')] = i
            self._workbook_relationships = result
        else:
            result = self._workbook_relationships

        return result

    def get_xl_international_char(self, flag_name):
        result = None
        if flag_name in self.xl_international_flags:
            result = self.xl_international_flags[flag_name]

        return result

    def get_files(self, file_name_filters=None):
        input_zip = ZipFile(self.xlsm_doc_path)
        result = {}
        if not file_name_filters:
            file_name_filters = ['*']
        for i in input_zip.namelist():
            for filter in file_name_filters:
                if i == filter or fnmatch.fnmatch(i, filter):
                    result[i] = input_zip.read(i)
        #Excel Crack is Wack... Excel converts \x5c to \x2f in zip file names and will happily eat it for you. Sample 51762ea84ac51f9e40b1902ebe22c306a732d77a5aa8f03650279d8b21271516
        if not result:
            for i in input_zip.namelist():
                for filter in file_name_filters:            
                    if i == filter.replace('\x2f','\x5c') or fnmatch.fnmatch(i, filter.replace('\x2f','\x5c')):
                        result[i.replace('\x5c','\x2f')] = input_zip.read(i)
        return result

    def get_xml_file(self, file_name, ignore_pattern=None):
        if file_name.startswith('/'):
            file_name = file_name[1:]
        result = None
        file_name = file_name
        files = self.get_files([file_name])
        if len(files) == 1:
            workbook_content = files[file_name].decode('utf_8')
            if ignore_pattern:
                workbook_content = re.sub(ignore_pattern, "", workbook_content)
            result = untangle.parse(StringIO(workbook_content))
        return result

    def get_content_types(self):
        if not self._content_types:
            content_type = self.get_xml_file('[Content_Types].xml')
            self._content_types = content_type
        return self._content_types

    def _get_workbook_path(self):
        workbook_path = 'xl/workbook.xml'
        if 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' in self._types:
            workbook_path = self._types[
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml']
        elif 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' in self._types:
            workbook_path = self._types['application/vnd.ms-excel.sheet.macroEnabled.main+xml']
        workbook_path = workbook_path.lstrip('/')

        path = ''
        name = workbook_path

        if '/' in workbook_path:
            path = workbook_path[:workbook_path.index('/')]
            name = workbook_path[workbook_path.index('/') + 1:]
        return workbook_path, path, name

    def get_workbook(self):
        if not self._workbook:
            workbook_path, _, _ = self._get_workbook_path()
            workbook = self.get_xml_file(workbook_path)
            self._workbook = workbook
        return self._workbook

    def get_style(self):
        if not self._style:
            types = self._get_types()
            rel_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
            if rel_type in types:
                style = self.get_xml_file(types[rel_type])
                self._style = style

        return self._style

    def get_theme(self):
        if not self._theme:
            types = self._get_types()
            rel_type = "application/vnd.openxmlformats-officedocument.theme+xml"
            if rel_type in types:
                style = self.get_xml_file(types[rel_type])
                self._theme = style

        return self._theme

    def get_workbook_style(self):
        if not self._workbook_style:
            style_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
            relationships = self._get_relationships()
            if style_type in relationships:
                style_sheet_path = relationships[style_type]
                _, base_dir, _ = self._get_workbook_path()
                style_sheet = self.get_xml_file(base_dir+'/'+style_sheet_path)
            self._workbook_style = style_sheet

        return self._workbook_style

    def _get_workbook_rels(self):
        if not self._workbook_rels:

            type = 'rels'
            if 'application/vnd.openxmlformats-package.relationships+xml' in self._types:
                type = self._types['application/vnd.openxmlformats-package.relationships+xml']

            workbook_path, base_dir, name = self._get_workbook_path()

            path = '{}/_{}/{}.{}'.format(base_dir, type, name, type)
            workbook = self.get_xml_file(path)
            self._workbook_rels = workbook
        return self._workbook_rels

    def get_sheet_info(self, rId):
        sheet_type = None
        sheet_path = None
        nsmap = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        workbook_rels = self._get_workbook_rels()

        relationships = self._get_relationships()

        if rId in relationships:
            sheet_path = relationships[rId].get_attribute('Target')
            type = relationships[rId].get_attribute('Type')
            if type == "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" or \
                    type == 'http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet':
                sheet_type = 'Macrosheet'
            elif type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
                sheet_type = 'Worksheet'
            else:
                sheet_type = 'Unknown'

        return sheet_type, sheet_path

    def get_defined_names(self):
        if self._defined_names is None:
            workbook_obj = self.get_workbook()
            self._defined_names = {}
            if hasattr(workbook_obj, 'workbook') \
                    and hasattr(workbook_obj.workbook, 'definedNames')\
                    and hasattr(workbook_obj.workbook.definedNames, 'definedName'):
                for defined_name in workbook_obj.workbook.definedNames.definedName:
                    self._defined_names[defined_name.get_attribute(
                        'name').replace('_xlnm.', '').lower()] = defined_name.cdata

        return self._defined_names

    def get_sheet_infos(self, types):
        result = []
        workbook_obj = self.get_workbook()
        sheet_names = set()

        _, base_dir, _ = self._get_workbook_path()
        if hasattr(workbook_obj, 'workbook'):
            for sheet_elm in workbook_obj.workbook.sheets.sheet:
                rId = sheet_elm.get_attribute('r:id')
                name = sheet_elm.get_attribute('name')
                sheet_type, rel_path = self.get_sheet_info(rId)
                if rel_path is not None:
                    path = base_dir + '/' + rel_path
                    if sheet_type in types and name not in sheet_names:
                        sheet = Boundsheet(name, sheet_type)
                        result.append({'sheet': sheet,
                                       'sheet_path': path,
                                       'sheet_xml': self.get_xml_file(path, ignore_pattern="<c[^>]+/>")})
                        sheet_names.add(name)
                else:
                    print("Sheet('{}') does not have a valid rId('{}')".format(name, rId))

        return result

    def get_macrosheet_infos(self):
        return self.get_sheet_infos(['Macrosheet'])

    def get_worksheet_infos(self):
        return self.get_sheet_infos(['Worksheet'])

    def get_shared_strings(self):
        if self._shared_strings is None:
            _, base_dir, _ = self._get_workbook_path()
            content = self.get_xml_file(base_dir + '/sharedStrings.xml')
            if content is not None:
                if hasattr(content, 'sst') and hasattr(content.sst, 'si'):
                    for str in content.sst.si:
                        if self._shared_strings is None:
                            self._shared_strings = []
                        if hasattr(str, 't'):
                            self._shared_strings.append(str.t.cdata)
                        elif hasattr(str, 'r') and len(str.r) > 0 and hasattr(str.r[0], 't'):
                            self._shared_strings.append(str.r[0].t.cdata)

        return self._shared_strings

    def load_macro_cells(self, macrosheet, macrosheet_obj, macrosheet_names):
        strings = self.get_shared_strings()

        sheet = macrosheet_obj.xm_macrosheet if hasattr(macrosheet_obj, 'xm_macrosheet') else macrosheet_obj.worksheet
        if not hasattr(sheet.sheetData, 'row'):
            return
        for row in sheet.sheetData.row:
            row_attribs = {}
            for attr in row._attributes:
                if attr == 'ht':
                    row_attribs[RowAttribute.Height] = row.get_attribute('ht')
                elif attr == 'spans':
                    row_attribs[RowAttribute.Spans] = row.get_attribute('spans')
            if len(row_attribs) > 0:
                macrosheet.row_attributes[row.get_attribute('r')] = row_attribs
            if hasattr(row, 'c'):
                for cell_elm in row.c:
                    formula_text = None
                    if hasattr(cell_elm, 'f'):
                        formula = cell_elm.f
                        if formula.get_attribute('bx') == "1":
                            text = formula.cdata
                            formula_text = None
                            if text:
                                first_eq_sign = text.find('=')
                                if first_eq_sign > 0:
                                    formula_text = '=SET.NAME("{}",{})'.format(text[:first_eq_sign], text[first_eq_sign+1:])
                        else:
                            formula_text = ('=' + formula.cdata) if formula is not None else None
                            
                    if formula_text:
                        for name in macrosheet_names:
                            if name.lower() + '!' in formula_text.lower():
                                formula_text = re.sub('{}\!'.format(name), "'{}'!".format(name), formula_text)
                    value_text = None
                    is_string = False
                    if 't' in cell_elm._attributes and cell_elm.get_attribute('t') == 's':
                        is_string = True

                    cached_str = False
                    if 't' in cell_elm._attributes and cell_elm.get_attribute('t') == 'str':
                        cached_str = True

                    if hasattr(cell_elm, 'v'):
                        value = cell_elm.v
                        value_text = value.cdata if value is not None else None
                        if value_text is not None and is_string:
                            value_text = strings[int(value_text)]
                    location = cell_elm.get_attribute('r')
                    if formula_text or value_text:
                        cell = Cell()
                        sheet_name, cell.column, cell.row = Cell.parse_cell_addr(location)
                        cell.sheet = macrosheet
                        if not cached_str:
                            cell.formula = formula_text
                        cell.value = value_text
                        macrosheet.cells[location] = cell

                        for attrib in cell_elm._attributes:
                            if attrib != 'r':
                                cell.attributes[attrib] = cell_elm._attributes[attrib]

    def load_worksheet_cells(self, macrosheet, macrosheet_obj):
        strings = self.get_shared_strings()
        if not hasattr(macrosheet_obj.worksheet.sheetData, 'row'):
            return
        for row in macrosheet_obj.worksheet.sheetData.row:
            row_attribs = {}
            for attr in row._attributes:
                if attr == 'ht':
                    row_attribs[RowAttribute.Height] = row.get_attribute('ht')
                elif attr == 'spans':
                    row_attribs[RowAttribute.Spans] = row.get_attribute('spans')
            if len(row_attribs) > 0:
                macrosheet.row_attributes[row.get_attribute('r')] = row_attribs
            if hasattr(row, 'c'):
                for cell_elm in row.c:
                    formula_text = None
                    if hasattr(cell_elm, 'f'):
                        formula = cell_elm.f
                        formula_text = ('=' + formula.cdata) if formula is not None else None
                    value_text = None
                    is_string = False
                    if 't' in cell_elm._attributes and cell_elm.get_attribute('t') == 's':
                        is_string = True

                    if hasattr(cell_elm, 'v'):
                        value = cell_elm.v
                        value_text = value.cdata if value is not None else None
                        if value_text is not None and is_string:
                            value_text = strings[int(value_text)]
                    location = cell_elm.get_attribute('r')
                    cell = Cell()
                    sheet_name, cell.column, cell.row = Cell.parse_cell_addr(location)
                    cell.sheet = macrosheet
                    cell.formula = formula_text
                    cell.value = value_text
                    macrosheet.cells[location] = cell

                    for attrib in cell_elm._attributes:
                        if attrib != 'r':
                            cell.attributes[attrib] = cell_elm._attributes[attrib]

    def get_defined_name(self, name, full_match=True):
        result = []
        name = name.lower()

        if full_match:
            if name in self.get_defined_names():
                result = self._defined_names[name]
        else:
            for defined_name, cell_address in self.get_defined_names().items():
                if defined_name.startswith(name):
                    result.append((defined_name, cell_address))

        return result

    def get_macrosheets(self):
        if self._macrosheets is None:
            self._macrosheets = {}
            macrosheets = self.get_macrosheet_infos()

            macrosheet_names = []
            for macrosheet in macrosheets:
                macrosheet_names.append(macrosheet['sheet'].name)

            for macrosheet in macrosheets:
                # if the actual file exist
                if macrosheet['sheet_xml']:
                    self.load_macro_cells(macrosheet['sheet'], macrosheet['sheet_xml'], macrosheet_names)
                    sheet = macrosheet['sheet_xml'].xm_macrosheet if hasattr(macrosheet['sheet_xml'],
                                                                    'xm_macrosheet') else macrosheet['sheet_xml'].worksheet
                    if hasattr(sheet, 'sheetFormatPr'):
                        macrosheet['sheet'].default_height = sheet.sheetFormatPr.get_attribute(
                            'defaultRowHeight')

                self._macrosheets[macrosheet['sheet'].name] = macrosheet['sheet']

        return self._macrosheets

    def get_worksheets(self):
        if self._worksheets is None:
            self._worksheets = {}
            _worksheets = self.get_worksheet_infos()
            for worksheet in _worksheets:
                self.load_worksheet_cells(worksheet['sheet'], worksheet['sheet_xml'])
                if hasattr(worksheet['sheet_xml'].worksheet, 'sheetFormatPr'):
                    worksheet['sheet'].default_height = worksheet['sheet_xml'].worksheet.sheetFormatPr.get_attribute(
                        'defaultRowHeight')

                self._worksheets[worksheet['sheet'].name] = worksheet['sheet']

        return self._worksheets

    def get_color_index(self, rgba_str):

        r, g, b = int('0x' + rgba_str[2:4], base=16), int('0x' + rgba_str[4:6], base=16), int(
            '0x' + rgba_str[6:8], base=16)

        if self.color_maps is None:
            colors = [
                (0, 0, 0, 1), (255, 255, 255, 2), (255, 0, 0, 3), (0, 255, 0, 4),
                (0, 0, 255, 5), (255, 255, 0, 6), (255, 0, 255, 7), (0, 255, 255, 8),
                (128, 0, 0, 9), (0, 128, 0, 10), (0, 0, 128, 11), (128, 128, 0, 12),
                (128, 0, 128, 13), (0, 128, 128, 14), (192, 192, 192, 15), (128, 128, 128, 16),
                (153, 153, 255, 17), (153, 51, 102, 18), (255, 255, 204, 19), (204, 255, 255, 20),
                (102, 0, 102, 21), (255, 128, 128, 22), (0, 102, 204, 23), (204, 204, 255, 24),
                (0, 0, 128, 25), (255, 0, 255, 26), (255, 255, 0, 27), (0, 255, 255, 28),
                (128, 0, 128, 29), (128, 0, 0, 30), (0, 128, 128, 31), (0, 0, 255, 32),
                (0, 204, 255, 33), (204, 255, 255, 34), (204, 255, 204, 35), (255, 255, 153, 36),
                (153, 204, 255, 37), (255, 153, 204, 38), (204, 153, 255, 39), (255, 204, 153, 40),
                (51, 102, 255, 41), (51, 204, 204, 42), (153, 204, 0, 43), (255, 204, 0, 44),
                (255, 153, 0, 45), (255, 102, 0, 46), (102, 102, 153, 47), (150, 150, 150, 48),
                (0, 51, 102, 49), (51, 153, 102, 50), (0, 51, 0, 51), (51, 51, 0, 52),
                (153, 51, 0, 53), (153, 51, 102, 54), (51, 51, 153, 55), (51, 51, 51, 56)
            ]
            self.color_maps = {}

            for i in colors:
                c_r, c_g, c_b, index = i
                if (c_r, c_g, c_b) not in self.color_maps:
                    self.color_maps[(c_r, c_g, c_b)] = index

        color_index = None

        if (r, g, b) in self.color_maps:
            color_index = self.color_maps[(r, g, b)]

        return color_index

    def get_cell_info(self, sheet_name, col, row, info_type_id):
        data = None
        not_exist = True
        not_implemented = False

        sheet = self._macrosheets[sheet_name]
        cell_addr = col+str(row)
        if info_type_id == 17:
            style = self.get_style()
            if row in sheet.row_attributes and RowAttribute.Height in sheet.row_attributes[row]:
                not_exist = False
                data = sheet.row_attributes[row][RowAttribute.Height]
            elif sheet.default_height is not None:
                data = sheet.default_height
                NotImplemented = True
            data = round(float(data) * 4) / 4

        else:
            style = self.get_style()
            cell_format = None
            font = None
            not_exist = False

            if cell_addr in sheet.cells:
                cell = sheet.cells[cell_addr]
                if 's' in cell.attributes:
                    index = int(cell.attributes['s'])
                    cell_format = style.styleSheet.cellXfs.xf[index]
                    if 'fontId' in cell_format._attributes:
                        font_index = int(cell_format.get_attribute('fontId'))
                        font = style.styleSheet.fonts.font[font_index]
            else:
                for cell_style in style.styleSheet.cellStyles.cellStyle:
                    if cell_style.get_attribute('name') == 'Normal':
                        index = int(cell_style.get_attribute('xfId'))
                        if type(style.styleSheet.cellStyleXfs.xf) is list:
                            cell_format = style.styleSheet.cellStyleXfs.xf[index]
                        else:
                            cell_format = style.styleSheet.cellStyleXfs.xf
                        if 'fontId' in cell_format._attributes:
                            font_index = int(cell_format.get_attribute('fontId'))
                            font = style.styleSheet.fonts.font[font_index]
                        break
                NotImplemented = True

            if info_type_id == 8:
                h_align_map = {
                    'general': 1,
                    'left': 2,
                    'center': 3,
                    'right': 4,
                    'fill': 5,
                    'justify': 6,
                    'centerContinuous': 7,
                    'distributed': 8
                }

                if hasattr(cell_format, 'alignment'):
                    horizontal_alignment = cell_format.alignment.get_attribute('horizontal')
                    data = h_align_map[horizontal_alignment.lower()]

                else:
                    data = 1

            elif info_type_id == 19:
                if hasattr(font, 'sz'):
                    size = font.sz
                    data = float(size.get_attribute('val'))

            elif info_type_id == 24:
                if 'rgb' in font.color._attributes:
                    rgba_str = font.color.get_attribute('rgb')
                    data = self.get_color_index(rgba_str)
                else:
                    data = 1

            elif info_type_id == 38:
                # Font Background Color
                fill_id = int(cell_format.get_attribute('fillId'))
                fill = style.styleSheet.fills.fill[fill_id]
                if hasattr(fill.patternFill, 'fgColor'):
                    rgba_str = fill.patternFill.fgColor.get_attribute('rgb')
                    data = self.get_color_index(rgba_str)
                else:
                    data = 0

            elif info_type_id == 50:
                if hasattr(cell_format, 'alignment'):
                    vertical_alignment = cell_format.alignment.get_attribute('vertical')
                else:
                    vertical_alignment = 'bottom'  # default

                v_alignment = {
                    'top': 1,
                    'center': 2,
                    'bottom': 3,
                    'justify': 4,
                    'distributed': 5,
                }
                data = v_alignment[vertical_alignment.lower()]

            else:
                not_implemented = True

        # return None, None, True
        return data, not_exist, not_implemented


if __name__ == '__main__':

    path = r"tmp\xlsb\6644bcba091c3104aebc0eab93d4247a884028aad389803d71f26541df325cf8.xlsm"

    xlsm_doc = XLSMWrapper(path)
    macrosheets = xlsm_doc.get_macrosheets()

    auto_open_labels = xlsm_doc.get_defined_name('auto_open', full_match=False)
    for label in auto_open_labels:
        print('auto_open: {}->{}'.format(label[0], label[1]))

    for macrosheet_name in macrosheets:
        print('SHEET: {}\t{}'.format(macrosheets[macrosheet_name].name,
                                     macrosheets[macrosheet_name].type))
        for formula_loc, info in macrosheets[macrosheet_name].cells.items():
            if info.formula is not None:
                print('{}\t{}\t{}'.format(formula_loc, info.formula, info.value))

        for formula_loc, info in macrosheets[macrosheet_name].cells.items():
            if info.formula is None:
                print('{}\t{}\t{}'.format(formula_loc, info.formula, info.value))

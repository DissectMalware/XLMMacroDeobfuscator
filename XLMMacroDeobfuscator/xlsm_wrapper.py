from XLMMacroDeobfuscator.excel_wrapper import XlApplicationInternational
from zipfile import ZipFile
from glob import fnmatch
from xml.etree import ElementTree
from XLMMacroDeobfuscator.excel_wrapper import ExcelWrapper
from XLMMacroDeobfuscator.boundsheet import *
import untangle
import os


class XLSMWrapper(ExcelWrapper):
    def __init__(self, xlsm_doc_path):
        self.xlsm_doc_path = xlsm_doc_path
        self._content_types = None
        self._types = None
        self._workbook = None
        self._workbook_rels = None
        self._workbook_relationships = None
        self._workbook_style = None
        self._defined_names = None
        self._macrosheets = None
        self.xl_international_flags = {XlApplicationInternational.xlLeftBracket: '[',
                                       XlApplicationInternational.xlListSeparator: ',',
                                       XlApplicationInternational.xlRightBracket: ']'}

        self._types = self._get_types()


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

        return result

    def get_xml_file(self, file_name):
        result = None
        file_name = file_name
        files = self.get_files([file_name])
        if len(files) == 1:
            workbook_content = files[file_name].decode('utf_8')
            result = untangle.parse(workbook_content)
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

        path=''
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

    def get_workbook_style(self):
        if not self._workbook_style:
            style_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
            relationships = self._get_relationships()
            if style_type in relationships:
                style_sheet_path= relationships[style_type]
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
        result = {}
        workbook_obj = self.get_workbook()
        if hasattr(workbook_obj.workbook, 'definedNames'):
            for defined_name in workbook_obj.workbook.definedNames.definedName:
                result[defined_name.get_attribute('name').replace('_xlnm.', '').lower()] = defined_name.cdata
        return result

    def get_macrosheet_infos(self):
        result = []
        workbook_obj = self.get_workbook()
        sheet_names = set()

        _, base_dir, _ = self._get_workbook_path()

        for sheet_elm in workbook_obj.workbook.sheets.sheet:
            rId = sheet_elm.get_attribute('r:id')
            name = sheet_elm.get_attribute('name')
            sheet_type, rel_path = self.get_sheet_info(rId)
            if rel_path is not None:
                path = base_dir + '/' + rel_path
                if sheet_type == 'Macrosheet' and name not in sheet_names:
                    sheet = Boundsheet(name, sheet_type)
                    result.append({'sheet': sheet,
                                   'sheet_path': path,
                                   'sheet_xml': self.get_xml_file(path)})
                    sheet_names.add(name)
            else:
                print("Sheet('{}') does not have a valid rId('{}')".format(name, rId))

        return result

    def load_cells(self, macrosheet, macrosheet_obj):
        for row in macrosheet_obj.xm_macrosheet.sheetData.row:
            row_attribs = {}
            for attr in row._attributes:
                if attr == 'ht':
                    row_attribs[RowAttributes.Height] = row.get_attribute('ht')
                elif attr == 'spans':
                    row_attribs[RowAttributes.Spans] = row.get_attribute('spans')
            if len(row_attribs) > 0:
                macrosheet.row_attributes[row.get_attribute('r')] = row_attribs
            for cell_elm in row:
                formula = cell_elm.c.f
                formula_text = ('=' + formula.cdata) if formula is not None else None
                value = cell_elm.c.v
                value_text = value.cdata if value is not None else None
                location = cell_elm.c.get_attribute('r')
                cell = Cell()
                sheet_name, cell.column, cell.row = Cell.parse_cell_addr(location)
                cell.sheet = macrosheet
                cell.formula = formula_text
                cell.value = value_text
                macrosheet.cells[location] = cell

    def get_defined_name(self, name, full_match=True):
        result = []
        name = name.lower()
        if self._defined_names is None:
            self._defined_names = self.get_defined_names()

        if full_match:
            if name in self._defined_names:
                result = self._defined_names[name]
        else:
            for defined_name, cell_address in self._defined_names.items():
                if defined_name.startswith(name):
                    result.append((defined_name, cell_address))

        return result

    def get_macrosheets(self):
        if self._macrosheets is None:
            self._macrosheets = {}
            macrosheets = self.get_macrosheet_infos()
            for macrosheet in macrosheets:
                self.load_cells(macrosheet['sheet'], macrosheet['sheet_xml'])
                self._macrosheets[macrosheet['sheet'].name] = macrosheet['sheet']

        return self._macrosheets

    def get_cell_info(self, sheet_name, col, row, info_type_id):
        data = None
        not_exist = True
        not_implemented = False

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

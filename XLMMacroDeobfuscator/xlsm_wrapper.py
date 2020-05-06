from XLMMacroDeobfuscator.excel_wrapper import XlApplicationInternational
from zipfile import ZipFile
from glob import fnmatch
from xml.etree import ElementTree
from XLMMacroDeobfuscator.excel_wrapper import ExcelWrapper
from XLMMacroDeobfuscator.boundsheet import *


class XLSMWrapper(ExcelWrapper):
    def __init__(self, xlsm_doc_path):
        self.xlsm_doc_path = xlsm_doc_path
        self._workbook = None
        self._workbook_rels = None
        self._defined_names = None
        self._macrosheets = None
        self.xl_international_flags = {XlApplicationInternational.xlLeftBracket: '[',
                                       XlApplicationInternational.xlListSeparator: ',',
                                       XlApplicationInternational.xlRightBracket: ']'}

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
                if fnmatch.fnmatch(i, filter):
                    result[i] = input_zip.read(i)

        return result

    def get_xml_file(self, file_name):
        result = None
        files = self.get_files([file_name])
        if len(files) == 1:
            workbook_content = files[file_name].decode('utf_8')
            result = ElementTree.fromstring(workbook_content)
        return result

    def get_workbook(self):
        if not self._workbook:
            workbook = self.get_xml_file('xl/workbook.xml')
            self._workbook = workbook
        return self._workbook

    def get_workbook_rels(self):
        if not self._workbook_rels:
            workbook = self.get_xml_file('xl/_rels/workbook.xml.rels')
            self._workbook_rels = workbook
        return self._workbook_rels

    def get_sheet_info(self, rId):
        sheet_type = None
        sheet_path = None
        nsmap = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        workbook_rels = self.get_workbook_rels()
        relationships = workbook_rels.findall('.//r:Relationship', namespaces=nsmap)

        for relationship in relationships:
            if relationship.attrib['Id'] == rId:
                sheet_path = relationship.attrib['Target']
                if relationship.attrib[
                    'Type'] == "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" or \
                        relationship.attrib[
                            'Type'] == 'http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet':
                    sheet_type = 'Macrosheet'
                elif relationship.attrib[
                    'Type'] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
                    sheet_type = 'Worksheet'
                else:
                    sheet_type = 'Unknown'
                break

        return sheet_type, sheet_path

    def get_defined_names(self):
        result = {}
        nsmap = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        workbook = self.get_workbook()
        names = workbook.findall('.//main:definedName', namespaces=nsmap)

        for name in names:
            result[name.attrib['name'].replace('_xlnm.', '').lower()] = name.text

        return result

    def get_macrosheet_infos(self):
        result = []
        nsmap = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        workbook = self.get_workbook()
        sheets = workbook.findall('.//main:sheet', namespaces=nsmap)
        sheet_names = set()
        for sheet_elm in sheets:
            rId = sheet_elm.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            if rId:
                name = sheet_elm.attrib['name']
                sheet_type, rel_path = self.get_sheet_info(rId)
                path = 'xl/' + rel_path
                if sheet_type == 'Macrosheet' and name not in sheet_names:
                    sheet = Boundsheet(name, sheet_type)
                    result.append({'sheet': sheet,
                                   'sheet_path': path,
                                   'sheet_xml': self.get_xml_file(path)})
                    sheet_names.add(name)

        return result

    def load_cells(self, macrosheet, macrosheet_xml):
        nsmap = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        cells = macrosheet_xml.findall('.//main:c', namespaces=nsmap)
        for cell_elm in cells:
            formula = cell_elm.find('./main:f', namespaces=nsmap)
            formula_text = ('=' + formula.text) if formula is not None else None
            value = cell_elm.find('./main:v', namespaces=nsmap)
            value_text = value.text if value is not None else None
            location = cell_elm.attrib['r']
            cell = Cell()
            sheet_name, cell.column, cell.row = Cell.parse_cell_addr(location)
            cell.sheet = macrosheet
            cell.formula = formula_text
            cell.value = value_text
            cell.attribs = cell_elm.attrib
            macrosheet.cells[location] = cell

    def load_row_heights(self, macrosheet, macrosheet_xml):
        nsmap = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        rows = macrosheet_xml.findall('.//main:row[@ht]', namespaces=nsmap)
        for row in rows:
            macrosheet.row_heights[row.attrib['r']] = row.attrib['ht']



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
                self.load_row_heights(macrosheet['sheet'], macrosheet['sheet_xml'])
                self._macrosheets[macrosheet['sheet'].name] = macrosheet['sheet']

        return self._macrosheets


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

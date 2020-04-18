from zipfile import ZipFile
from glob import fnmatch
from xml.etree import ElementTree
import xlm_wrapper


class XLSMWrapper(xlm_wrapper.XLMWrapper):
    def __init__(self, xlsm_doc_path):
        self.xlsm_doc_path = xlsm_doc_path
        self._workbook = None
        self._workbook_rels = None
        self._defined_names = None

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
            result[name.attrib['name'].lower()] = name.text

        return result

    def get_macrosheets(self):
        result = []
        nsmap = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        workbook = self.get_workbook()
        sheets = workbook.findall('.//main:sheet', namespaces=nsmap)
        sheet_names = set()
        for sheet in sheets:
            rId = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            if rId:
                name = sheet.attrib['name']
                sheet_type, rel_path = self.get_sheet_info(rId)
                path = 'xl/' + rel_path
                if sheet_type == 'Macrosheet' and name not in sheet_names:
                    result.append({'sheet_name': name,
                                   'sheet_type': sheet_type,
                                   'sheet_path': path,
                                   'sheet_xml': self.get_xml_file(path)})
                    sheet_names.add(name)

        return result

    def get_xlm_macro(self, macrosheet_xml):
        # formula_cells = {}
        # value_cells = {}
        result_cells = {}
        nsmap = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        cells = macrosheet_xml.findall('.//main:c', namespaces=nsmap)

        for cell in cells:
            formula = cell.find('./main:f', namespaces=nsmap)
            formula_text = ('=' + formula.text) if formula is not None else None
            value = cell.find('./main:v', namespaces=nsmap)
            value_text = value.text if value is not None else None
            location = cell.attrib['r']
            # if formula_text is not None:
            #     formula_cells[location] = {'formula': formula_text,
            #                                'value': value_text}
            # else:
            #     value_cells[location] = {'formula': formula_text,
            #                              'value': value_text}
            result_cells[location] = {'formula': formula_text,
                                      'value': value_text}
        return result_cells

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

    def get_xlm_macros(self):
        result = {}
        auto_open_labels = self.get_defined_name('_xlnm.auto_open', full_match=False)
        for label in auto_open_labels:
            print('auto_open: {}->{}'.format(label[0], label[1]))
        macrosheets = self.get_macrosheets()
        for macrosheet in macrosheets:
            print('SHEET: {}\t{}\t{}'.format(macrosheet['sheet_name'],
                                             macrosheet['sheet_type'],
                                             macrosheet['sheet_path']))
            # formula_cells, value_cells = self.get_xlm_macro(macrosheet['sheet_xml'])
            # result[macrosheet['sheet_name']] = {'formulas': formula_cells,
            #                                     'values': value_cells}
            cells = self.get_xlm_macro(macrosheet['sheet_xml'])
            result[macrosheet['sheet_name']] = {'cells': cells}
        return result


if __name__ == '__main__':

    path = r"C:\Users\user\Downloads\samples\analyze\01558388b33abe05f25afb6e96b0c899221fe75b037c088fa60fe8bbf668f606.xlsm"
    
    xlsm_doc = XLSMWrapper(path)
    macros = xlsm_doc.get_xlm_macros()

    for macrosheet_name in macros:
        print(macrosheet_name)
        for formula_loc, info in macros[macrosheet_name]['cells'].items():
            if info['formula'] is not None:
                print('{}\t{}\t{}'.format(formula_loc, info['formula'], info['value']))

        for formula_loc, info in macros[macrosheet_name]['cells'].items():
            if info['formula'] is None:
                print('{}\t{}\t{}'.format(formula_loc, info['formula'], info['value']))

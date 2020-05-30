from XLMMacroDeobfuscator.excel_wrapper import ExcelWrapper
from XLMMacroDeobfuscator.excel_wrapper import XlApplicationInternational
from pyxlsb2 import open_workbook
from pyxlsb2.formula import Formula
import os
from XLMMacroDeobfuscator.boundsheet import *


class XLSBWrapper(ExcelWrapper):
    def __init__(self, xlsb_doc_path):
        self._xlsb_workbook = open_workbook(xlsb_doc_path)
        self._macrosheets = None
        self._defined_names = None
        self.xl_international_flags = {XlApplicationInternational.xlLeftBracket: '[',
                                       XlApplicationInternational.xlListSeparator: ',',
                                       XlApplicationInternational.xlRightBracket: ']'}

    def get_xl_international_char(self, flag_name):
        result = None
        if flag_name in self.xl_international_flags:
            result = self.xl_international_flags[flag_name]

        return result

    def get_defined_names(self):
        if self._defined_names is None:
            names = {}
            for key, val in self._xlsb_workbook.defined_names.items():
                names[key.lower()] = key.lower(), val.formula
            self._defined_names = names
        return self._defined_names

    def get_defined_name(self, name, full_match=True):
        result = []
        if full_match:
            if name in self.get_defined_names():
                result.append(self.get_defined_names()[name])
        else:
            for defined_name, cell_address in self.get_defined_names().items():
                if defined_name.startswith(name):
                    result.append(cell_address)
        return result

    def load_cells(self, boundsheet):
        with self._xlsb_workbook.get_sheet_by_name(boundsheet.name) as sheet:
            for row in sheet:
                for cell in row:
                    tmp_cell = Cell()
                    tmp_cell.row = cell.row_num + 1
                    tmp_cell.column = Cell.convert_to_column_name(cell.col + 1)

                    tmp_cell.value = cell.value
                    tmp_cell.sheet = boundsheet
                    formula_str = Formula.parse(cell.formula)
                    if formula_str._tokens:
                        try:
                            tmp_cell.formula = '=' + formula_str.stringify(self._xlsb_workbook)
                        except NotImplementedError as exp:
                            print('ERROR({}) {}'.format(exp, str(cell)))
                        except Exception:
                            print('ERROR ' + str(cell))
                    if tmp_cell.value is not None or tmp_cell.formula is not None:
                        boundsheet.cells[tmp_cell.get_local_address()] = tmp_cell

    def get_macrosheets(self):
        if self._macrosheets is None:
            self._macrosheets = {}
            for xlsb_sheet in self._xlsb_workbook.sheets:
                if xlsb_sheet.type == 'macrosheet':
                    with self._xlsb_workbook.get_sheet_by_name(xlsb_sheet.name) as sheet:
                        macrosheet = Boundsheet(xlsb_sheet.name, 'macrosheet')
                        self.load_cells(macrosheet)
                        self._macrosheets[macrosheet.name] = macrosheet

                # self.load_cells(macrosheet, workbook)
                # self._macrosheets[workbook.name] = macrosheet

        return self._macrosheets

    def get_cell_info(self, sheet_name, col, row, info_type_id):
        data = None
        not_exist = False
        not_implemented = True

        return data, not_exist, not_implemented


if __name__ == '__main__':

    # path = r"tmp\xlsb\179ef8970e996201815025c1390c88e1ab2ea59733e1c38ec5dbed9326d7242a"
    path = r"C:\Users\dan\PycharmProjects\xlm\TMP\Doc55752.xlsb"

    path = os.path.abspath(path)
    excel_doc = XLSBWrapper(path)

    macrosheets = excel_doc.get_macrosheets()

    auto_open_labels = excel_doc.get_defined_name('auto_open', full_match=False)
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

from XLMMacroDeobfuscator.excel_wrapper import ExcelWrapper, XlApplicationInternational
from XLMMacroDeobfuscator.boundsheet import Boundsheet
from XLMMacroDeobfuscator.boundsheet import Cell
import xlrd2
import os


class XLSWrapper2(ExcelWrapper):
    XLEXCEL4MACROSHEET = 3

    def __init__(self, xls_doc_path):
        self.xls_workbook = xlrd2.open_workbook(xls_doc_path)
        self._macrosheets = None
        self._defined_names = None
        self.xl_international_flags = {}
        self.xl_international_flags = {XlApplicationInternational.xlLeftBracket: '[',
                                       XlApplicationInternational.xlListSeparator: ',',
                                       XlApplicationInternational.xlRightBracket: ']'}

    def get_xl_international_char(self, flag_name):
        result = None
        if flag_name in self.xl_international_flags:
            result = self.xl_international_flags[flag_name]

        return result

    def get_defined_names(self):
        result = {}

        name_objects = self.xls_workbook.name_map

        for name_obj, cell in name_objects.items():
            result[name_obj.lower()] = cell[0].result.text

        return result

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

    def load_cells(self, macrosheet, xls_sheet):
        try:
            for row in xls_sheet.get_rows():
                if row is not None:
                    for xls_cell in row:
                        cell = Cell()
                        cell.sheet = macrosheet
                        if xls_cell.formula is not None and len(xls_cell.formula)>0:
                            cell.formula = '=' + xls_cell.formula
                        cell.value = xls_cell.value
                        cell.row = xls_cell.row + 1
                        cell.column = Cell.convert_to_column_name(xls_cell.column + 1)
                        if cell.value is not None or cell.formula is not None:
                            macrosheet.add_cell(cell)

        except Exception as error:
            print('CELL(Formula): ' + str(error.args[2]))


    def get_macrosheets(self):
        if self._macrosheets is None:
            self._macrosheets = {}
            for sheet in self.xls_workbook.sheets():
                if sheet.boundsheet_type == xlrd2.biffh.XL_MACROSHEET:
                    macrosheet = Boundsheet(sheet.name, 'Macrosheet')
                    self.load_cells(macrosheet, sheet)
                    self._macrosheets[sheet.name] = macrosheet

        return self._macrosheets


if __name__ == '__main__':

    path = r"C:\Users\user\Downloads\bf58dc1c6ee61d7370c3dfaed7efd98435aed215dfed58e7d90a25b195584b33.xls"

    path = os.path.abspath(path)
    excel_doc = XLSWrapper2(path)

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


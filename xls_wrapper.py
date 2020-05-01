from excel_wrapper import ExcelWrapper
from excel_wrapper import XlApplicationInternational
from boundsheet import Boundsheet
from boundsheet import Cell
from win32com.client import Dispatch
import pywintypes
from enum import Enum
import os


class XlCellType(Enum):
    xlCellTypeFormulas = -4123
    xlCellTypeConstants = 2


class XLSWrapper(ExcelWrapper):
    XLEXCEL4MACROSHEET = 3

    def __init__(self, xls_doc_path):
        self._excel = Dispatch("Excel.Application")
        self.xls_workbook = self._excel.Workbooks.Open(xls_doc_path)
        self._macrosheets = None
        self._defined_names = None
        self.xl_international_flags = {}
        self._international_flags = None

    def get_xl_international_char(self, flag_name):
        if flag_name not in self.xl_international_flags:
            if self._international_flags is None:
                self._international_flags = self._excel.Application.International
            # flag value starts at 1, list index starts at 0
            self.xl_international_flags[flag_name] = self._international_flags[flag_name.value - 1]

        result = self.xl_international_flags[flag_name]
        return result

    def get_defined_names(self):
        result = {}

        name_objects = self.xls_workbook.Excel4MacroSheets.Application.Names

        for name_obj in name_objects:
            result[name_obj.NameLocal.lower()] = str(name_obj.RefersToLocal).strip('=')

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
        cells = {}
        try:
            self._excel.Application.ScreenUpdating = False
            col_offset = xls_sheet.UsedRange.Column
            row_offset = xls_sheet.UsedRange.Row
            for row_no, row in enumerate(xls_sheet.UsedRange.Formula):
                for col_no, col in enumerate(row):
                    if len(col) > 0:
                        cell = Cell()
                        cell.sheet = macrosheet
                        cell.formula = col
                        row_addr = row_offset + row_no
                        col_addr = col_offset + col_no
                        cell.row = row_addr
                        cell.column = Cell.convert_to_column_name(col_addr)
                        if cell.formula is not None:
                            cells[(col_addr, row_addr)] = cell
            self._excel.Application.ScreenUpdating = True

        except pywintypes.com_error as error:
            print('CELL(Formula): ' + str(error.args[2]))

        try:
            for row_no, row in enumerate(xls_sheet.UsedRange.Value):
                for col_no, col in enumerate(row):
                    row_addr = row_offset + row_no
                    col_addr = col_offset + col_no

                    if col is not None:
                        if (col_addr, row_addr) in cells:
                            cell = cells[(col_addr, row_addr)]
                            cell.value = col
                        else:
                            cell = Cell()
                            cell.sheet = macrosheet
                            cell.value = col
                            if cell.value is not None:
                                cells[(col_addr, row_addr)] = cell
        except pywintypes.com_error as error:
            print('CELL(Constant): ' + str(error.args[2]))

        for cell in cells:
            macrosheet.add_cell(cells[cell])

    def get_macrosheets(self):
        if self._macrosheets is None:
            self._macrosheets = {}
            for sheet in self.xls_workbook.Excel4MacroSheets:
                macrosheet = Boundsheet(sheet.name, 'Macrosheet')
                self.load_cells(macrosheet, sheet)
                self._macrosheets[sheet.name] = macrosheet

        return self._macrosheets


if __name__ == '__main__':

    # path = r"tmp\xls\1ed44778fbb022f6ab1bb8bd30849c9e4591dc16f9c0ac9d99cbf2fa3195b326.xls"
    path = r"tmp\xls\edd554502033d78ac18e4bd917d023da2fd64843c823c1be8bc273f48a5f3f5f.xls"

    path = os.path.abspath(path)
    excel_doc = XLSWrapper(path)
    try:
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
    finally:
        excel_doc._excel.Application.DisplayAlerts = False
        excel_doc._excel.Application.Quit()

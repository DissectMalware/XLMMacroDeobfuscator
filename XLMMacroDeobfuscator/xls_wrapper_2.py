from XLMMacroDeobfuscator.excel_wrapper import ExcelWrapper, XlApplicationInternational
from XLMMacroDeobfuscator.boundsheet import Boundsheet
from XLMMacroDeobfuscator.boundsheet import Cell
import xlrd2
import os
import string
import re
import math


class XLSWrapper2(ExcelWrapper):
    XLEXCEL4MACROSHEET = 3

    def __init__(self, xls_doc_path):
        self.xls_workbook = xlrd2.open_workbook(xls_doc_path, formatting_info=True)
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

        for index, (name_obj, cell) in enumerate(name_objects.items()):
            name = name_obj.replace('\x00', '').lower()
            if name in result:
                name = name + index
            result[name] = cell[0].result.text

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
            for xls_cell in xls_sheet.get_used_cells():
                cell = Cell()
                cell.sheet = macrosheet
                if xls_cell.formula is not None and len(xls_cell.formula) > 0:
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

    def get_color(self, color_index):
        return self.xls_workbook.colour_map.get(color_index)

    def get_cell_info(self, sheet_name, col, row, info_type_id):
        sheet = self.xls_workbook.sheet_by_name(sheet_name)
        row = int(row) - 1
        column = Cell.convert_to_column_index(col) - 1
        info_type_id = int(float(info_type_id))

        data = None
        not_exist = True
        not_implemented = False

        if (row, column) in sheet.used_cells:
            cell = sheet.cell(row, column)

            if cell.xf_index is not None and cell.xf_index < len(self.xls_workbook.xf_list):
                fmt = self.xls_workbook.xf_list[cell.xf_index]
                font = self.xls_workbook.font_list[fmt.font_index]
                not_exist = False

                if info_type_id == 8:
                    data = fmt.alignment.hor_align + 1

                # elif info_type_id == 9:
                #     data = fmt.border.left_line_style
                #
                # elif info_type_id == 10:
                #     data = fmt.border.right_line_style
                #
                # elif info_type_id == 11:
                #     data = fmt.border.top_line_style
                #
                # elif info_type_id == 12:
                #     data = fmt.border.bottom_line_style
                #
                # elif info_type_id == 13:
                #     data = fmt.border.fill_pattern
                #
                # elif info_type_id == 14:
                #     data = fmt.protection.cell_locked
                #
                # elif info_type_id == 15:
                #     data = fmt.protection.formula_hidden
                #     return data
                #
                # elif info_type_id == 18:
                #     data = font.name
                #     return data

                elif info_type_id == 19:
                    data = font.height
                    data = Cell.convert_twip_to_point(data)

                # elif info_type_id == 20:
                #     data = font.bold
                #
                # elif info_type_id == 21:
                #     data = font.italic
                #
                # elif info_type_id == 22:
                #     data = font.underlined
                #
                # elif info_type_id == 23:
                #     data = font.struck_out
                
                elif info_type_id == 24:
                    data = font.colour_index - 7

                # elif info_type_id == 25:
                #     data = font.outline
                #
                # elif info_type_id == 26:
                #     data = font.shadow

                # elif info_type_id == 34:
                #     # Left Color index
                #     data = fmt.border.left_colour_index
                #
                # elif info_type_id == 35:
                #     # Right Color index
                #     data = fmt.border.right_colour_index
                #
                # elif info_type_id == 36:
                #     # Top Color index
                #     data = fmt.border.top_colour_index
                #
                # elif info_type_id == 37:
                #     # Bottom Color index
                #     data = fmt.border.bottom_colour_index

                elif info_type_id == 38:
                    data = fmt.background.pattern_colour_index - 7

                elif info_type_id == 50:
                    data = fmt.alignment.vert_align + 1

                # elif info_type_id == 51:
                #     data = fmt.alignment.rotation
                else:
                    not_implemented = True

        elif info_type_id == 17:
            if row in sheet.rowinfo_map:
                not_exist = False
                data = sheet.rowinfo_map[row].height
                data = Cell.convert_twip_to_point(data)
                data = round(float(data) * 4) / 4

        return data, not_exist, not_implemented


if __name__ == '__main__':

    path = r"C:\Users\dan\PycharmProjects\XLMMacroDeobfuscator\tmp\xls\Doc55752.xls"

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

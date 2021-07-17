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

    def __init__(self, xls_doc_path, logfile="default"):
        # Not interested in logging
        if logfile == "default":
            self.xls_workbook = xlrd2.open_workbook(xls_doc_path, formatting_info=True)
        else:
            self.xls_workbook = xlrd2.open_workbook(xls_doc_path, formatting_info=True, logfile=logfile)
        self.xls_workbook_name = os.path.basename(xls_doc_path)
        self._macrosheets = None
        self._worksheets = None
        self._defined_names = None
        self.xl_international_flags = {}
        self.xl_international_flags = {XlApplicationInternational.xlLeftBracket: '[',
                                       XlApplicationInternational.xlListSeparator: ',',
                                       XlApplicationInternational.xlRightBracket: ']'}

        control_chars = ''.join(map(chr, range(0, 32)))
        control_chars += ''.join(map(chr, range(127, 160)))
        control_chars += '\ufefe\uffff\ufeff\ufffe\uffef\ufff0\ufff1\ufff6\ufefd\udddd\ufffd'
        self._control_char_re = re.compile('[%s]' % re.escape(control_chars))

    # from xlrd2
    oBOOL = 3
    oERR = 4
    oMSNG = 5  # tMissArg
    oNUM = 2
    oREF = -1
    oREL = -2
    oSTRG = 1
    oUNK = 0
    oARR = 6

    def get_xl_international_char(self, flag_name):
        result = None
        if flag_name in self.xl_international_flags:
            result = self.xl_international_flags[flag_name]

        return result

    def replace_nonprintable_chars(self, input_str, replace_char=''):
        input_str = input_str.encode("utf-16").decode('utf-16', 'ignore')
        return self._control_char_re.sub(replace_char, input_str)

    def get_defined_names(self):
        if self._defined_names is None:
            self._defined_names = {}

            name_objects = self.xls_workbook.name_map

            for index, (name_obj, cells) in enumerate(name_objects.items()):
                name = name_obj.lower()
                if len(cells) > 1:
                    index = 1
                else:
                    index = 0

                # filtered_name = self.replace_nonprintable_chars(name, replace_char='_').lower()
                filtered_name = name.lower()
                if name != filtered_name:
                    if filtered_name in self._defined_names:
                        filtered_name = filtered_name + str(index)
                    if cells[0].result is not None:
                        self._defined_names[filtered_name] = cells[0].result.text

                if name in self._defined_names:
                    name = name + str(index)
                if cells[0].result is not None:
                    cell_location = cells[0].result.text
                    if cells[0].result.kind == XLSWrapper2.oNUM:
                        self._defined_names[name] = cells[0].result.value
                    elif cells[0].result.kind == XLSWrapper2.oSTRG:
                        self._defined_names[name] = cells[0].result.text
                    elif cells[0].result.kind == XLSWrapper2.oARR:
                        self._defined_names[name] = cells[0].result.value
                    elif cells[0].result.kind == XLSWrapper2.oREF:
                        if '$' in cell_location:
                            self._defined_names[name] = cells[0].result.text
                        else:
                            # By @JohnLaTwC:
                            # handled mangled cell name as in:
                            # 8a868633be770dc26525884288c34ba0621170af62f0e18c19b25a17db36726a
                            # defined name auto_open found at Operand(kind=oREF, value=[Ref3D(coords=(1, 2, 321, 322, 14, 15))], text='sgd7t!\x00\x00\x00\x00\x00\x00\x00\x00')

                            curr_cell = cells[0].result
                            if ('auto_open' in name):
                                coords = curr_cell.value[0].coords
                                r = int(coords[3])
                                c = int(coords[5])
                                sheet_name = curr_cell.text.split('!')[0].replace("'", '')
                                cell_location_xlref = sheet_name + '!' + self.xlref(row=r, column=c, zero_indexed=False)
                                self._defined_names[name] = cell_location_xlref

        return self._defined_names

    def xlref(self, row, column, zero_indexed=True):

        if zero_indexed:
            row += 1
            column += 1
        return '$' + Cell.convert_to_column_name(column) + '$' + str(row)

    def get_defined_name(self, name, full_match=True):
        result = []
        name = name.lower().replace('[', '')

        if full_match:
            if name in self.get_defined_names():
                result = self._defined_names[name]
        else:
            for defined_name, cell_address in self.get_defined_names().items():
                if defined_name.startswith(name):
                    result.append((defined_name, cell_address))

        # By @JohnLaTwC:
        # if no matches, try matching 'name' by looking for its characters
        # in the same order (ignoring junk chars from UTF16 etc in between. Eg:
        # Auto_open:
        #   match:    'a_u_t_o___o__p____e_n'
        #   not match:'o_p_e_n_a_u_to__'
        # Reference: https://malware.pizza/2020/05/12/evading-av-with-excel-macros-and-biff8-xls/
        # Sample: e23f9f55e10f3f31a2e76a12b174b6741a2fa1f51cf23dbd69cf169d92c56ed5
        if isinstance(result, list) and len(result) == 0:
            for defined_name, cell_address in self.get_defined_names().items():
                lastidx = 0
                fMatch = True
                for c in name:
                    idx = defined_name.find(c, lastidx)
                    if idx == -1:
                        fMatch = False
                    lastidx = idx
                if fMatch:
                    result.append((defined_name, cell_address))
                ##print("fMatch for %s in %s is %d:" % (name,defined_name, fMatch))

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

    def get_workbook_name(self):
        return self.xls_workbook_name

    def get_worksheets(self):
        if self._worksheets is None:
            self._worksheets = {}
            for sheet in self.xls_workbook.sheets():
                if sheet.boundsheet_type == xlrd2.biffh.XL_WORKSHEET:
                    worksheet = Boundsheet(sheet.name, 'Worksheet')
                    self.load_cells(worksheet, sheet)
                    self._worksheets[sheet.name] = worksheet

        return self._worksheets

    def get_color(self, color_index):
        return self.xls_workbook.colour_map.get(color_index)

    def get_cell_info(self, sheet_name, col, row, info_type_id):
        sheet = self.xls_workbook.sheet_by_name(sheet_name)
        row = int(row) - 1
        column = Cell.convert_to_column_index(col) - 1
        info_type_id = int(float(info_type_id))

        data = None
        not_exist = False
        not_implemented = False

        if info_type_id == 5:
            data = sheet.cell(row, column).value

        elif info_type_id == 17:
            not_exist = False
            if row in sheet.rowinfo_map:
                data = sheet.rowinfo_map[row].height
            else:
                data = sheet.default_row_height
            data = Cell.convert_twip_to_point(data)
            data = round(float(data) * 4) / 4
        else:
            if (row, column) in sheet.used_cells:
                cell = sheet.cell(row, column)
                if cell.xf_index is not None and cell.xf_index < len(self.xls_workbook.xf_list):
                    fmt = self.xls_workbook.xf_list[cell.xf_index]
                    font = self.xls_workbook.font_list[fmt.font_index]

                else:
                    normal_style = self.xls_workbook.style_name_map['Normal'][1]
                    fmt = self.xls_workbook.xf_list[normal_style]
                    font = self.xls_workbook.font_list[fmt.font_index]
            else:
                normal_style = self.xls_workbook.style_name_map['Normal'][1]
                fmt = self.xls_workbook.xf_list[normal_style]
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
                data = font.colour_index - 7 if font.colour_index > 7 else font.colour_index

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
                data = fmt.background.pattern_colour_index - 7 if font.colour_index > 7 else font.colour_index

            elif info_type_id == 50:
                data = fmt.alignment.vert_align + 1

            # elif info_type_id == 51:
            #     data = fmt.alignment.rotation
            else:
                not_implemented = True

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

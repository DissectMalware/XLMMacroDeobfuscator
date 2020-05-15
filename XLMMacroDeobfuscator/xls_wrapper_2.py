from XLMMacroDeobfuscator.excel_wrapper import ExcelWrapper, XlApplicationInternational
from XLMMacroDeobfuscator.boundsheet import Boundsheet
from XLMMacroDeobfuscator.boundsheet import Cell
import xlrd2
import os
import string
import re


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

        for index, (name_obj, cell) in enumerate(name_objects.items()):
            name = name_obj.replace('\x00','').lower()
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

    def col2num(self,col):
        num = 0
        for c in col:
            if c in string.ascii_letters:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num

    def get_color(self,color_index):
        return self.xls_workbook.colour_map.get(color_index)

    def twipToPoint(self,twips):
        #Xlrd has the data for font height and row height in twips.
        #Excel GET.CELL(17 OR 19) returns in points.
        #I change the twips (which are 1/20 of a point) into points so it matches what excel would output

        point = int(twips) * 0.050000283464388
        return point

    def cell_info(self,sheet, cell, type_ID):
        #sheet, column, row = Cell.parse_cell_addr(cell)
        sht =  self.xls_workbook.sheet_by_name(sheet)
        cellParse = re.compile("([a-zA-Z]+)([0-9]+)")
        cellData = cellParse.match(cell).groups()
        column = cellData[0]
        row = cellData[1]
        print(row,column)
        row = int(row) - 1
        column = Cell.convert_to_column_index(column) - 1
        w = sht.computed_column_width(0)
        cell = sht.cell(row, column)
        fmt = self.xls_workbook.xf_list[cell.xf_index]
        font = self.xls_workbook.font_list[fmt.font_index]
        border = fmt.border
        #
        # if int(type_ID) == 2:
        #     data = sht.Range(cell).Row
        #     print(data)
        #     return data
        #
        # elif int(type_ID) == 3:
        #     data = sht.Range(cell).Column
        #     print(data)
        #     return data


        if int(type_ID) == 8:
            data = fmt.alignment.hor_align
            return data

        elif int(type_ID) == 9:
            # GET.CELL(9,cell)
            data = border.left_line_style
            return data

        elif int(type_ID) == 10:
            # GET.CELL(9,cell)
            data = border.right_line_style
            return data

        elif int(type_ID) == 11:
            # GET.CELL(9,cell)
            data = border.top_line_style
            return data

        elif int(type_ID) == 12:
            # GET.CELL(9,cell)
            data = border.bottom_line_style
            return data

        elif int(type_ID) == 13:
            # GET.CELL(9,cell)
            data = border.fill_pattern
            return data

        elif int(type_ID) == 14:
            # GET.CELL(9,cell)
            data = fmt.protection.cell_locked
            return data

        elif int(type_ID) == 15:
            data = fmt.protection.formula_hidden
            return data

        elif int(type_ID) == 17:
            #get row height
            data = sht.rowinfo_map[row].height
            data = self.twipToPoint(data)
            return data

        elif int(type_ID) == 18:
            #get font name
            data = font.name
            return data

        elif int(type_ID) == 19:
            #get font height
            data = font.height
            data = self.twipToPoint(data)
            return data

        elif int(type_ID) == 20:
            #check if bold
            data = font.bold
            return data

        elif int(type_ID) == 21:
            #check if italic
            data = font.italic
            return data

        elif int(type_ID) == 22:
            data = font.underlined
            return data

        elif int(type_ID) == 23:
            #Check if font has strikethrough
            data = font.struck_out
            return data

        elif int(type_ID) == 24:
            #NOT FINISHED
            data = self.get_color(font.colour_index)
            return data

        elif int(type_ID) == 25:
            data = font.outline
            return data

        elif int(type_ID) == 26:
            data = font.shadow
            return data

        elif int(type_ID) == 34:
            #Left Color index
            data = border.left_colour_index
            return data

        elif int(type_ID) == 35:
            #Right Color index
            data = border.right_colour_index
            return data


        elif int(type_ID) == 36:
            #Top Color index
            data = border.top_colour_index
            return data


        elif int(type_ID) == 37:
            #Bottom Color index
            data = border.bottom_colour_index
            return data


        elif int(type_ID) == 50:
            data = fmt.alignment.vert_align
            return data


        elif int(type_ID) == 51:
            data = fmt.alignment.rotation
            return data


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


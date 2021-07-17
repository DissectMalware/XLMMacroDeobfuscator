import re


class Cell:
    _a1_cell_addr_regex_str = r"((?P<sheetname>[^\s]+?|'.+?')!)?\$?(?P<column>[a-zA-Z]+)\$?(?P<row>\d+)"
    _a1_cell_addr_regex = re.compile(_a1_cell_addr_regex_str)

    _r1c1_abs_cell_addr_regex_str = r"((?P<sheetname>[^\s]+?|'.+?')!)?R(?P<row>\d+)C(?P<column>\d+)"
    _r1c1_abs_cell_addr_regex = re.compile(_r1c1_abs_cell_addr_regex_str)

    _r1c1_cell_addr_regex_str = r"((?P<sheetname>[^\s]+?|'.+?')!)?R(\[?(?P<row>-?\d+)\]?)?C(\[?(?P<column>-?\d+)\]?)?"
    _r1c1_cell_addr_regex = re.compile(_r1c1_cell_addr_regex_str)

    _range_addr_regex_str = r"((?P<sheetname>[^\s]+?|'.+?')[!|$])?\$?(?P<column1>[a-zA-Z]+)\$?(?P<row1>\d+)\:?\$?(?P<column2>[a-zA-Z]+)\$?(?P<row2>\d+)"
    _range_addr_regex = re.compile(_range_addr_regex_str)

    def __init__(self):
        self.sheet = None
        self.column = ''
        self.row = 0
        self.formula = None
        self.value = None
        self.attributes = {}
        self.is_set = False

    def get_attribute(self, attribute_name):
        # return default value if attributes doesn't cointain the attribute_name
        pass
    
    def __deepcopy__(self, memodict={}):
        copy = type(self)()
        memodict[id(self)] = copy
        copy.sheet = self.sheet
        copy.column = self.column
        copy.row = self.row
        copy.formula = self.formula
        copy.value = self.value
        copy.attributes = self.attributes
        return copy

    def get_local_address(self):
        return self.column + str(self.row)

    def __str__(self):
        return "'{}'!{}".format(self.sheet.name,self.get_local_address())

    @staticmethod
    def convert_to_column_index(s):
        number = 0
        power = 1
        for character in reversed(s):
            character = character.upper()
            digit = ((ord(character) - ord('A'))+1) * power
            number = number + digit
            power = power * 26

        return number

    @staticmethod
    def convert_to_column_name(n):
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(ord('A') + remainder) + string
        return string

    @staticmethod
    def parse_cell_addr(cell_addr_str):
        cell_addr_str = cell_addr_str.strip('\"')
        alternate_res = Cell._r1c1_abs_cell_addr_regex.match(cell_addr_str)
        if alternate_res is not None:
            sheet_name = alternate_res.group('sheetname')
            sheet_name = sheet_name.strip('\'') if sheet_name is not None else sheet_name
            column = Cell.convert_to_column_name(int(alternate_res.group('column')))
            row = alternate_res.group('row')
            return sheet_name, column, row
        else:
            res = Cell._a1_cell_addr_regex.match(cell_addr_str)
            if res is not None:
                sheet_name = res.group('sheetname')
                sheet_name = sheet_name.strip('\'') if sheet_name is not None else sheet_name
                column = res.group('column')
                row = res.group('row')
                return sheet_name, column, row
            else:
                return None, None, None

    @staticmethod
    def parse_range_addr(range_addr_str):
        res = Cell._range_addr_regex.match(range_addr_str)
        if res is not None:
            sheet_name = res.group('sheetname')
            sheet_name = sheet_name.strip('\'') if sheet_name is not None else sheet_name
            startcolumn = res.group('column1')
            startrow = res.group('row1')
            endcolumn = res.group('column2')
            endrow = res.group('row2')
            return sheet_name, startcolumn, startrow, endcolumn, endrow
        else:
            return None, None, None

    @staticmethod
    def convert_twip_to_point(twips):
        # A twip is 1/20 of a point
        point = int(twips) * 0.05
        return point

    @staticmethod
    def get_abs_addr(base_addr, offset_addr):
        _, base_col, base_row = Cell.parse_cell_addr(base_addr)
        offset_addr_match = Cell._r1c1_cell_addr_regex.match(offset_addr)
        column_offset = row_offset = 0
        if offset_addr_match is not None:
            column_offset = int(offset_addr_match.group('column'))
            row_offset = int(offset_addr_match.group('row'))

        res_col_index = Cell.convert_to_column_index(base_col) + column_offset
        res_row_index = int(base_row) + row_offset

        return Cell.convert_to_column_name(res_col_index)+str(res_row_index)


class Boundsheet:
    def __init__(self, name, type):
        self.name = name
        self.type = type
        self.cells = {}
        self.row_attributes = {}
        self.col_attributes = {}
        self.default_height = None

    def get_row_attribute(self, row, attrib_name):
        # default values if row doesn't exist in row_attributes
        pass

    def get_col_attribute(self, col, attrib_name):
        # default value if row doesn't exist in row_attributes

        pass


    def add_cell(self, cell):
        cell.sheet = self
        self.cells[cell.get_local_address()] = cell

    def get_cell(self, local_address):
        result = None
        if local_address in self.cells:
            result = self.cells[local_address]
        return


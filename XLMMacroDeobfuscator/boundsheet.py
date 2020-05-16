import re


class Cell:
    _cell_addr_regex_str = r"((?P<sheetname>[^\s]+?|'.+?')!)?\$?(?P<column>[a-zA-Z]+)\$?(?P<row>\d+)"
    _cell_addr_regex = re.compile(_cell_addr_regex_str)

    def __init__(self):
        self.sheet = None
        self.column = ''
        self.row = 0
        self.formula = None
        self.value = None
        self.attributes = {}

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
        for character in s:
            character = character.upper()
            digit = (ord(character) - ord('A')) * power
            number = number + digit
            power = power * 26

        return number + 1

    @staticmethod
    def convert_to_column_name(n):
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    @staticmethod
    def parse_cell_addr(cell_addr_str):
        res = Cell._cell_addr_regex.match(cell_addr_str)
        sheet_name = res.group('sheetname')
        sheet_name = sheet_name.strip('\'') if sheet_name is not None else sheet_name
        column = res.group('column')
        row = res.group('row')

        return sheet_name, column, row

    @staticmethod
    def convert_twip_to_point(twips):
        # A twip is 1/20 of a point
        point = int(twips) * 0.05
        return point


class Boundsheet:
    def __init__(self, name, type):
        self.name = name
        self.type = type
        self.cells = {}
        self.row_attributes = {}
        self.col_attributes = {}

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

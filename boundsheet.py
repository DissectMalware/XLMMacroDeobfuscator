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

    def get_local_address(self):
        return self.column + str(self.row)

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
        sheet_name = (res['sheetname'].strip('\'')) if (res['sheetname'] is not None) else None
        column = res['column'] if 'column' in res.re.groupindex else None
        row = res['row'] if 'row' in res.re.groupindex else None

        return sheet_name, column, row

class Boundsheet:
    def __init__(self, name, type):
        self.name = name
        self.type = type
        self.cells = {}

    def add_cell(self, cell):
        cell.sheet = self
        self.cells[cell.get_local_address()] = cell

    def get_cell(self, local_address):
        result = None
        if local_address in self.cells:
            result = self.cells[local_address]
        return
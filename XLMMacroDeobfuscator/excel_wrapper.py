from enum import Enum


class ExcelWrapper:
    def get_xl_international_char(self, flag_name):
        pass

    def get_defined_name(self, name, full_match):
        pass

    def get_defined_names(self):
        pass

    def get_macrosheets(self):
        pass

    def get_row_attribute(self, row, attrib_name):
        pass

    def get_col_attribute(self, row, attrib_name):
        pass

    def get_cell_attribute(self, row, attrib_name):
        pass


class XlApplicationInternational(Enum):
    # https://docs.microsoft.com/en-us/office/vba/api/excel.xlapplicationinternational
    xlLeftBracket = 10
    xlListSeparator = 5
    xlRightBracket = 11

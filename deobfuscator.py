from win32com.client import Dispatch
import re
import os


def column_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


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


if __name__ == '__main__':
    XLEXCEL4MACROSHEET = 3

    excel = Dispatch("Excel.Application")

    zloader_samples_dir = r"C:\Users\user\Downloads\samples\zloader"

    with open(r'result\zloader.txt', 'w', encoding='utf_8') as output:
        for excel_file in os.listdir(zloader_samples_dir):
            try:
                print(excel_file)
                output.write('SHA256: ' + excel_file + '\n')
                wb = excel.Workbooks.Open(os.path.join(zloader_samples_dir, excel_file))
                for workbook in wb.Sheets:
                    if workbook.type == XLEXCEL4MACROSHEET:
                        workbook = wb.Sheets[1]
                        print(workbook.name)
                        output.write(workbook.name + '\n')

                        formulas = get_formula_cells(workbook)
                        for loc, formula in formulas.items():
                            exist, formulas[loc] = interpret_char_function(formula)

                        for loc, formula in formulas.items():
                            exist, result = interpret_formula_function(formula, formulas)
                            if exist:
                                print(result)
                                output.write(result + '\n')
            except Exception as error:
                print(error)
            finally:
                output.write('\n')
                output.flush()

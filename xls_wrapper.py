import excel_wrapper
from utils import *

class XLSWrapper(excel_wrapper.ExcelWrapper):
    pass


def get_formula_cells(sheet):
    formulas = {}
    for cell in sheet.UsedRange:
        formula = cell.FormulaLocal
        if formula and formula.startswith('='):
            formulas[column_string(cell.column) + str(cell.row)] = formula
    return formulas


def process_cell(sheet, cell, depth):

    next_cell = None
    next_depth = None
    value = None

    text = cell.Text

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


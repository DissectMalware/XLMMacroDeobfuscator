import xlrd2
import string
import re
wb = xlrd2.open_workbook(r"C:\Users\dan\PycharmProjects\XLMMacroDeobfuscator\tmp\xls\Doc55752.xls", formatting_info=True)
stuff = "S90"

cellParse = re.compile("([a-zA-Z]+)([0-9]+)")
cellData = cellParse.match(stuff).groups()
column = cellData[0]
row = cellData[1]


def column_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num
#print(column,row,col2num(column))
#print(column)
sheet = wb.sheet_by_index(1)
print(sheet.name)

w = sheet.computed_column_width(0)
cell = sheet.cell(80,0)
fmt = wb.xf_list[cell.xf_index]
font = wb.font_list[fmt.font_index]
border = fmt.border
#print(fmt.dump(),font.dump())
print(font.height, sheet.rowinfo_map[80].height )

print(font.colour_index, wb.colour_map.get(font.colour_index))
# if int(type_ID) == 2:
#     data = sht.Range(cell).Row
#     print(data)
#     return data
#
# elif int(type_ID) == 3:
#     data = sht.Range(cell).Column
#     print(data)
#     return data
#
#
# elif int(type_ID) == 8:
#     data = fmt.alignment.hor_align
#     print(data)
#     return data
#
# elif int(type_id) == 9:
#     # GET.CELL(9,cell)
#     data = fmt.border.left_line_style
#     print("left border " + str(fmt.border.left_line_style))
#     return data
#
# elif int(type_id) == 10:
#     # GET.CELL(9,cell)
#     data = fmt.border.right_line_style
#     return data
#
# elif int(type_id) == 11:
#     # GET.CELL(9,cell)
#     data = fmt.border.top_line_style
#     return data
#
# elif int(type_id) == 12:
#     # GET.CELL(9,cell)
#     data = fmt.border.bottom_line_style
#     return data
#
# elif int(type_id) == 13:
#     # GET.CELL(9,cell)
#     data = fmt.border.fill_pattern
#     return data
#
# elif int(type_id) == 14:
#     # GET.CELL(9,cell)
#     data = fmt.protection.cell_locked
#     return data
#
# elif int(type_id) == 15:
#     data = fmt.protection.formula_hidden
#     return data
#
# elif int(type_ID) == 17:
#     data = sheet.rowinfo_map[row].height
#     print(data)
#     return data
# elif int(type_ID) == 19:
#     data = font.height
#     print(data)
#     return data
# elif int(type_ID) == 20:
#     data = font.bold
#     print(data)
#     return data
# elif int(type_ID) == 21:
#
#     data =  font.italic
#     print(data)
#     return data
# elif int(type_ID) == 22:
#     data = font.underlined
#     return data
#
# elif int(type_ID) == 23:
#     data = font.struck_out
#
#     print(data)
#     return data
#
# elif int(type_ID) == 24:
#     colour_index = font.colour_index
#     data = self.xls_workbook.colour_map.get(colour_index)
#     print(data)
#     return data
#
#
# elif int(type_ID) == 25:
#     data = font.outline
#
#     print(data)
#     return data
#
#
#
# elif int(type_ID) == 26:
#     data = font.shadow
#
#     print(data)
#     return data
#
#
# # GET.CELL(8,cell)
# print(fmt.alignment.hor_align)
#
# # GET.CELL(9,cell)
# left_border = fmt.border.left_line_style
# print("left border " + str(fmt.border.left_line_style))
#
# # GET.CELL(10,cell)
# right_border = fmt.border.right_line_style
# print("right border " + str(fmt.border.right_line_style))
#
# # GET.CELL(11,cell)
# top_border = fmt.border.top_line_style
# print("top border " + str(fmt.border.top_line_style))
#
# # GET.CELL(12,cell)
# bottom_border = fmt.border.bottom_line_style
# print("bottom border " + str(fmt.border.bottom_line_style))
#
# # GET.CELL(13,cell)
# pattern = fmt.background.fill_pattern
# print("pattern " + str(pattern))
#
# # GET.CELL(14,cell)
# cell_locked = fmt.protection.cell_locked
# print("Cell locked: " + str(cell_locked))
#
# # GET.CELL(15,cell)
# formula_hidden = fmt.protection.formula_hidden
# print("formula hidden:" + str(formula_hidden))
#
# # GET.CELL(17,cell)
# row_height = sheet.rowinfo_map[5].height
# print(row_height)
# # GET.CELL(19,cell)
# font_height = wb.font_list[fmt.font_index].height
# print(wb.font_list[fmt.font_index].height)
#
# # GET.CELL(20,cell)
# cell_bold = wb.font_list[fmt.font_index].bold
# print("is bold {}".format(cell_bold))
# # GET.CELL(21,cell)
# cell_italic =
# print("is italics {}".format(cell_italic))
# # GET.CELL(22,cell)
# cell_underlined = font.underlined
# print("is underline {}".format(cell_underlined))
# # GET.CELL(23,cell)
# cell_strike = font.struck_out
# print("struck out: " + str(cell_strike))
#
# print(cell)
#
# # elif parse_tree_root.data == 'method_call':
# # text = self.convert_parse_tree_to_str(parse_tree_root)
# # current_sheet = current_cell.sheet.name
# # if "GET.CELL" in text:
# #     test = text.replace("GET.CELL(", "").replace(")", "").split(",")
# #     type_num = test[0]
# #     location = test[1]
# #     # location = re.findall("\$?([a-zA-Z])+(\$?\d+)",location)
# #     # col = location[0][0]
# #     # row = location[0][1]
# #     print(location, current_sheet, current_cell.formula)
# #     data = self.xlm_wrapper.cell_info(current_sheet, location, type_num)
# #     text = int(data)
# #     # text = self.xlm_parser.parse(text)
# #     # cell_type_num.type_num(self.xlm_wrapper,type_num, current_sheet, location)
#
#
# print(col2num("bs"))
#
# if int(type_ID) == 2:
#     data = sht.Range(cell).Row
#     print(data)
#     return data
#
# elif int(type_ID) == 3:
#     data = sht.Range(cell).Column
#     print(data)
#     return data
#
#
# elif int(type_ID) == 8:
#     data = sht.Range(cell).HorizontalAlignment
#     print(data)
#     return data
#
# elif int(type_ID) == 17:
#     data = sht.Range(cell).Height
#     print(data)
#     return data
# elif int(type_ID) == 19:
#     data = sht.Range(cell).Font.Size
#     print(data)
#     return data
# elif int(type_ID) == 20:
#     data = sht.Range(cell).Font.Bold
#     print(data)
#     return data
# elif int(type_ID) == 21:
#     data = sht.Range(cell).Font.Italic
#     print(data)
#     return data
# elif int(type_ID) == 23:
#     data = sht.Range(cell).Font.Italic
#     print(data)
#     return data
# elif int(type_ID) == 24:
#     data = sht.Range(cell).Font.ColorIndex
#     print(data)
#     return data

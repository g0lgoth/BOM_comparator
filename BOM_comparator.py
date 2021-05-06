import os
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl import load_workbook

excel_directory_move = os.getcwd()
print(excel_directory_move)

wb1 = load_workbook('{}\\E0000237-Auxiliary_Module_1-04_03_21.xls'.format(excel_directory_move))
wb2 = load_workbook('{}\\BOM-E0000237-A01_initial.xlsx'.format(excel_directory_move))
wb3 = Workbook()

active_row = 13
id_column = 6
cursor_value = 0
check_table = []

class excelManagement():

    def __init__(self, wb):
        ws = self.wb.active
        
    def get_value(self, row_cursor, col_cursor)
        location = ws.cell(row=row_cursor, column=col_cursor)
        cell_value= location.value
        return (cell_value)

def checkValue(wb, row_cursor, col_cursor):
    ws = wb.active
    location = ws.cell(row=row_cursor, column=col_cursor)
    cell_value= location.value
    return (cell_value)

print(checkValue(wb1, active_row, id_column))
value = checkValue(wb1, active_row, id_column)
check_table.append(value)
print(check_table)
# for row in ws1.values:
#     for value in row:
#         print(value)

# wb2 = Workbook()
# xlrd.open_workbook('{}\\BOM-E0000237-A01.xlsx'.format(excel_directory_move))
# rb2 = xlrd.open_workbook('E0000237-Auxiliary_Modile_1-04_03_21.xlsx')
# sheet1 = rb1.sheet_by_index(0)
# sheet2 = rb2.sheet_by_index(0)

# value_min = 13

# def maxRowValue(sheet):
#     for col in sheet:
#         for row in sheet:
#             if row != None:
#                 continue
#             else:
#                 return row.value

# sheet1_max_value = maxRowValue(sheet1)
# sheet2_max_value = maxRowValue(sheet2)
# print(sheet1_max_value, sheet2_max_value)
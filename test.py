from readwrite_class import *

readxl = Readxl("input.xlsx")

read_workbook = readxl.read_xlsx_workbook()
print(read_workbook)

read_sheet = readxl.read_wlsx_worksheet(read_workbook, 0)
print(read_sheet)

(A,B) = readxl.get_rowcol_len(read_sheet)
print(A)
print(B)
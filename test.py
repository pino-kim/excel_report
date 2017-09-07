from readwrite_class import *

readxl = Readxl("input.xlsx")

read_workbook = readxl.read_xlsx_workbook()
print(read_workbook)

read_sheet = readxl.read_wlsx_worksheet(read_workbook, 0)
print(read_sheet)

# length and data
print(readxl.get_rowcol_len(read_sheet))
print(readxl.get_row_by_Index(read_sheet, 2))
print(readxl.get_col_by_Index(read_sheet, 0))

#worksheet = read_sheet
#c0 = [worksheet.row_values(i)[0] for i in range(worksheet.nrows ) if worksheet.row_values(i)[0]]
#print (c0)
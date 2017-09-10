from readwrite_class import *

readxl = Readxl("input.xlsx")

read_workbook = readxl.read_xlsx_workbook()
print(read_workbook)

read_sheet = readxl.read_wlsx_worksheet(read_workbook, 0)
print(read_sheet)

# length and data
print(readxl.get_rowcol_len(read_sheet))
print(readxl.get_data_by_row_index(read_sheet, 2))
(a,b) = readxl.get_data_by_row_index(read_sheet, 0)
print(type(a))
print(type(b))

#worksheet = read_sheet
#c0 = [worksheet.row_values(i)[0] for i in range(worksheet.nrows ) if worksheet.row_values(i)[0]]
#print (c0)
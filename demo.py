from Readxl import *
from Writexl import *

# Readxl class example
readxl = Readxl("input.xlsx")

# read workbook, worksheet
rd_workbook = readxl.read_xlsx_workbook()
#print(rd_workbook)
rd_worksheet = readxl.read_xlsx_worksheet(rd_workbook, 0)
#print(rd_worksheet)

# get data first  section A colum
(data_len, data) = readxl.read_data_from_col(rd_worksheet,0)

# close readed workbook
readxl.read_xlsx_close(rd_workbook)
del readxl


# Writexl class example
writexl = Writexl('output.xlsx')

# write workbook, worksheet
wt_workbook = writexl.write_xlsx_workbook()
#print(wt_workbook)
wt_worksheet = writexl.add_xlsx_worksheet(wt_workbook)
#print(wt_workbook)

# write string 'data length' to A1. Cell point  is (0,0).
format = writexl.set_cell_format(wt_workbook, 'title')
format2 = writexl.set_cell_format(wt_workbook, 'label')

#set row and column size
writexl.set_row_size(wt_worksheet, 1, 10,)
writexl.set_col_size(wt_worksheet,'A', 'F', 20)

writexl.write_cell_by_cellname(wt_worksheet ,'A1', 'data length', format)

# write integer 'data length' to A2.  Cell point  is (0,1).
writexl.write_cell_by_rowcal(wt_worksheet, 1,0,  data_len, format)

# write string 'data list' to B1. Cell point  is (1,0).
writexl.write_cell_by_cellname(wt_worksheet ,'B1', 'data list', format)

# write integer data to AB.  Cell point  is (0,1).
#Write numbers down from cell' B2.'
writexl.write_data_to_col(wt_worksheet, 'B2', data, format2)

#close workbook
writexl.write_xlsx_close(wt_workbook)
del writexl
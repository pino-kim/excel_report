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
(data_len, data) = readxl.get_data_from_col(rd_worksheet,0)

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
writexl.write_cell_by_cellname(wt_worksheet ,'A1', 'data length')

# write integer 'data length' to A2.  Cell point  is (0,1).
writexl.write_cell_by_rowcal(wt_worksheet, 1,0,  data_len)

# write string 'data list' to B1. Cell point  is (1,0).
writexl.write_cell_by_cellname(wt_worksheet ,'B1', 'data list')

# write integer data to AB.  Cell point  is (0,1).
#Write numbers down from cell' B2.'
writexl.write_data_to_col(wt_worksheet, 'B2', data)

#close workbook
writexl.write_xlsx_close(wt_workbook)
del writexl
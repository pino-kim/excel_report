from Readxl import *
from Writexl import *

#file_name : "input.xlsx"
#file_name = str(input())

file_name = "input.xlsx"
(data_len1, data1) = read_col_data(file_name, 0, 0)
(data_len2, data2) = read_col_data(file_name, 0, 3)

writexl = Writexl('output.xlsx')


(workbook, worksheet) = writexl_init(writexl)

writexl_label_title(writexl, workbook, worksheet)

writexl.write_xlsx_close(workbook)




from excel_con import *

#demo : http://xlsxwriter.readthedocs.io/worksheet.html#write_formula

writexl = Writexl()

writexl.set_xlsx_title("AAA.xlsx", "BBB")

writexl.open_xlsx_workbook()

workbook = writexl.open_xlsx_workbook()

worksheet = writexl.set_xlsx_worksheet(workbook)

writexl.write_cell_by_point(worksheet, 'A1', 123)

writexl.close_xlsx(workbook)



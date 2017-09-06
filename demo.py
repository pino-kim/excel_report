from itertools import takewhile
import xlrd
import xlsxwriter
from openpyxl import worksheet

book = xlrd.open_workbook('input.xlsx', on_demand = True)
sheet = book.sheet_by_index(0)

#row and col length
row_len = sheet.nrows
col_len = sheet.ncols

#read all data colum
data = [[str(c.value) for c in sheet.col(i)] for i in range(col_len)]
print(data)

#release excel
book.release_resources()
del book

workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()

data[0].pop(0)

worksheet.write_column('A2', data[0])
worksheet.write_column('B3', data[2])

workbook.close()
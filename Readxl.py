import xlrd

class Readxl :
    def __init__(self, workbook_titile) :
        self.workbook_title = workbook_titile


    def read_xlsx_workbook(self) :
        workbook = xlrd.open_workbook( self.workbook_title, on_demand = True )
        return workbook


    def read_xlsx_worksheet(self,workbook, sheet_num) :
        sheet = workbook.sheet_by_index(sheet_num)
        return sheet


    def get_rowcol_len(self, sheet) :
        row_len = sheet.nrows
        col_len = sheet.ncols
        return (row_len, col_len)


    def read_data_from_row(self, sheet, index, start_row = 0) :
        data = sheet.row_values(index, start_row)
        data_len = len(data)
        return (data_len, data)


    def read_data_from_col(self, sheet, index, start_col = 0) :
        data = sheet.col_values(index, start_col)
        data_len = len(data)
        return (data_len, data)

    def read_xlsx_close(self,workbook) :
        workbook.release_resources()


def read_col_data(file_name, sheet_num, col_num) :
    # Readxl class example
    readxl = Readxl(file_name)

    # read workbook, worksheet
    rd_workbook = readxl.read_xlsx_workbook()
    # print(rd_workbook)
    rd_worksheet = readxl.read_xlsx_worksheet(rd_workbook, sheet_num)
    # print(rd_worksheet)

    # get data first  section A colum
    (data_len, data) = readxl.read_data_from_col(rd_worksheet, col_num)

    # close readed workbook
    readxl.read_xlsx_close(rd_workbook)
    del readxl

    return (data_len, data)
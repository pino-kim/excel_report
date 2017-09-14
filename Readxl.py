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


    def get_data_from_row(self, sheet, index) :
        data = sheet.row_values(index)
        data_len = len(data)
        return (data_len, data)


    def get_data_from_col(self, sheet, index) :
        data = sheet.col_values(index)
        data_len = len(data)
        return (data_len, data)

    def read_xlsx_close(self,workbook) :
        workbook.release_resources()

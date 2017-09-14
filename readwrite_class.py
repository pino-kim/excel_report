import xlrd
import xlsxwriter

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


    def get_data_by_row_index(self, sheet, index) :
        data = sheet.row_values(index)
        data_len = len(data)
        return (data_len, data)


    def get_data_by_col_index(self, sheet, index) :
        data = sheet.col_values(index)
        data_len = len(data)
        return (data_len, data)




class Writexl :
    def __init__(self, workbook_titile):
        self.workbook_title = workbook_titile
        self.worksheet_page = 0

    def write_xlsx_workbook(self):
        workbook = xlrd.open_workbook( self.workbook_title, on_demand=True)
        return workbook

    def add_xlsx_worksheet(self, workbook, worksheet_title='Sheet') :
        worksheet_title = worksheet_title + str(self.worksheet_page)
        self.worksheet_page = self.worksheet_page + 1
        print(worksheet_title)



    def open_xlsx_workbook(self) :
        workbook = xlsxwriter.Workbook(self.workbook_title)
        return workbook


    def set_xlsx_worksheet(self, workbook):
        worksheet = workbook.add_worksheet()
        return worksheet


    def write_cell_by_point(self, worksheet, cell_point, content) :
        worksheet.write(cell_point, content)


    def write_cell_by_rowcal(self, worksheet, row, cal, content) :
        worksheet.write(row, cal, content)

    #def close_xlsx(self, workbook) :
    #    workbook.close()




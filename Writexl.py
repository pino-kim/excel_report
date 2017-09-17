import xlsxwriter

class Writexl :
    def __init__(self, workbook_titile):
        self.workbook_title = workbook_titile
        self.worksheet_page = 0

    def write_xlsx_workbook(self):
        workbook = xlsxwriter.Workbook( self.workbook_title)
        return workbook

    def add_xlsx_worksheet(self, workbook) :
        worksheet = workbook.add_worksheet()
        return worksheet

    def set_cell_format(self, workbook, type = 'defauilt'):
        if type == 'test' :
            format = workbook.add_format({'bold': True, 'font_color': 'red'})
            return format
        elif type == 'defauilt' :
            format = workbook.add_format({'bold': False, 'font_color': 'blue'})
            return format


    #def write_cell_w_merge

    def write_cell_by_cellname(self, worksheet, cell_name, content, format):
        worksheet.write(cell_name, content, format)

    def write_cell_by_rowcal(self, worksheet, row, cal, content) :
        worksheet.write(row, cal, content)

    def write_data_to_row(self, sheet, start_point, data):
        sheet.write_row(start_point, data)

    def write_data_to_col(self, sheet, start_point, data):
        sheet.write_column(start_point, data)


    def write_xlsx_close(self, workbook):
        workbook.close()
    #def close_xlsx(self, workbook) :
    #    workbook.close()




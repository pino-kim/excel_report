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
        if type == 'title':
            format = workbook.add_format({'bold' : True,
                                          'border' : 1,
                                          'bg_color':'silver',
                                          'font_color': 'red',
                                          'font_size': 15,
                                          'align': 'center', })
            return format

        elif type == 'label' :
            format = workbook.add_format({'bold': False,
                                          'border': 1,
                                          'bg_color':'yellow',
                                          'font_color': 'blue',
                                          'font_size': 12,
                                          'align':'center', })
            return format

        elif type == 'defauilt' :
            format = workbook.add_format({'bold': False,
                                          'border': 1,
                                          'font_color': 'black'})
            return format

    def set_row_size(self, worksheet, row_pos, row_size) :
        worksheet.set_row(row_pos, row_size)

    def set_col_size(self, worksheet, col_start, col_end, col_size, ):
        worksheet.set_column( col_start + ':' + col_end, col_size)

    #def write_cell_w_merge



    def write_cell_by_cellname(self, worksheet, cell_name, content, format):
        worksheet.write(cell_name, content, format)

    def write_cell_by_rowcal(self, worksheet, row, cal, content, format) :
        worksheet.write(row, cal, content, format)

    def write_data_to_row(self, sheet, start_point, data, format):
        sheet.write_row(start_point, data, format)

    def write_data_to_col(self, sheet, start_point, data, format):
        sheet.write_column(start_point, data, format)


    def write_xlsx_close(self, workbook):
        workbook.close()





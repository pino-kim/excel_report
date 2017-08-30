import xlrd
import xlsxwriter

#class Readxl:
#    workbook = None
#    worksheet = None

#    secret = "영구는 외계인이다."

#    def read_xlsx(self, file_name):
#        workbook = xlsxwriter.Workbook('filename.xlsx')

#    def

#    def write_cell(self, cell_point, value):

#    def sum(self, a, b):
#        result = a + b
#        print("%s님 %s + %s = %s입니다." % (self.name, a, b, result))


class Writexl :
    workbook_title = None
    worksheet_title = None

    def set_xlsx_title(self, workbook_title, worksheet_title):
        self.workbook_title = workbook_title
        self.worksheet_title = worksheet_title

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

    def close_xlsx(self, workbook) :
        workbook.close()




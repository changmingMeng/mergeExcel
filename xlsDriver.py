# encoding: utf-8
import xlrd
from xlutils.copy import copy

class ExcelDriver(object):


    #相比于old_getTargets，使用了with...as..语句来代替try...except语句
    def read_file_1(self,filename1):
        self.readfilename1 = filename1

    def read_file_2(self, filename2):
        self.readfilename2 = filename2

    def write_in_file2(self):
        with xlrd.open_workbook(self.readfilename2) as workbook2:
            file2sheet = workbook2.sheet_by_index(0)
            file2rown = file2sheet.nrows
            file2coln = file2sheet.ncols

        with xlrd.open_workbook(self.readfilename1) as workbook1:
            file1sheet = workbook1.sheet_by_index(0)
            file1rown = file1sheet.nrows
            file1coln = file1sheet.ncols

            writebook = copy(workbook2)
            writesheet = writebook.get_sheet(0)
            for r in range(file1rown):
                for c in range(file1coln):
                    print file1sheet.cell_value(r, c)
                    writesheet.write(file2rown+r, c, file1sheet.cell_value(r, c))

        writebook.save(self.readfilename2)

def copy_excel(f1, f2):
    ed = ExcelDriver()
    ed.read_file_1(f1)
    ed.read_file_2(f2)
    ed.write_in_file2()

if __name__ == "__main__":
    copy_excel(r"E:\developtools\PyCharm 2016.3\projects\mergeExcel\test1.xlsx",
               r"E:\developtools\PyCharm 2016.3\projects\mergeExcel\test2.xls")

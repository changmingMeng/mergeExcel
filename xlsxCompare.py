# encoding: utf-8

import xlrd
from openpyxl import Workbook
from openpyxl import load_workbook
from functools import wraps
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import utils

class ExcelCompare(object):


    def read_file_1(self,filename1):
        self.readfilename1 = filename1

    def read_file_2(self, filename2):
        self.readfilename2 = filename2

    def log_file(self, filename):
        self.logfile = filename
        f = open(self.logfile, 'r+')
        f.truncate()

    def logout(self,r,c,v1,v2):
        print (r,c,v1,v2)
        log = "row="+str(r+1)+", col="+str(c+1)+", v1="+str(v1).encode('utf-8')+", v2="+str(v2)+"\n\r"
        #log = "row={}, col={}, v1={}, v2={} \n".format(r,c,v1,v2)
        logfile = open(self.logfile, 'a')
        logfile.write(log)
        logfile.close()

    @staticmethod
    def value_equal(v1, v2):
        return abs(v1-v2)<0.1

    def compare(self):
        with xlrd.open_workbook(self.readfilename2) as workbook2:
            file2sheet = workbook2.sheet_by_index(0)
            file2rown = file2sheet.nrows
            file2coln = file2sheet.ncols

        with xlrd.open_workbook(self.readfilename1) as workbook1:
            file1sheet = workbook1.sheet_by_index(0)
            file1rown = file1sheet.nrows
            file1coln = file1sheet.ncols

        print "file1 rn=", file1rown, "cn=", file1coln
        print "file2 rn=", file2rown, "cn=", file2coln
        if not (file1coln == file2coln and file1rown == file2rown):
            return False

        isxlsxsame = True
        for r in range(file2rown):
            for c in range(file2coln):
                v1 = file1sheet.cell_value(r, c)
                v2 = file2sheet.cell_value(r, c)
                if c==0:
                    pass
                elif c in [6, 7, 8, 9, 10, 11]:
                    if not ( r== 0 or self.value_equal(float(v1), float(v2))):
                        self.logout(r, c, v1, v2)
                        isxlsxsame = False
                elif c in [1, 2]:
                    if not ( r==0 or self.value_equal(int(v1),int(v2))):
                        self.logout(r, c, v1, v2)
                        isxlsxsame = False
                else:
                    if not v1 == v2:
                        self.logout(r, c, v1, v2)
                        isxlsxsame = False
        return isxlsxsame








def compare_excel(f1, f2, f3):
    ed = ExcelCompare()
    ed.read_file_1(f1)
    ed.read_file_2(f2)
    ed.log_file(f3)
    if ed.compare():
        print "same"
    else:
        print "not same"

def test1():
    compare_excel(
        r"W网监控常用指标-20170306.xlsx".decode("utf-8").encode("GBK"),
        r"Wtest1.xlsx",
        r"comparelog.txt")


if __name__ == "__main__":
    test1()


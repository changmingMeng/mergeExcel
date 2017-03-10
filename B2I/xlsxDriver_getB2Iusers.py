# encoding: utf-8

import xlrd
import csv
from openpyxl import Workbook
from openpyxl import load_workbook
import chardet
import os
import re
import chardet

class B2IUserFile(object):


    def read_file(self, filename):
        self.readfilename = filename

    def write_file(self):
        userdict = {}
        namelist = ["大王卡",
                    "大小王产品（A）",
                    "大小王产品（B）",
                    "滴滴大王卡-58元套餐",
                    "滴滴大王卡-118元套餐",
                    "蚂蚁大宝卡",
                    "蚂蚁小宝卡"]
        with xlrd.open_workbook(self.readfilename) as workbook:
            filesheet = workbook.sheet_by_index(0)
            filerown = filesheet.nrows
            filecoln = filesheet.ncols

            wb = Workbook()
            ws = wb.active

            for r in range(1,filerown):
                product = filesheet.cell_value(r, 7).encode("utf-8")
                cellnumber = str(filesheet.cell_value(r, 1))

                if not product in userdict.keys():
                    userdict[product]=[]

                userdict[product].append(cellnumber)
        #按主要产品名称保存号码
        for product in userdict.keys():
            if product in namelist:
                filename = product+".txt"
                f = open(filename.decode("utf-8"), 'w')
                for num in userdict[product]:
                    f.write(num[:11]+'\n')
                f.close()
        #保存其他产品号码
        filename = "其他.txt"
        f = open(filename.decode("utf-8"), 'w')
        for product in userdict.keys():
            if product not in namelist:
                for num in userdict[product]:
                    f.write(num[:11] + '\n')
        f.close()

        #保存全部号码
        filename = "总体.txt"
        f = open(filename.decode("utf-8"), 'w')
        for product in userdict.keys():
            for num in userdict[product]:
                f.write(num[:11] + '\n')
        f.close()




def make_user_files(inputfile):
    ad = B2IUserFile()
    ad.read_file(inputfile)
    ad.write_file()

def test1():
    #os.walk(os.)
    for root, dirs, files in os.walk(os.getcwd()):
        for name in files:
            print name, chardet.detect(name), name.decode("GB2312")
            if not re.search("B2I项目充值数据", name.decode("GB2312").encode("utf-8")) == None:
                print "read file:", name
                make_user_files(name)



if __name__ == "__main__":
    test1()

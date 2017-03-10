# encoding: utf-8
import datetime
import Sqlite3Diver
import ExcelDriver
import utils
import os

class MainControl(object):


    def __init__(self):
        self.dbfile = os.path.abspath('.')+"\\datebase.txt" #':memory:'#
        self.tabledesc = ("exceldata", "name varchar(128), date date, erl float, data float")

        self.excel = ExcelDriver.ExcelDriver()

    def save2DB(self, r):
        #ut = utils.Utils()
        name, date, erl, data = r
        #date = ut.tuple2Sqlite3Timestring(old_date)
        self.sqlite.execDB("insert into exceldata(name, date, erl, data) "
                            "values ('"+name+"', DATE('"+date+"'), "+str(erl)+","+str(data)+")")

    def dropTable(self):
        self.sqlite = Sqlite3Diver.Sqlite3Driver(self.dbfile, self.tabledesc)
        self.sqlite.cerateDB()
        self.sqlite.execDB("drop table if exists exceldata")

    def closeDB(self):
        self.sqlite.closeDB()

    def saveData2DB(self, loadfilename, targetfilename):
        self.sqlite.cerateDB()

        self.excel.getTargets(targetfilename)
        self.excel.readFile(loadfilename)
        rows = self.excel.pickResult()

        for r in rows:
            #print r
            self.save2DB(r)

    def selectDB(self, tableName, colName, beginDate, endDate):
        select = "select name, avg(erl), sum(erl), avg(data), sum(data) " \
                 "from "+tableName+" "\
                 "where name='"+colName+"' and DATE(date) Between DATE('"+str(beginDate)+"') and DATE('"+str(endDate)+"')"
        #select = "select avg(erl), avg(data), sum(erl), sum(data) from " + tableName + " where name='" + colName + "'"
        #select = "select * from " + tableName + " where name='" + colName + "'"
        #select = "select * from "+tableName
        # print beginDate, endDate
        # print select
        return self.sqlite.getResult(select)

    def selectItemFromDB(self, beginDate, endDate):
        self.sqlite.cerateDB()
        self.results = []
        targets = self.excel.targets
        for t in targets:
            self.results.append(self.selectDB('exceldata', t, beginDate, endDate))

        # for r in self.results:
        #     print r

    def outputResults(self, begindate, enddate, outputfilename):
        rowNames = ('小区名称', '平均话务量', '话务量总量', '平均流量', '总流量' )
        self.excel.writeFile(self.results, rowNames, (begindate, enddate), outputfilename)

    def loadFile(self, loadfilename, targetfilename):
        self.dropTable()
        self.saveData2DB(loadfilename, targetfilename)

    def outputFile(self, begindate, enddate, outputfilename):
        self.selectItemFromDB(begindate, enddate)
        self.outputResults(begindate, enddate, outputfilename)

if __name__ == "__main__":
    mc = MainControl()
    #mc.dropTable()
    #mc.saveData2DB()
    mc.selectItemFromDB('2017-01-01', '2017-01-07')
    mc.outputResults()



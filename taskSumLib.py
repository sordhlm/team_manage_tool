#!/usr/bin/env python
#coding=utf-8
'''
Created on 2017年2月27日

@author: beike
'''

import os
import time,datetime
import re
import xlrd as xr
QA_MEM = {'Liangmin Huang','Hongyi Ai','Yingjun Yang','Beishi Chen','Shaobin Shu','Fang Wu','Chenchen Xu','Weili Han','Weifeng Guo','Zhilong Qian'}
def strConvert(inp):
    out = inp
    if isinstance(out,basestring):
        out = out.encode('utf8')
    else:
        out = unicode(out).encode('utf8')
    return out.lower()
def findFullName(name):
    for own_full in QA_MEM:
        m = re.search(name,own_full,re.IGNORECASE)
        if m:
            own = own_full
            break
        own = name
    print "[findFullName]src: %s, dst: %s"%(name,own)
    return own

class PersonTask(object):
    def __init__(self,owner,status,task,degree,comment=''):
        self.owner = owner
        self.status = status
        self.task = task
        self.degree = degree
        self.comment = comment
    def info_print(self):
        print '%s %s %s %s %s'%(self.owner.ljust(10),self.status.ljust(8),self.task.ljust(80),self.degree,self.comment)
class WeeklyTask(object):
    def __init__(self,date):
        self.date = date
        self.Ltask = []
        self.Perf = {}
        self.Num = {}
    def getOwner(self,ostr):
        olist = []
        olist_src = []
        owner_string = strConvert(ostr)
        olist_src = re.split('/|;',owner_string)
        for own in olist_src:
            olist.append(findFullName(own))
        return olist
    def addTask(self,owner,status,job,degree,comment=''):
        olist = self.getOwner(owner)
        num = len(olist)
        m = re.match("[0-9]+","%s"%degree)
        if not m:
            degree = 0
        degr = round(degree,2)/num
        for own in olist:
            task = PersonTask(strConvert(own),strConvert(status),strConvert(job),degr,strConvert(comment))
            self.Ltask.append(task)
    def listTask(self,owner,status):
        for task in self.Ltask:
            if (owner in task.owner ) and (task.status == status):
                task.info_print()
    def calPerf(self):
        for task in self.Ltask:
            if task.owner not in self.Perf.keys():
                self.Perf[task.owner] = 0
                self.Num[task.owner] = 0
            if task.status.strip() == 'done':
                if task.degree == 0:
                    task.info_print()
                self.Perf[task.owner] += task.degree
                self.Num[task.owner] += 1
class AllTask(object):
    def __init__(self):
        self.Lweek = []
        self.Perf = {}
        self.Num = {}
    def addWeek(self,weekly):
        self.Lweek.append(weekly)
    def listTaskForOwner(self,owner,status):
        for weekly in self.Lweek:
            #print weekly.date
            weekly.listTask(owner,status)
    def sumPerf(self,week):
        for owner in week.Perf:
            if owner not in self.Perf.keys():
                self.Perf[owner] = 0
                self.Num[owner] = 0
            self.Perf[owner] += week.Perf[owner]
            self.Num[owner] += week.Num[owner]
    def calPerf(self):
        for weekly in self.Lweek:
            weekly.calPerf()
            self.sumPerf(weekly)
        self.printPerf()
    def printPerf(self):
        for i in self.Perf:
            if self.Num[i] == 0:
                self.Num[i] = 1
            print "%10s  score:%6.2f  tasknum:%6.2f  perDegree:%6.2f"%(i,self.Perf[i],self.Num[i],self.Perf[i]/self.Num[i])



class taskSumLib(object):
    '''
    Function:
    1. read weekly xlsx file;
    2. calculate every tasks for everyone;
    3. output summary result;
    Input:
        src: weekly xlsx file
        dst: Target to generate. If file was exists, it would be regenerate.
    Output:
        None
    '''
    def __init__(self, src, dst):
#         Parameter Checking
        if not os.path.exists(src):
            print "Invalid Input File: %s"%src
            exit(-1)
#         Member variables setting
        self.src = src
        self.dst = dst
        self.task = AllTask()
#         1.Create or open a excle file;
        if not self.__openXls():
            exit(-1)
        else:
            self.__parseTask()

    def __openXls(self):
        if not os.path.isfile(self.src):
            raise StandardError("%s: open file failed!" % self.src)
        try:
            self.xlsx = xr.open_workbook(self.src)
            return True
        except Exception as e:
            print "Create a excle file failed: %s, %s"%(self.dst, e)
            return False
    def __parseTask(self):
        sheet = self.xlsx.sheets()[0]
        nrows = sheet.nrows
        ncols = sheet.ncols
        nodes = 0
        if ncols < 5:
            nodes = 1
        s_flag = 0
        e_flag = 0
        sub_flag = 0
        des = ''
        for i in range(1,nrows):
            #print "#####%s#######line:%d"%(sheet.cell(i,2),i)
            if sheet.cell(i,0).ctype == 3:
                if s_flag == 1:
                    self.task.addWeek(weekly)
                date = xr.xldate.xldate_as_datetime(sheet.cell(i,0).value,0)
                weekly = WeeklyTask(date)
                s_flag = 1
            task = sheet.cell(i,2).value
            status = sheet.cell(i,3).value
            degree = sheet.cell(i,4).value
            owner = sheet.cell(i,5).value
            if nodes != 1:
                des = sheet.cell(i,6).value
            m = re.match("^[a-zA-Z]+",task)
            if m:
                sub_flag = 0
                title = task
            else:
                sub_flag = 1

            m = re.match("\w+",owner)
            print '%20s %20s %s %s %s sub:%d'%(strConvert(owner),strConvert(status),strConvert(task),strConvert(degree),strConvert(des),sub_flag)
            if m:
                if sub_flag == 1:
                    task = title+" "+task
                weekly.addTask(owner,status,task,degree,des)
        self.task.addWeek(weekly)
        self.task.calPerf()


from optparse import OptionParser
if __name__ == '__main__':
    parser = OptionParser()
    parser.add_option("-m", "--mode", type="string", dest="mode", default="chk", help="work mode")
    parser.add_option("-n", "--name", type="string", dest="name", default="none", help="the name who need to check")
    (options, args) = parser.parse_args()
    name = options.name
    mode = options.mode
    #xlsx = "./../Workplan/QA_Weekly.xlsx"
    xlsx = "./weekly.xlsx"
    task = taskSumLib(xlsx, "issues_grb.xlsx")
    if mode == 'chk':
        task.task.listTaskForOwner(name,'done')



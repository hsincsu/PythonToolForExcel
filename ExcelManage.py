import pathlib
import os
import numpy as np
import xlwings as xw
import time
from decimal import *

class SheetManager(object):
    def __init__(self,wb,name,index):
        self.sheetname = name
        self.sheetindex = index
        self.IsInit = False
        self.wb = wb
        self.titlearry={}
    
    def init_sheet(self,sht,title_row):
        self.title_row = title_row
        self.sheet = sht
        self.nrows = self.sheet.used_range.last_cell.row
        self.ncols = self.sheet.used_range.last_cell.column
        for i in range(1,50):
            if self.sheet.range((title_row,i)).value is not None:
                self.titlearry[self.sheet.range((title_row,i)).value] = i
                print(self.sheet.range((title_row,i)).value + " : " + str(i))
            elif title_row > 1 and self.sheet.range((title_row-1,i)).value is not None:
                self.titlearry[self.sheet.range((title_row-1,i)).value] = i
                print(self.sheet.range((title_row-1,i)).value + " : " + str(i))
            else:
                pass
        self.IsInit = True
    
    def get_title_column(self,name):
        return self.titlearry[name]
    
    def get_range_rows_add_list(self,value,column):
        valuelist = self.sheet.range((1,column),(self.nrows,column)).value
        list_rows = []
        for i in range(self.nrows):
            if str(valuelist[i]).upper() == value.upper():
                list_rows.append(i+1)
        print(len(list_rows))
        return list_rows
    
    def get_range_rows_del_list(self,value,column,list_rows):
        valuelist = self.sheet.range((1,column),(self.nrows,column)).value
        list_work = []
        for i in list_rows:
            if valuelist[i-1] == value:
                list_work.append(i)
        print("rows:",len(list_work))
        return list_work



class ExcelManager(object):
    def __init__(self):
        self._filename = None
        self._FilePath = None
        self._IsOpen = False
        self.IsBusy = False
        self.wb = None
        self.app = None
        self.sheetarr={}
        

    def OpenFile(self):
        if self._IsOpen:
            return -1 #means error.
        else:
            self.app = xw.App(visible=False,add_book=False)
            self.app.display_alerts=False
            self.app.screen_updating=False
            self.wb = self.app.books.open(self._FilePath[0])
            if self.wb is not None:
                self.sheet_len = len(self.wb.sheets)
                for i in range(self.sheet_len):
                    self.sheetarr[self.wb.sheets[i].name] = SheetManager(self.wb,self.wb.sheets[i].name,i)
                self._IsOpen = True
                return 0
            else:
                return -2
           
            
    def CloseFile(self):
        if not self._IsOpen:
            return -1
        else:
            self.wb.close()
            self.app.quit()
            self.clear()
            return 0
    
    def clear(self):
        self._filename = None
        self._FilePath = None
        self._IsOpen = False
        self.IsBusy = False
        self.wb = None
        self.app = None

    def CreateFile(self,name):
        print("创建文件")

    def SetFilename(self,filename):
        self._filename = filename
    
    def SetFilePath(self,FilePath):
        self._FilePath = FilePath
    
    def SetIsOpen(self,IsOpen):
        self._IsOpen = IsOpen

    def GetEmployeeWithId(self,sheetname,idlist):
        if not self._IsOpen:
            return -1
        
        listid = idlist
        #get sheets
        sheet = self.wb.sheets[sheetname]
        sheetmanager = self.sheetarr[sheetname]

        if sheetmanager.IsInit is not True:
            sheetmanager.init_sheet(sheet,4) #actually should can be written by user.but I am lazy.

        column_cell = sheetmanager.titlearry['工号']
        row_cell = sheetmanager.title_row-1

        print((row_cell, column_cell))
        
        wb_new = self.app.books.add()
        sht_new = wb_new.sheets[0]
        sheet.range((1,1),(row_cell+1,sheetmanager.ncols)).api.Copy(sht_new.range('A1').api)
        row_new = row_cell+2
        # row_new += 1
        # #get id so we can check what we want.
        colid_list = sheet.range((row_new,column_cell),(sheetmanager.nrows,column_cell)).value

        for id in listid:
            for row2 in range(len(colid_list)):
                if int(id) == int(colid_list[row2]):
                    #storage it
                    sheet.api.Rows(row2+row_cell+2).Copy(sht_new.api.Rows(row_new))
                    row_new +=1
        
        wb_new.save(self._filename+'-pythoned.xlsx')
        wb_new.close()
        return 0
    
    def DeleteWhoIstheLucky(self,sheetname,idlist,month,projectid):
        if not self._IsOpen:
            return -1
        
        listid = idlist
        #get sheets
        sheet = self.wb.sheets[sheetname]
        sheetmanager = self.sheetarr[sheetname]

        if sheetmanager.IsInit is not True:
            sheetmanager.init_sheet(sheet,4) #actually should can be written by user.but I am lazy.

        #get month list
        projectlist = sheetmanager.get_range_rows_add_list(projectid,2) #projectid is 2
        worklist = sheetmanager.get_range_rows_del_list(month,1,projectlist)
        
        droplistid = []
        #del what we want
        for workid in listid:
            for id in worklist:
                if sheet.range((id,4)).value == workid:
                    droplistid.append(id)
                    break
    
        #add money to everyone who is lucky
        AllMoneyBefore = 0.00
        AllMoneyAfter = 0.00
        droplist = []
        #1.get all sum
    
        for id in worklist:
            if float(sheet.range((id,14)).value) != 1:
                AllMoneyBefore = AllMoneyBefore + sheet.range((id,9)).value
                AllMoneyAfter = AllMoneyAfter + sheet.range((id,13)).value   
            else:
                droplist.append(id)
        
        for id in droplist:
            worklist.remove(id)
        
        for id in droplistid:
            worklist.remove(id)
            AllMoneyBefore = AllMoneyBefore - sheet.range(id,9).value

        # print(len(worklist))
        AllMoneyBefore = round(AllMoneyBefore,2)
        AllMoneyAfter = round(AllMoneyAfter,2)
        getcontext().prec = 8
        ratio = Decimal(str(AllMoneyAfter)) / Decimal(str(AllMoneyBefore))
        ratio = float(ratio)
        print(AllMoneyAfter , AllMoneyBefore)
        print(ratio)

        for id in worklist:
            
            listvalue1 = sheet.range((id,6),(id,8)).value
            listvalue2 = sheet.range((id,10),(id,12)).value
            
            listvalue2[0] = round(listvalue1[0] * ratio,2)
            listvalue2[1] = round(listvalue1[1] * ratio,2)
            listvalue2[2] = round(listvalue1[2] * ratio,2)

            sheet.range((id,10),(id,12)).value = listvalue2

        idlast = sheetmanager.nrows + 10 #make sure no one is bigger.
        for id in droplistid: #every delete make id changed , need caculate.
            if idlast < id : #impossible that there is situation that equal.
                sheet.api.Rows(id - 1).Delete()
            else:
                sheet.api.Rows(id).Delete()
            idlast = id
            # print("deleted") 
        
        return 0

    
    def AddWhoIstheLucky(self,sheetname,idlist,month,projectid):
        if not self._IsOpen:
            return -1
        
        listid = idlist
        #get sheets
        sheet = self.wb.sheets[sheetname]
        sheetmanager = self.sheetarr[sheetname]

        if sheetmanager.IsInit is not True:
            sheetmanager.init_sheet(sheet,4) #actually should can be written by user.but I am lazy.

        #get month list
        projectlist = sheetmanager.get_range_rows_add_list(projectid,2) #projectid is 2
        worklist = sheetmanager.get_range_rows_del_list(month,1,projectlist)
        
        droplistid = []
        #del what we want
        for workid in listid:
            for id in worklist:
                if sheet.range((id,4)).value == workid:
                    droplistid.append(id)
                    break
    
        #add money to everyone who is lucky
        AllMoneyBefore = 0.00
        AllMoneyAfter = 0.00
        droplist = []
        #1.get all sum
    
        for id in worklist:
            if float(sheet.range((id,14)).value) != 1:
                AllMoneyBefore = AllMoneyBefore + sheet.range((id,9)).value
                AllMoneyAfter = AllMoneyAfter + sheet.range((id,13)).value   
            else:
                droplist.append(id)
        
        for id in droplist:
            worklist.remove(id)
        
        for id in droplistid:
            AllMoneyAfter = AllMoneyAfter - sheet.range(id,13).value

        # print(len(worklist))
        AllMoneyBefore = round(AllMoneyBefore,2)
        AllMoneyAfter = round(AllMoneyAfter,2)
        getcontext().prec = 8
        ratio = Decimal(str(AllMoneyAfter)) / Decimal(str(AllMoneyBefore))
        ratio = float(ratio)
        print(AllMoneyAfter , AllMoneyBefore)
        print(ratio)
        
        for id in worklist:
            
            listvalue1 = sheet.range((id,6),(id,8)).value
            listvalue2 = sheet.range((id,10),(id,12)).value
            
            listvalue2[0] = round(listvalue1[0] * ratio,2)
            listvalue2[1] = round(listvalue1[1] * ratio,2)
            listvalue2[2] = round(listvalue1[2] * ratio,2)

            print(listvalue2)
            sheet.range((id,10),(id,12)).value = listvalue2
        
        return 0

    
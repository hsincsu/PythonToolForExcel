import pathlib
import os
import numpy as np
import xlwings as xw
import time

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
        print(len(list_work))
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
        if not self._IsOpen:
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
        
        listid = idlist.split(' ')
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
        
        listid = idlist.split(' ')
        #get sheets
        sheet = self.wb.sheets[sheetname]
        sheetmanager = self.sheetarr[sheetname]

        if sheetmanager.IsInit is not True:
            sheetmanager.init_sheet(sheet,4) #actually should can be written by user.but I am lazy.

        #get month list
        # print("start get monthlist")
        projectlist = sheetmanager.get_range_rows_add_list(projectid,2) #projectid is 2
        worklist = sheetmanager.get_range_rows_del_list(month,1,projectlist)
        
        FromBeforeMoney = 0
        FromShebao = 0
        FromGongJiJin = 0
        ToBeforeMoney = 0
        ToShebao = 0
        ToGongJiJin = 0
        droplistid = []
        #get money 
        for workid in listid:
            for id in worklist:
                if sheet.range((id,4)).value == workid:
                    listvalue1 = sheet.range((id,6),(id,8)).value
                    listvalue2 = sheet.range((id,10),(id,12)).value
                    TmpFromBeforeMoney = float(listvalue1[0])
                    TmpFromShebao = float(listvalue1[1])
                    TmpFromGongJiJin = float(listvalue1[2])
                    TmpToBeforeMoney = float(listvalue2[0])
                    TmpToShebao = float(listvalue2[1])
                    TmpToGongJiJin = float(listvalue2[2])
                    if listvalue1[0] is not None and listvalue1[0] != 0:
                        FromBeforeMoney = FromBeforeMoney + listvalue1[0]
                    if listvalue1[1] is not None and listvalue1[1] != 0:
                        FromShebao = FromShebao + listvalue1[1]
                    if listvalue1[2] is not None and listvalue1[2] != 0:
                        FromGongJiJin = FromGongJiJin + listvalue1[2]
                    if listvalue2[0] is not None and listvalue2[0] != 0:
                        ToBeforeMoney = ToBeforeMoney + listvalue2[0]
                    if listvalue2[1] is not None and listvalue2[1] != 0:
                        ToShebao = ToShebao + listvalue2[1]
                    if listvalue2[2] is not None and listvalue2[2] != 0:
                        ToGongJiJin = ToGongJiJin + listvalue2[2]
                    droplistid.append(id)
                    break
        for id in droplistid:
            worklist.remove(id)
    
        # print(len(worklist))
        # print(FromBeforeMoney,FromShebao,FromGongJiJin,ToBeforeMoney,ToShebao,ToGongJiJin)
        #add money to everyone who is lucky
        AllFromBeforeMoney = 0.00
        AllFromShebao = 0.00
        AllFromGongJiJin = 0.00
        AllToBeforeMoney = 0.00
        AllToShebao = 0.00
        AllToGongJiJin = 0.00
        ratioFromBeforeMoney = {}
        ratioFromShebao = {}
        ratioFromGongJiJin = {}
        ratioToBeforeMoney = {}
        ratioToShebao = {}
        ratioToGongJiJin = {}
        droplist = []
        #1.get all sum
    
        for id in worklist:
            if float(sheet.range((id,14)).value) != 1:
                listvalue1 = sheet.range((id,6),(id,8)).value
                listvalue2 = sheet.range((id,10),(id,12)).value
                TmpFromBeforeMoney = float(listvalue1[0])
                TmpFromShebao = float(listvalue1[1])
                TmpFromGongJiJin = float(listvalue1[2])
                TmpToBeforeMoney = float(listvalue2[0])
                TmpToShebao = float(listvalue2[1])
                TmpToGongJiJin = float(listvalue2[2])

                if TmpFromBeforeMoney is not None and TmpFromBeforeMoney != 0:
                    AllFromBeforeMoney = AllFromBeforeMoney + TmpFromBeforeMoney
                if TmpFromShebao is not None and TmpFromShebao != 0:
                    AllFromShebao = AllFromShebao + TmpFromShebao
                if TmpFromGongJiJin is not None and TmpFromGongJiJin != 0:
                    AllFromGongJiJin = AllFromGongJiJin + TmpFromGongJiJin
                if TmpToBeforeMoney is not None and TmpToBeforeMoney != 0:
                    AllToBeforeMoney = AllToBeforeMoney + TmpToBeforeMoney
                if TmpToShebao is not None and TmpToShebao != 0:
                    AllToShebao = AllToShebao + TmpToShebao
                if TmpToGongJiJin is not None and TmpToGongJiJin != 0:
                    AllToGongJiJin = AllToGongJiJin + TmpToGongJiJin
                
            else:
                droplist.append(id)
        
        for id in droplist:
            worklist.remove(id)

        # print(len(worklist))
        AllFromBeforeMoney = round(AllFromBeforeMoney,2)
        AllFromShebao = round(AllFromShebao,2)
        AllFromGongJiJin = round(AllFromGongJiJin,2)
        AllToBeforeMoney = round(AllToBeforeMoney,2)
        AllToShebao = round(AllToShebao,2)
        AllToGongJiJin = round(AllToGongJiJin,2)

        TmpFromBeforeMoney = 0
        TmpFromShebao = 0
        TmpFromGongJiJin = 0
        TmpToBeforeMoney = 0
        TmpToShebao = 0
        TmpToGongJiJin = 0
        sum1 = 0
        sum2 = 0
        sum3 = 0
        sum4 = 0
        sum5 = 0
        sum6 = 0
        
        for id in worklist:
            
            listvalue1 = sheet.range((id,6),(id,8)).value
            listvalue2 = sheet.range((id,10),(id,12)).value
            TmpFromBeforeMoney = float(listvalue1[0])
            TmpFromShebao = float(listvalue1[1])
            TmpFromGongJiJin = float(listvalue1[2])
            TmpToBeforeMoney = float(listvalue2[0])
            TmpToShebao = float(listvalue2[1])
            TmpToGongJiJin = float(listvalue2[2])
            
            if id == worklist[-1]:
                listvalue1[0] = TmpFromBeforeMoney + FromBeforeMoney - sum1
                listvalue1[1] = TmpFromShebao + FromShebao - sum2
                listvalue1[2] = TmpFromGongJiJin + FromGongJiJin - sum3
                listvalue2[0] = TmpToBeforeMoney + ToBeforeMoney - sum4
                listvalue2[1] = TmpToShebao + ToShebao - sum5
                listvalue2[2] = TmpToGongJiJin + ToGongJiJin - sum6
            else:
                if TmpFromBeforeMoney is not None and TmpFromBeforeMoney != 0:
                    ratioFromBeforeMoney[id] = abs(TmpFromBeforeMoney / AllFromBeforeMoney)
                else:
                    ratioFromBeforeMoney[id] = 0.00
                if TmpFromShebao is not None and TmpFromShebao != 0:
                    ratioFromShebao[id] = abs(TmpFromShebao / AllFromShebao)
                else:
                    ratioFromShebao[id] = 0.00
                if TmpFromGongJiJin is not None and TmpFromGongJiJin != 0:
                    ratioFromGongJiJin[id] = abs(TmpFromGongJiJin / AllFromGongJiJin)
                else:
                    ratioFromGongJiJin[id] = 0.00
                if TmpToBeforeMoney is not None and TmpToBeforeMoney != 0:
                    ratioToBeforeMoney[id] = abs(TmpToBeforeMoney / AllToBeforeMoney)
                else:
                    ratioToBeforeMoney[id] = 0.00
                if TmpToShebao is not None and TmpToShebao != 0:
                    ratioToShebao[id] = abs(TmpToShebao / AllToShebao)
                else:
                    ratioToShebao[id] = 0.00
                if TmpToGongJiJin is not None and TmpToGongJiJin != 0:
                    ratioToGongJiJin[id] = abs(TmpToGongJiJin / AllToGongJiJin)
                else:
                    ratioToGongJiJin[id] = 0.00
                
                tmpvalue1 = round(ratioFromBeforeMoney[id] * FromBeforeMoney,2)
                listvalue1[0] = TmpFromBeforeMoney + tmpvalue1
                sum1 = sum1 + tmpvalue1
                tmpvalue2 = round(ratioFromShebao[id] * FromShebao,2)
                listvalue1[1] = TmpFromShebao + tmpvalue2
                sum2 = sum2 + tmpvalue2
                tmpvalue3 = round(ratioFromGongJiJin[id] * FromGongJiJin,2)
                listvalue1[2] = TmpFromGongJiJin + tmpvalue3
                sum3 = sum3 + tmpvalue3
                tmpvalue4 = round(ratioToBeforeMoney[id] * ToBeforeMoney,2)
                listvalue2[0] = TmpToBeforeMoney + tmpvalue4
                sum4 = sum4 + tmpvalue4
                tmpvalue5 = round(ratioToShebao[id] * ToShebao,2)
                listvalue2[1] = TmpToShebao + tmpvalue5
                sum5 = sum5 + tmpvalue5
                tmpvalue6 = round(ratioToGongJiJin[id] * ToGongJiJin,2)
                listvalue2[2] = TmpToGongJiJin + tmpvalue6
                sum6 = sum6 + tmpvalue6

            sheet.range((id,6),(id,8)).value = listvalue1
            sheet.range((id,10),(id,12)).value = listvalue2

            # print(sheet.range((id,6)).value, sheet.range((id,7)).value,sheet.range((id,8)).value,
            #         sheet.range((id,10)).value,sheet.range((id,11)).value,sheet.range((id,12)).value)

        idlast = sheetmanager.nrows + 1 #make sure no one is bigger.
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
        
        listid = idlist.split(' ')
        #get sheets
        sheet = self.wb.sheets[sheetname]
        sheetmanager = self.sheetarr[sheetname]

        if sheetmanager.IsInit is not True:
            sheetmanager.init_sheet(sheet,4) #actually should can be written by user.but I am lazy.

        #get month list
        # print("start get monthlist")
        projectlist = sheetmanager.get_range_rows_add_list(projectid,2) #projectid is 2
        worklist = sheetmanager.get_range_rows_del_list(month,1,projectlist)
        
        FromBeforeMoney = 0
        FromShebao = 0
        FromGongJiJin = 0
        ToBeforeMoney = 0
        ToShebao = 0
        ToGongJiJin = 0
        droplistid = []
        TmpFromBeforeMoney = 0
        TmpFromShebao = 0
        TmpFromGongJiJin = 0
        TmpToBeforeMoney = 0
        TmpToShebao = 0
        TmpToGongJiJin = 0
        #get money 
        for workid in listid:
            for id in worklist:
                if sheet.range((id,4)).value == workid:
                    listvalue1 = sheet.range((id,6),(id,8)).value
                    listvalue2 = sheet.range((id,10),(id,12)).value
                    TmpFromBeforeMoney = float(listvalue1[0])
                    TmpFromShebao = float(listvalue1[1])
                    TmpFromGongJiJin = float(listvalue1[2])
                    TmpToBeforeMoney = float(listvalue2[0])
                    TmpToShebao = float(listvalue2[1])
                    TmpToGongJiJin = float(listvalue2[2])
                    if listvalue1[0] is not None and listvalue1[0] != 0:
                        FromBeforeMoney = FromBeforeMoney + listvalue1[0]
                    if listvalue1[1] is not None and listvalue1[1] != 0:
                        FromShebao = FromShebao + listvalue1[1]
                    if listvalue1[2] is not None and listvalue1[2] != 0:
                        FromGongJiJin = FromGongJiJin + listvalue1[2]
                    if listvalue2[0] is not None and listvalue2[0] != 0:
                        ToBeforeMoney = ToBeforeMoney + listvalue2[0]
                    if listvalue2[1] is not None and listvalue2[1] != 0:
                        ToShebao = ToShebao + listvalue2[1]
                    if listvalue2[2] is not None and listvalue2[2] != 0:
                        ToGongJiJin = ToGongJiJin + listvalue2[2]
                    droplistid.append(id)
                    break
        for id in droplistid:
            worklist.remove(id)
    
        # print(len(worklist))
        # print(FromBeforeMoney,FromShebao,FromGongJiJin,ToBeforeMoney,ToShebao,ToGongJiJin)
        #add money to everyone who is lucky
        AllFromBeforeMoney = 0.00
        AllFromShebao = 0.00
        AllFromGongJiJin = 0.00
        AllToBeforeMoney = 0.00
        AllToShebao = 0.00
        AllToGongJiJin = 0.00
        TmpFromBeforeMoney = 0
        TmpFromShebao = 0
        TmpFromGongJiJin = 0
        TmpToBeforeMoney = 0
        TmpToShebao = 0
        TmpToGongJiJin = 0
        ratioFromBeforeMoney = {}
        ratioFromShebao = {}
        ratioFromGongJiJin = {}
        ratioToBeforeMoney = {}
        ratioToShebao = {}
        ratioToGongJiJin = {}
        droplist = []
        #1.get all sum
    
        for id in worklist:
            if float(sheet.range((id,14)).value) != 1:
                listvalue1 = sheet.range((id,6),(id,8)).value
                listvalue2 = sheet.range((id,10),(id,12)).value
                TmpFromBeforeMoney = float(listvalue1[0])
                TmpFromShebao = float(listvalue1[1])
                TmpFromGongJiJin = float(listvalue1[2])
                TmpToBeforeMoney = float(listvalue2[0])
                TmpToShebao = float(listvalue2[1])
                TmpToGongJiJin = float(listvalue2[2])

                if TmpFromBeforeMoney is not None and TmpFromBeforeMoney != 0:
                    AllFromBeforeMoney = AllFromBeforeMoney + TmpFromBeforeMoney
                if TmpFromShebao is not None and TmpFromShebao != 0:
                    AllFromShebao = AllFromShebao + TmpFromShebao
                if TmpFromGongJiJin is not None and TmpFromGongJiJin != 0:
                    AllFromGongJiJin = AllFromGongJiJin + TmpFromGongJiJin
                if TmpToBeforeMoney is not None and TmpToBeforeMoney != 0:
                    AllToBeforeMoney = AllToBeforeMoney + TmpToBeforeMoney
                if TmpToShebao is not None and TmpToShebao != 0:
                    AllToShebao = AllToShebao + TmpToShebao
                if TmpToGongJiJin is not None and TmpToGongJiJin != 0:
                    AllToGongJiJin = AllToGongJiJin + TmpToGongJiJin
                
            else:
                droplist.append(id)
        
        for id in droplist:
            worklist.remove(id)

        # print(len(worklist))
        AllFromBeforeMoney = round(AllFromBeforeMoney,2)
        AllFromShebao = round(AllFromShebao,2)
        AllFromGongJiJin = round(AllFromGongJiJin,2)
        AllToBeforeMoney = round(AllToBeforeMoney,2)
        AllToShebao = round(AllToShebao,2)
        AllToGongJiJin = round(AllToGongJiJin,2)

        TmpFromBeforeMoney = 0
        TmpFromShebao = 0
        TmpFromGongJiJin = 0
        TmpToBeforeMoney = 0
        TmpToShebao = 0
        TmpToGongJiJin = 0
        sum1 = 0
        sum2 = 0
        sum3 = 0
        sum4 = 0
        sum5 = 0
        sum6 = 0
        
        for id in worklist:
            
            listvalue1 = sheet.range((id,6),(id,8)).value
            listvalue2 = sheet.range((id,10),(id,12)).value
            TmpFromBeforeMoney = float(listvalue1[0])
            TmpFromShebao = float(listvalue1[1])
            TmpFromGongJiJin = float(listvalue1[2])
            TmpToBeforeMoney = float(listvalue2[0])
            TmpToShebao = float(listvalue2[1])
            TmpToGongJiJin = float(listvalue2[2])
            
            if id == worklist[-1]:
                listvalue1[0] = TmpFromBeforeMoney - FromBeforeMoney + sum1
                listvalue1[1] = TmpFromShebao - FromShebao + sum2
                listvalue1[2] = TmpFromGongJiJin - FromGongJiJin + sum3
                listvalue2[0] = TmpToBeforeMoney - ToBeforeMoney + sum4
                listvalue2[1] = TmpToShebao - ToShebao + sum5
                listvalue2[2] = TmpToGongJiJin - ToGongJiJin + sum6
            else:
                if TmpFromBeforeMoney is not None and TmpFromBeforeMoney != 0:
                    ratioFromBeforeMoney[id] = abs(TmpFromBeforeMoney / AllFromBeforeMoney)
                else:
                    ratioFromBeforeMoney[id] = 0.00
                if TmpFromShebao is not None and TmpFromShebao != 0:
                    ratioFromShebao[id] = abs(TmpFromShebao / AllFromShebao)
                else:
                    ratioFromShebao[id] = 0.00
                if TmpFromGongJiJin is not None and TmpFromGongJiJin != 0:
                    ratioFromGongJiJin[id] = abs(TmpFromGongJiJin / AllFromGongJiJin)
                else:
                    ratioFromGongJiJin[id] = 0.00
                if TmpToBeforeMoney is not None and TmpToBeforeMoney != 0:
                    ratioToBeforeMoney[id] = abs(TmpToBeforeMoney / AllToBeforeMoney)
                else:
                    ratioToBeforeMoney[id] = 0.00
                if TmpToShebao is not None and TmpToShebao != 0:
                    ratioToShebao[id] = abs(TmpToShebao / AllToShebao)
                else:
                    ratioToShebao[id] = 0.00
                if TmpToGongJiJin is not None and TmpToGongJiJin != 0:
                    ratioToGongJiJin[id] = abs(TmpToGongJiJin / AllToGongJiJin)
                else:
                    ratioToGongJiJin[id] = 0.00
                
                tmpvalue1 = round(ratioFromBeforeMoney[id] * FromBeforeMoney,2)
                listvalue1[0] = TmpFromBeforeMoney - tmpvalue1
                sum1 = sum1 + tmpvalue1
                tmpvalue2 = round(ratioFromShebao[id] * FromShebao,2)
                listvalue1[1] = TmpFromShebao - tmpvalue2
                sum2 = sum2 + tmpvalue2
                tmpvalue3 = round(ratioFromGongJiJin[id] * FromGongJiJin,2)
                listvalue1[2] = TmpFromGongJiJin - tmpvalue3
                sum3 = sum3 + tmpvalue3
                tmpvalue4 = round(ratioToBeforeMoney[id] * ToBeforeMoney,2)
                listvalue2[0] = TmpToBeforeMoney - tmpvalue4
                sum4 = sum4 + tmpvalue4
                tmpvalue5 = round(ratioToShebao[id] * ToShebao,2)
                listvalue2[1] = TmpToShebao - tmpvalue5
                sum5 = sum5 + tmpvalue5
                tmpvalue6 = round(ratioToGongJiJin[id] * ToGongJiJin,2)
                listvalue2[2] = TmpToGongJiJin - tmpvalue6
                sum6 = sum6 + tmpvalue6

            sheet.range((id,6),(id,8)).value = listvalue1
            sheet.range((id,10),(id,12)).value = listvalue2

            # print(sheet.range((id,6)).value, sheet.range((id,7)).value,sheet.range((id,8)).value,
            #         sheet.range((id,10)).value,sheet.range((id,11)).value,sheet.range((id,12)).value)

        idlast = sheetmanager.nrows + 1 #make sure no one is bigger.
        for id in droplistid: #every delete make id changed , need caculate.
            if idlast < id : #impossible that there is situation that equal.
                sheet.api.Rows(id - 1).Delete()
            else:
                sheet.api.Rows(id).Delete()
            idlast = id
            # print("deleted") 
        
        return 0

    
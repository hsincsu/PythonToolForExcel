# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'D:\AppList\python3.9\Lib\site-packages\qt5_applications\Qt\bin\2022_12_12_excel_check_template.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from ExcelManage import ExcelManager
from Ui_Form import Ui_Form
from Ui_Delete import Ui_Delete
from PyQt5 import QtCore, QtGui, QtWidgets,Qt
import pathlib

class Ui_MainWindow(QtWidgets.QMainWindow):
    def setupUi(self):
        self.setObjectName("MainWindow")
        self.resize(805, 617)
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 30, 251, 141))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.textBrowser = QtWidgets.QTextBrowser(self.gridLayoutWidget)
        self.textBrowser.setObjectName("textBrowser")
        self.gridLayout.addWidget(self.textBrowser, 0, 0, 1, 1)
        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QtCore.QRect(270, 10, 521, 591))
        self.groupBox.setObjectName("groupBox")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.groupBox)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(0, 20, 521, 221))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.scrollArea = QtWidgets.QScrollArea(self.verticalLayoutWidget)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName("scrollArea")
        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 517, 217))
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")
        self.textBrowser_2 = QtWidgets.QTextBrowser(self.scrollAreaWidgetContents)
        self.textBrowser_2.setGeometry(QtCore.QRect(0, 0, 521, 221))
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.verticalLayout.addWidget(self.scrollArea)
        self.verticalLayoutWidget_2 = QtWidgets.QWidget(self.groupBox)
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(0, 250, 521, 331))
        self.verticalLayoutWidget_2.setObjectName("verticalLayoutWidget_2")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_2)
        self.verticalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.calendarWidget = QtWidgets.QCalendarWidget(self.verticalLayoutWidget_2)
        self.calendarWidget.setObjectName("calendarWidget")
        self.verticalLayout_4.addWidget(self.calendarWidget)
        self.groupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_2.setGeometry(QtCore.QRect(10, 10, 261, 591))
        self.groupBox_2.setObjectName("groupBox_2")
        self.tableWidget = QtWidgets.QTableWidget(self.groupBox_2)
        self.tableWidget.setGeometry(QtCore.QRect(0, 170, 251, 411))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.label_2 = QtWidgets.QLabel(self.groupBox_2)
        self.label_2.setGeometry(QtCore.QRect(10, 240, 131, 16))
        self.label_2.setObjectName("label_2")
        self.pushButton_3 = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_3.setGeometry(QtCore.QRect(10, 260, 111, 31))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_4 = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_4.setGeometry(QtCore.QRect(130, 260, 111, 31))
        self.pushButton_4.setMouseTracking(False)
        self.pushButton_4.setTabletTracking(False)
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_5 = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_5.setGeometry(QtCore.QRect(10, 300, 111, 31))
        self.pushButton_5.setObjectName("pushButton_5")
        self.pushButton_6 = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_6.setGeometry(QtCore.QRect(130, 300, 111, 31))
        self.pushButton_6.setMouseTracking(False)
        self.pushButton_6.setTabletTracking(False)
        self.pushButton_6.setObjectName("pushButton_6")
        self.pushButton_7 = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_7.setGeometry(QtCore.QRect(10, 340, 111, 31))
        self.pushButton_7.setObjectName("pushButton_7")

        self.label = QtWidgets.QLabel(self.groupBox_2)
        self.label.setGeometry(QtCore.QRect(10, 180, 131, 16))
        self.label.setObjectName("label")
        self.pushButton_2 = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_2.setGeometry(QtCore.QRect(10, 200, 111, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton.setGeometry(QtCore.QRect(130, 200, 111, 31))
        self.pushButton.setMouseTracking(False)
        self.pushButton.setTabletTracking(False)
        self.pushButton.setObjectName("pushButton")
        self.tableWidget.raise_()
        self.label.raise_()
        self.pushButton_2.raise_()
        self.pushButton.raise_()
        self.label_2.raise_()
        self.pushButton_3.raise_()
        self.pushButton_4.raise_()
        self.pushButton_5.raise_()
        self.pushButton_6.raise_()
        self.pushButton_7.raise_()
        self.groupBox_2.raise_()
        self.groupBox.raise_()
        self.gridLayoutWidget.raise_()
        self.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)
        

        self.retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self)
        self.pushButton_2.clicked.connect(self.showOpenFileDialog)
        self.pushButton.clicked.connect(self.showCloseFileDialog)
        self.pushButton_3.clicked.connect(self.CheckEmployeeWithId)
        self.pushButton_4.clicked.connect(self.DeleteWhoIsLucky)
        self.pushButton_5.clicked.connect(self.savefile)
        self.pushButton_6.clicked.connect(self.AddWhoIsLucky)
        self.pushButton_7.clicked.connect(self.BatchProcess)
        #init excel 
        self.excelmanage = ExcelManager()
        self.cwd = '.'
        self.listdeleted = []
        self.listadded = []

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "EXCEL小工具-思泥专用版"))
        self.groupBox.setTitle(_translate("MainWindow", "控制台信息"))
        self.groupBox_2.setTitle(_translate("MainWindow", "输入区"))
        self.label_2.setText(_translate("MainWindow", "命令："))
        self.pushButton_3.setText(_translate("MainWindow", "查找-员工编号"))
        self.pushButton_4.setText(_translate("MainWindow", "分散-员工工资"))
        self.pushButton_5.setText(_translate("MainWindow", "保存"))
        self.pushButton_6.setText(_translate("MainWindow", "添加-员工工资"))
        self.pushButton_7.setText(_translate("MainWindow", "批量操作"))
        self.label.setText(_translate("MainWindow", "文件操作："))
        self.pushButton_2.setText(_translate("MainWindow", "  打开excel文件   "))
        self.pushButton.setText(_translate("MainWindow", "关闭excel文件"))
    
    def savefile(self):
        if self.excelmanage._IsOpen:
            self.excelmanage.wb.save()
            QtWidgets.QMessageBox.warning(self, "信息", "保存成功", QtWidgets.QMessageBox.Ok)
            return
        else:
            QtWidgets.QMessageBox.warning(self, "警告", "请先打开待处理的excel文件", QtWidgets.QMessageBox.Ok)
            return
    
    def showOpenFileDialog(self):
        if not self.excelmanage._IsOpen:
            fname = QtWidgets.QFileDialog.getOpenFileName(self, '打开文件',self.cwd,"*.xlsx;*.xls")
        else:
            QtWidgets.QMessageBox.warning(self, "警告", "请关闭已打开的excel文件", QtWidgets.QMessageBox.Cancel)
            return
        if fname[0]:
            path = pathlib.Path(fname[0])
            filename = path.name
            if filename != "":
                self.textBrowser.setText("文件名："+filename)
                self.setConsoleInfo("打开了文件"+filename)
                self.cwd = fname[0]
                self.excelmanage.SetFilePath(fname)
                self.excelmanage.SetFilename(filename)
                ret = self.excelmanage.OpenFile()
                if ret != 0:
                    QtWidgets.QMessageBox.warning(self, "错误", "无法打开文件,请选择正确的excel文件", QtWidgets.QMessageBox.Cancel)
                    self.excelmanage.clear()
    
    def setConsoleInfo(self,text):
        self.textBrowser_2.append(text)

    def setTextInfo(self,text,option=0):
        if option == 0:
            self.textBrowser.append(text)
        else:
            self.textBrowser.setText(text)
    
    def showCloseFileDialog(self):
        if self.excelmanage._IsOpen:
            reply = QtWidgets.QMessageBox.question(self,'Warning','是否保存?',QtWidgets.QMessageBox.Yes,QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.Yes:
                self.excelmanage.wb.save()
                self.textBrowser.clear()
                self.textBrowser_2.clear()
                self.excelmanage.CloseFile()
            elif reply == QtWidgets.QMessageBox.No:
                self.textBrowser.clear()
                self.textBrowser_2.clear()
                self.excelmanage.CloseFile()
            else:
                return
        else:
            self.textBrowser.clear()
            self.textBrowser_2.clear()
        
    
    def CheckEmployeeWithId(self):
        if not self.excelmanage._IsOpen:
            QtWidgets.QMessageBox.warning(self, "警告", "请先打开待处理的excel文件", QtWidgets.QMessageBox.Ok)
            return
        else:
            if self.excelmanage.IsBusy:
                QtWidgets.QMessageBox.warning(self, "警告", "忙等待...", QtWidgets.QMessageBox.Ok)
                return
            self.excelmanage.IsBusy=True
            self.setConsoleInfo("处理文件: "+ self.excelmanage._filename+"......")
            self.infolist = Ui_Form()
            self.infolist.setWindowModality(QtCore.Qt.ApplicationModal)
            self.infolist.exec_()

            if self.infolist.sheetname == '' and self.infolist.idlist =='':
                self.excelmanage.IsBusy=False
                return
            self.setTextInfo("----操作-查找-员工编号----")
            self.setTextInfo("sheet名称: " + self.infolist.sheetname)
            self.setTextInfo("员工编号: " + self.infolist.idlist)
            QtWidgets.QApplication.processEvents()
            self.infolist.idlist = self.infolist.idlist.split(' ')
            ret = self.excelmanage.GetEmployeeWithId(self.infolist.sheetname,self.infolist.idlist)
            if ret != 0:
                QtWidgets.QMessageBox.warning(self, "错误", "处理发生错误", QtWidgets.QMessageBox.Ok)
                self.infolist = None
                self.excelmanage.IsBusy=False
            self.excelmanage.IsBusy=False
            # self.GetEmployeeWithId()
            self.setConsoleInfo("文件: "+ self.excelmanage._filename+"处理成功")
    
    def DeleteWhoIsLucky(self):
        if not self.excelmanage._IsOpen:
            QtWidgets.QMessageBox.warning(self, "警告", "请先打开待处理的excel文件", QtWidgets.QMessageBox.Ok)
            return
        else:
            if self.excelmanage.IsBusy:
                QtWidgets.QMessageBox.warning(self, "警告", "忙等待...", QtWidgets.QMessageBox.Ok)
                return
            self.excelmanage.IsBusy=True
            self.setConsoleInfo("处理文件: "+ self.excelmanage._filename+"......")
            self.infolist = Ui_Delete()
            self.infolist.setWindowModality(QtCore.Qt.ApplicationModal)
            self.infolist.exec_()

            if self.infolist.sheetname == '' or self.infolist.idlist =='' \
                or self.infolist.month =='' or self.infolist.projectid == '':
                self.excelmanage.IsBusy=False
                return

            for deleted in self.listdeleted:
                if deleted == str(self.infolist.month+self.infolist.projectid):
                    QtWidgets.QMessageBox.warning(self, "警告", deleted+"已处理", QtWidgets.QMessageBox.Ok)
                    return
                
            self.setTextInfo("----操作-分散-员工工资----",1)
            self.setTextInfo("sheet名称: " + self.infolist.sheetname)
            self.setTextInfo("员工编号: " + self.infolist.idlist)
            self.setTextInfo("月份：" + self.infolist.month)
            self.setTextInfo("项目编号：" + self.infolist.projectid)
            QtWidgets.QApplication.processEvents()
            self.infolist.idlist = self.infolist.idlist.split(' ')
            ret = self.excelmanage.DeleteWhoIstheLucky(self.infolist.sheetname,self.infolist.idlist,
                                                    self.infolist.month,self.infolist.projectid)
            if ret != 0:
                QtWidgets.QMessageBox.warning(self, "错误", "处理发生错误", QtWidgets.QMessageBox.Ok)
                self.infolist = None
                self.excelmanage.IsBusy=False

            self.listdeleted.append(str(self.infolist.month+self.infolist.projectid))
            self.setConsoleInfo("已处理"+str(self.infolist.month+self.infolist.projectid))
            self.excelmanage.IsBusy=False
            # self.GetEmployeeWithId()
            self.setConsoleInfo("文件: "+ self.excelmanage._filename+"...处理成功")
        
    def AddWhoIsLucky(self):
        if not self.excelmanage._IsOpen:
            QtWidgets.QMessageBox.warning(self, "警告", "请先打开待处理的excel文件", QtWidgets.QMessageBox.Ok)
            return
        else:
            if self.excelmanage.IsBusy:
                QtWidgets.QMessageBox.warning(self, "警告", "忙等待...", QtWidgets.QMessageBox.Ok)
                return
            self.excelmanage.IsBusy=True
            self.setConsoleInfo("处理文件: "+ self.excelmanage._filename+"......")
            self.infolist = Ui_Delete()
            self.infolist.setWindowModality(QtCore.Qt.ApplicationModal)
            self.infolist.exec_()

            if self.infolist.sheetname == '' or self.infolist.idlist =='' \
                or self.infolist.month =='' or self.infolist.projectid == '':
                self.excelmanage.IsBusy=False
                return

            for added in self.listadded:
                if added == str(self.infolist.month+self.infolist.projectid):
                    QtWidgets.QMessageBox.warning(self, "警告", added+"已处理", QtWidgets.QMessageBox.Ok)
                    return
                
            self.setTextInfo("----操作-添加-员工工资----",1)
            self.setTextInfo("sheet名称: " + self.infolist.sheetname)
            self.setTextInfo("员工编号: " + self.infolist.idlist)
            self.setTextInfo("月份：" + self.infolist.month)
            self.setTextInfo("项目编号：" + self.infolist.projectid)
            QtWidgets.QApplication.processEvents()
            self.infolist.idlist = self.infolist.idlist.split(' ')
            ret = self.excelmanage.AddWhoIstheLucky(self.infolist.sheetname,self.infolist.idlist,
                                                    self.infolist.month,self.infolist.projectid)
            if ret != 0:
                QtWidgets.QMessageBox.warning(self, "错误", "处理发生错误", QtWidgets.QMessageBox.Ok)
                self.infolist = None
                self.excelmanage.IsBusy=False

            self.listadded.append(str(self.infolist.month+self.infolist.projectid))
            self.setConsoleInfo("已处理"+str(self.infolist.month+self.infolist.projectid))
            self.excelmanage.IsBusy=False
            # self.GetEmployeeWithId()
            self.setConsoleInfo("文件: "+ self.excelmanage._filename+"...处理成功")

    def BatchProcess(self):
        if self.excelmanage._IsOpen:
            fname = QtWidgets.QFileDialog.getOpenFileName(self, '打开文件',self.cwd,"*.txt;")
        else:
            QtWidgets.QMessageBox.warning(self, "警告", "请先打开的excel文件", QtWidgets.QMessageBox.Cancel)
            return
        if fname[0]:
            path = pathlib.Path(fname[0])
            file = open(fname[0],encoding='utf-8')
            filename = path.name
            if filename != "":
                if self.excelmanage.IsBusy:
                    QtWidgets.QMessageBox.warning(self, "警告", "忙等待...", QtWidgets.QMessageBox.Ok)
                    return
                self.excelmanage.IsBusy=True
                self.textBrowser.setText("文件名："+filename)
                self.setConsoleInfo("打开了批处理配置文件"+filename)
                self.setConsoleInfo("正在进行批处理: \n")
                QtWidgets.QApplication.processEvents()
                for line in file.readlines():
                    list = line.strip().split(' ')
                    operation = list[0] #operation "Add  Del"
                    self.setConsoleInfo("opeartion: " + operation)
                    list.remove(operation)
                    sheetname = list[0] #sheet
                    self.setConsoleInfo("sheetname: " + sheetname)
                    list.remove(sheetname)
                    month = list[0] #month
                    self.setConsoleInfo("month: " + month)
                    list.remove(month)
                    projectid = list[0] #project id
                    self.setConsoleInfo("projectid: " + projectid)
                    list.remove(projectid)
                    listid = list
                    self.setConsoleInfo("listid: ")
                    self.setConsoleInfo(str(listid))
                    QtWidgets.QApplication.processEvents()
                    ret = 0
                    if operation == 'Add':
                        for added in self.listadded:
                            if added == str(month+projectid):
                                QtWidgets.QMessageBox.warning(self, "警告", added+":已处理", QtWidgets.QMessageBox.Ok)
                                ret = -1
                                break
                        if ret == -1:
                            continue
                        ret = self.excelmanage.AddWhoIstheLucky(sheetname,listid,month,projectid)
                    elif operation == 'Del':
                        for deleted in self.listdeleted:
                            if deleted == str(month+projectid):
                                QtWidgets.QMessageBox.warning(self, "警告", deleted+":已处理", QtWidgets.QMessageBox.Ok)
                                ret = -1
                                break
                        if ret == -1:
                            continue
                        ret = self.excelmanage.DeleteWhoIstheLucky(sheetname,listid,month,projectid)
                    else:
                        QtWidgets.QMessageBox.warning(self, "错误", "operation错误", QtWidgets.QMessageBox.Ok)

                    if ret != 0:
                        QtWidgets.QMessageBox.warning(self, "错误", "处理发生错误", QtWidgets.QMessageBox.Ok)
                        self.excelmanage.IsBusy=False
                        return 
                    
                    if operation == 'Add':
                        self.listadded.append(str(month+projectid))
                    elif operation == 'Del':
                        self.listdeleted.append(str(month+projectid))
                    else:
                        QtWidgets.QMessageBox.warning(self, "错误", "operation错误", QtWidgets.QMessageBox.Ok)
                        self.setConsoleInfo("处理失败")
                        self.excelmanage.IsBusy=False
                        return
        self.setConsoleInfo("处理成功")
        QtWidgets.QApplication.processEvents()
        return 0

    def closeEvent(self,event):
        if self.excelmanage._IsOpen:
            reply = QtWidgets.QMessageBox.question(self,'Warning','退出并保存或者取消并直接退出?',QtWidgets.QMessageBox.Yes,QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.Yes:
                self.excelmanage.wb.save()
                event.accept()
            else:
                event.accept()
        else:
            event.accept()


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ui = Ui_MainWindow()
    ui.setupUi()
    ui.show()
    sys.exit(app.exec_())


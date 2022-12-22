from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Delete(QtWidgets.QDialog):
    def __init__(self):
        super().__init__()
        self.sheetname = ''
        self.idlist = ''
        self.month = ''
        self.projectid = ''
        self.setupUi()

    def setupUi(self):
        self.setObjectName("Dialog")
        self.resize(527, 427)
        self.listView = QtWidgets.QListView(self)
        self.listView.setGeometry(QtCore.QRect(0, 0, 531, 431))
        self.listView.setAutoFillBackground(False)
        self.listView.setObjectName("listView")
        self.label = QtWidgets.QLabel(self)
        self.label.setGeometry(QtCore.QRect(18, 20, 101, 31))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.lineEdit = QtWidgets.QLineEdit(self)
        self.lineEdit.setGeometry(QtCore.QRect(120, 20, 391, 31))
        self.lineEdit.setObjectName("lineEdit")
        self.label_2 = QtWidgets.QLabel(self)
        self.label_2.setGeometry(QtCore.QRect(20, 70, 181, 31))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.pushButton = QtWidgets.QPushButton(self)
        self.pushButton.setGeometry(QtCore.QRect(30, 330, 101, 31))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self)
        self.pushButton_2.setGeometry(QtCore.QRect(380, 330, 101, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.textEdit = QtWidgets.QTextEdit(self)
        self.textEdit.setGeometry(QtCore.QRect(20, 100, 491, 81))
        self.textEdit.setObjectName("textEdit")
        self.label_3 = QtWidgets.QLabel(self)
        self.label_3.setGeometry(QtCore.QRect(30, 200, 121, 21))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.lineEdit_2 = QtWidgets.QLineEdit(self)
        self.lineEdit_2.setGeometry(QtCore.QRect(160, 190, 221, 31))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_3 = QtWidgets.QLineEdit(self)
        self.lineEdit_3.setGeometry(QtCore.QRect(160, 240, 221, 31))
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.label_4 = QtWidgets.QLabel(self)
        self.label_4.setGeometry(QtCore.QRect(30, 250, 121, 21))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")

        self.retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self)
        self.pushButton.clicked.connect(self.ProcessOk)
        self.pushButton_2.clicked.connect(self.ProcessCancel)

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label.setText(_translate("Dialog", "sheet名称："))
        self.label_2.setText(_translate("Dialog", "员工编号（空格隔开）："))
        self.pushButton.setText(_translate("Dialog", "确认"))
        self.pushButton_2.setText(_translate("Dialog", "取消"))
        self.label_3.setText(_translate("Dialog", "月份（x月）："))
        self.label_4.setText(_translate("Dialog", "  项目编号 ："))

    def ProcessOk(self):
        self.lineEdit.setText("附件1-人员人工费用台账")
        self.sheetname = self.lineEdit.text()
        self.idlist = self.textEdit.toPlainText()
        self.month = self.lineEdit_2.text() + "月"
        self.projectid = self.lineEdit_3.text().strip()
        strinfo = "sheet: "+self.sheetname + "\n" + "员工编号:" + self.idlist+"\n"+"月份："+self.month \
                  + "\n" + "项目编号："+self.projectid
        QtWidgets.QMessageBox.warning(self, "信息", strinfo, QtWidgets.QMessageBox.Ok)
        self.close()
    
    def ProcessCancel(self):
        self.close()
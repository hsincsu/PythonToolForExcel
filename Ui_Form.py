from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(QtWidgets.QDialog):
    def __init__(self):
        super().__init__()
        self.sheetname = ''
        self.idlist = ''
        self.setupUi()

    def setupUi(self):
        self.setObjectName("Dialog")
        self.resize(420, 340)
        self.listView = QtWidgets.QListView(self)
        self.listView.setGeometry(QtCore.QRect(0, 0, 421, 341))
        self.listView.setAutoFillBackground(False)
        self.listView.setObjectName("listView")
        self.label = QtWidgets.QLabel(self)
        self.label.setGeometry(QtCore.QRect(20, 40, 101, 31))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.lineEdit = QtWidgets.QLineEdit(self)
        self.lineEdit.setGeometry(QtCore.QRect(122, 40, 271, 31))
        self.lineEdit.setObjectName("lineEdit")
        self.label_2 = QtWidgets.QLabel(self)
        self.label_2.setGeometry(QtCore.QRect(20, 90, 181, 31))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.pushButton = QtWidgets.QPushButton(self)
        self.pushButton.setGeometry(QtCore.QRect(30, 290, 101, 31))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self)
        self.pushButton_2.setGeometry(QtCore.QRect(280, 290, 101, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.textEdit = QtWidgets.QTextEdit(self)
        self.textEdit.setGeometry(QtCore.QRect(20, 120, 381, 161))
        self.textEdit.setObjectName("textEdit")

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
    
    def ProcessOk(self):
        self.sheetname = self.lineEdit.text()
        self.idlist = self.textEdit.toPlainText()
        strinfo = "sheet: "+self.sheetname + "\n" + "员工编号" + self.idlist
        QtWidgets.QMessageBox.warning(self, "信息", strinfo, QtWidgets.QMessageBox.Ok)
        self.close()
    
    def ProcessCancel(self):
        self.close()

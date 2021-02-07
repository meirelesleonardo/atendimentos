# coding=utf-8
import sys
import os
from PyQt5.QtWidgets import  QMainWindow, QApplication, QTableWidgetItem, QFileDialog, QMessageBox, QComboBox
from PyQt5 import uic
from PyQt5.QtCore import  QDate, QDateTime, Qt
import datetime
#import openpyxl
#import EditTipo
import ConectarSqlite
#import WEditar
#from openpyxl.styles import Alignment
from PyQt5 import QtCore, QtGui, QtWidgets



# from openpyxl.styles import Border, Side


class Agenda(QMainWindow):

    def __init__(self):
        QMainWindow.__init__(self)
        # Carregando e configurando Formulário
        #uic.loadUi("./GSU_Agenda.ui", self)
        ########## copiado ############
        self.setObjectName("gsu")
        self.setEnabled(True)
        self.resize(1462, 691)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("L:/Projetos/Python/GSU2/Ico.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(icon)
        self.setAutoFillBackground(False)
        self.setStyleSheet("background-image: url(:/newPrefix/logo.jpg);")
        self.layout = QtWidgets.QWidget(self)
        self.layout.setObjectName("layout")
        self.gridLayout_4 = QtWidgets.QGridLayout(self.layout)
        self.gridLayout_4.setHorizontalSpacing(6)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.tabWidget = QtWidgets.QTabWidget(self.layout)
        self.tabWidget.setEnabled(True)
        self.tabWidget.setAcceptDrops(False)
        self.tabWidget.setWhatsThis("")
        self.tabWidget.setAccessibleName("")
        self.tabWidget.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.tabWidget.setStyleSheet("background-image: url(:/newPrefix/logo.jpg);\n"
"image: url(:/newPrefix/logo.jpg);\n"
"border-image: url(:/newPrefix/logo.jpg);")
        self.tabWidget.setTabsClosable(False)
        self.tabWidget.setMovable(False)
        self.tabWidget.setTabBarAutoHide(True)
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.tab.setObjectName("tab")
        self.gridLayout_7 = QtWidgets.QGridLayout(self.tab)
        self.gridLayout_7.setHorizontalSpacing(6)
        self.gridLayout_7.setObjectName("gridLayout_7")
        spacerItem = QtWidgets.QSpacerItem(467, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_7.addItem(spacerItem, 0, 1, 1, 1)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_4 = QtWidgets.QLabel(self.tab)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_2.addWidget(self.label_4)
        self.lineLook = QtWidgets.QLineEdit(self.tab)
        self.lineLook.setObjectName("lineLook")
        self.horizontalLayout_2.addWidget(self.lineLook)
        self.gridLayout_7.addLayout(self.horizontalLayout_2, 0, 0, 1, 1)
        spacerItem1 = QtWidgets.QSpacerItem(988, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_7.addItem(spacerItem1, 2, 0, 1, 1)
        spacerItem2 = QtWidgets.QSpacerItem(468, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_7.addItem(spacerItem2, 0, 2, 1, 1)
        spacerItem3 = QtWidgets.QSpacerItem(468, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_7.addItem(spacerItem3, 2, 2, 1, 1)
        spacerItem4 = QtWidgets.QSpacerItem(468, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_7.addItem(spacerItem4, 1, 2, 1, 1)
        spacerItem5 = QtWidgets.QSpacerItem(988, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_7.addItem(spacerItem5, 1, 0, 1, 1)
        self.tableTel = QtWidgets.QTableWidget(self.tab)
        self.tableTel.setObjectName("tableTel")
        self.tableTel.setColumnCount(7)
        self.tableTel.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.tableTel.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTel.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTel.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTel.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTel.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTel.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTel.setHorizontalHeaderItem(6, item)
        self.gridLayout_7.addWidget(self.tableTel, 5, 0, 1, 3)
        self.pushButton_3 = QtWidgets.QPushButton(self.tab)
        self.pushButton_3.setEnabled(False)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setObjectName("pushButton_3")
        self.gridLayout_7.addWidget(self.pushButton_3, 3, 0, 1, 2)
        self.btAddContato = QtWidgets.QPushButton(self.tab)
        self.btAddContato.setEnabled(True)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.btAddContato.setFont(font)
        self.btAddContato.setObjectName("btAddContato")
        self.gridLayout_7.addWidget(self.btAddContato, 4, 0, 1, 2)
        self.gridLayout_7.setColumnStretch(0, 3)
        self.tabWidget.addTab(self.tab, "")
        self.gridLayout_4.addWidget(self.tabWidget, 0, 0, 1, 1)
        self.setCentralWidget(self.layout)
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)

        self.retranslateUi(self)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(self)

    def retranslateUi(self, gsu):
        _translate = QtCore.QCoreApplication.translate
        gsu.setWindowTitle(_translate("gsu", "DDT_GSU    Agenda"))
        self.label_4.setText(_translate("gsu", "Localizar"))
        item = self.tableTel.horizontalHeaderItem(0)
        item.setText(_translate("gsu", "ID"))
        item = self.tableTel.horizontalHeaderItem(1)
        item.setText(_translate("gsu", "Sigla"))
        item = self.tableTel.horizontalHeaderItem(2)
        item.setText(_translate("gsu", "Unidade"))
        item = self.tableTel.horizontalHeaderItem(3)
        item.setText(_translate("gsu", "Setor"))
        item = self.tableTel.horizontalHeaderItem(4)
        item.setText(_translate("gsu", "Tel. Fixo"))
        item = self.tableTel.horizontalHeaderItem(5)
        item.setText(_translate("gsu", "Celular"))
        item = self.tableTel.horizontalHeaderItem(6)
        item.setText(_translate("gsu", "Local"))
        self.pushButton_3.setText(_translate("gsu", "Editar contato selecionado"))
        self.btAddContato.setText(_translate("gsu", "Adicionar Contato"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("gsu", "Lista Telefônica"))


        ########## Fim ############

        self.btAddContato.clicked.connect(self.addContato)

        self.tableTel.setColumnWidth(0, 0)
        self.tableTel.setColumnWidth(1, 300)
        self.tableTel.setColumnWidth(2, 330)
        self.tableTel.setColumnWidth(3, 220)
        self.tableTel.setColumnWidth(4, 150)
        self.tableTel.setColumnWidth(5, 150)
        self.tableTel.setColumnWidth(6, 300)

        Agenda.carregarTel(self)

        self.lineLook.textChanged.connect(self.locTel)

    def addContato(self):
        self.addContato = AddContato()
        self.addContato.lineEdit_setor.clear()
        self.addContato.lineEdit_TelFixo.clear()
        self.addContato.lineEdit_Celular.clear()

        self.addContato.show()

    def carregarTel(self):
        sql = ("SELECT a.id, a.unidade, a.descricao, b.setor, b.telFixo, b.telCel, a.localizacao "
               "FROM unidades a INNER JOIN listatel b on b.unidade = a.id ORDER BY a.unidade;")
        # print (sql2)
        result = (ConectarSqlite.conectarBd(sql))
        self.tableTel.setRowCount(0)
        # print (result)
        for row, form in enumerate(result):
            self.tableTel.insertRow(row)
            for column, item in enumerate(form):
                self.tableTel.setItem(row, column, QTableWidgetItem(str(item)))

        #print(result)


    def locTel(self):
        self.tableTel.setRowCount(0)
        texto = self.lineLook.text()

        if texto == "":
            self.carregarTel()
        else:

            sql = ("SELECT a.id, a.unidade, a.descricao, b.setor, b.telFixo, b.telCel, a.localizacao "
                   "FROM unidades a INNER JOIN listatel b on b.unidade = a.id ORDER BY a.unidade;")
            # print (sql2)
            result = (ConectarSqlite.conectarBd(sql))
            texto = texto.upper()
            row = 0
            #
            for cont, item in enumerate(result):

                if texto in item[1] or texto in item[2]:
                    self.tableTel.insertRow(row)
                    # print (item)
                    # self.tableTel.insertRow(1)
                    for column, item2 in enumerate(item):
                        self.tableTel.setItem(row, column, QTableWidgetItem(str(item2)))
                    row = row + 1
                else:
                    pass




class AddContato(QMainWindow):

    def __init__(self):
        QMainWindow.__init__(self)
        # Carregando e configurando Formulário
        #uic.loadUi("./addContato.ui", self)
        ########## copiado ############
        self.setObjectName("MainWindow")
        self.resize(699, 289)
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.gridLayout = QtWidgets.QGridLayout(self.frame)
        self.gridLayout.setObjectName("gridLayout")
        self.UNIDADE = QtWidgets.QLabel(self.frame)
        self.UNIDADE.setObjectName("UNIDADE")
        self.gridLayout.addWidget(self.UNIDADE, 2, 0, 1, 1)
        self.bt_addUnidade = QtWidgets.QPushButton(self.frame)
        self.bt_addUnidade.setObjectName("bt_addUnidade")
        self.gridLayout.addWidget(self.bt_addUnidade, 3, 7, 1, 1)
        self.label = QtWidgets.QLabel(self.frame)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 4, 0, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.frame)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 5, 0, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.frame)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 6, 0, 1, 1)
        spacerItem = QtWidgets.QSpacerItem(125, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem, 7, 0, 1, 1)
        spacerItem1 = QtWidgets.QSpacerItem(124, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem1, 7, 3, 1, 2)
        self.bt_addContato = QtWidgets.QPushButton(self.frame)
        self.bt_addContato.setObjectName("bt_addContato")
        self.gridLayout.addWidget(self.bt_addContato, 7, 5, 1, 1)
        spacerItem2 = QtWidgets.QSpacerItem(281, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem2, 7, 6, 1, 1)
        spacerItem3 = QtWidgets.QSpacerItem(118, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem3, 7, 7, 1, 1)
        spacerItem4 = QtWidgets.QSpacerItem(20, 33, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem4, 0, 1, 1, 2)
        self.label_4 = QtWidgets.QLabel(self.frame)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.gridLayout.addWidget(self.label_4, 1, 0, 1, 1)
        self.PesqUnidade = QtWidgets.QLineEdit(self.frame)
        self.PesqUnidade.setObjectName("PesqUnidade")
        self.gridLayout.addWidget(self.PesqUnidade, 1, 2, 1, 6)
        self.cbUnidade = QtWidgets.QComboBox(self.frame)
        self.cbUnidade.setObjectName("cbUnidade")
        self.gridLayout.addWidget(self.cbUnidade, 3, 0, 1, 7)
        self.lineEdit_setor = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_setor.setObjectName("lineEdit_setor")
        self.gridLayout.addWidget(self.lineEdit_setor, 4, 1, 1, 7)
        self.lineEdit_TelFixo = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_TelFixo.setObjectName("lineEdit_TelFixo")
        self.gridLayout.addWidget(self.lineEdit_TelFixo, 5, 1, 1, 7)
        self.lineEdit_Celular = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_Celular.setObjectName("lineEdit_Celular")
        self.gridLayout.addWidget(self.lineEdit_Celular, 6, 1, 1, 7)
        self.gridLayout_2.addWidget(self.frame, 0, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(self)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 699, 21))
        self.menubar.setObjectName("menubar")
        self.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)

        self.retranslateUi(self)
        QtCore.QMetaObject.connectSlotsByName(self)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Add Contato"))
        self.UNIDADE.setText(_translate("MainWindow", "UNIDADE:"))
        self.bt_addUnidade.setText(_translate("MainWindow", "Adicionar nova Unidade"))
        self.label.setText(_translate("MainWindow", "SETOR"))
        self.label_2.setText(_translate("MainWindow", "TELEFONE FIXO:"))
        self.label_3.setText(_translate("MainWindow", "CELULAR"))
        self.bt_addContato.setText(_translate("MainWindow", "SALVAR NOVO CONTATO"))
        self.label_4.setText(_translate("MainWindow", "Localizar Unidade"))

        ########## Fim ############
        self.PesqUnidade.textChanged.connect(self.locUnidade)
        self.bt_addContato.clicked.connect(self.addNovoContato)
        self.bt_addUnidade.clicked.connect(self.addUnidade)
        AddContato.atualizaCbUnidde(self)


    def addNovoContato(self):
        texto = self.cbUnidade.currentText()
        sql1 = ('SELECT id FROM unidades WHERE  unidade = "'+str(texto)+'";')
        result1 = ConectarSqlite.conectarBd(sql1)
        unidade = result1[0][0]
        setor = self.lineEdit_setor.text()
        telFixo = self.lineEdit_TelFixo.text()
        telCel = self.lineEdit_Celular.text()

        sql2 = ("INSERT INTO listatel VALUES ('"+str(unidade)+"', '"+setor+"', '"+telFixo+"', '"+telCel+"');")
        #print(sql2)
        ConectarSqlite.conectarBd(sql2)
        addContato.close()

        QMessageBox.about(self, 'Informação', 'Novo contato cadastrado com sucesso!')

    def addUnidade(self):
        self.close()
        self.addUnidade = AddUnidade()
        self.addUnidade.lineEdit_unidade.clear()
        self.addUnidade.lineEdit_descricao.clear()
        self.addUnidade.lineEdit_local.clear()

        self.addUnidade.show()

    def locUnidade(self):

        sql2 = ("SELECT id, unidade FROM unidades ORDER BY unidade;")
        result2 = (ConectarSqlite.conectarBd(sql2))
        texto = self.PesqUnidade.text()
        texto = texto.upper()
        n = 0

        for cont, item in enumerate(result2):
            if  item[1].count(texto) > 0:
                self.cbUnidade.setCurrentIndex(cont)
                return

    def atualizaCbUnidde(self):
        self.cbUnidade.clear()
        sql2 = ("SELECT id, unidade FROM unidades ORDER BY unidade;")
        result2 = (ConectarSqlite.conectarBd(sql2))
        for cont, item in enumerate(result2):
            self.cbUnidade.insertItem(cont,str(item[1]))

class AddUnidade(QMainWindow):

    def __init__(self):
        QMainWindow.__init__(self)
        # Carregando e configurando Formulário
        #uic.loadUi("./addUnidade.ui", self)
########## copiado ############
        self.setObjectName("MainWindow")
        self.resize(800, 271)
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.frame)
        self.gridLayout_2.setObjectName("gridLayout_2")
        spacerItem = QtWidgets.QSpacerItem(523, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_2.addItem(spacerItem, 3, 4, 1, 1)
        self.bt_addUnidade = QtWidgets.QPushButton(self.frame)
        self.bt_addUnidade.setObjectName("bt_addUnidade")
        self.gridLayout_2.addWidget(self.bt_addUnidade, 3, 3, 1, 1)
        self.label_4 = QtWidgets.QLabel(self.frame)
        self.label_4.setObjectName("label_4")
        self.gridLayout_2.addWidget(self.label_4, 2, 0, 1, 3)
        self.lineEdit_local = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_local.setObjectName("lineEdit_local")
        self.gridLayout_2.addWidget(self.lineEdit_local, 2, 3, 1, 2)
        self.lineEdit_descricao = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_descricao.setObjectName("lineEdit_descricao")
        self.gridLayout_2.addWidget(self.lineEdit_descricao, 1, 2, 1, 3)
        self.lineEdit_unidade = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_unidade.setObjectName("lineEdit_unidade")
        self.gridLayout_2.addWidget(self.lineEdit_unidade, 0, 2, 1, 3)
        self.label_3 = QtWidgets.QLabel(self.frame)
        self.label_3.setObjectName("label_3")
        self.gridLayout_2.addWidget(self.label_3, 1, 0, 1, 2)
        self.label = QtWidgets.QLabel(self.frame)
        self.label.setObjectName("label")
        self.gridLayout_2.addWidget(self.label, 0, 0, 1, 2)
        self.gridLayout.addWidget(self.frame, 0, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(self)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 21))
        self.menubar.setObjectName("menubar")
        self.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)

        self.retranslateUi(self)
        QtCore.QMetaObject.connectSlotsByName(self)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Cadastrar Nova Unidade"))
        self.bt_addUnidade.setText(_translate("MainWindow", "SALVAR NOVA UNIDADE"))
        self.label_4.setText(_translate("MainWindow", "Local / Endepreço:"))
        self.label_3.setText(_translate("MainWindow", "Descrição"))
        self.label.setText(_translate("MainWindow", "Nova Unidade"))

############ Fim ############
        self.bt_addUnidade.clicked.connect(self.addNovaUnidade)

    def addNovaUnidade(self):
        sql1 = ("SELECT MAX(id) FROM unidades;")
        result = ConectarSqlite.conectarBd(sql1)
        id = result[0][0] + 1
        unidade = self.lineEdit_unidade.text()
        descricao = self.lineEdit_descricao.text()
        local = self.lineEdit_local.text()
        unidade = unidade.upper()
        descricao = descricao.upper()
        local = local.upper()

        sql2 = ("INSERT INTO unidades (id, unidade, descricao, localizacao) \
        VALUES ("+str(id)+", '"+unidade+"', '"+descricao+"', '"+local+"');")

        ConectarSqlite.conectarBd(sql2)
        self.addUnidade.close()
        agenda.close()

        QMessageBox.about(self, 'Informação', 'Nova unidade cadastrada com sucesso!')

if __name__ == "__main__":
    app = QApplication(sys.argv)
    agenda = Agenda()
    #addContato = AddContato()
    #addUnidade = AddUnidade()
    agenda.show()
    app.exec_()

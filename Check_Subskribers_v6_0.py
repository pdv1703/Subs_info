import sys
import cx_Oracle
import os
import threading
import datetime
from PyQt5.QtCore import Qt, QDateTime, QDate, QDir
from PyQt5.QtWidgets import (
    QWidget, QLabel, QLineEdit, QTextEdit, QGridLayout, QApplication,
    QPushButton, QMainWindow, QApplication, QScrollArea, QTableWidget,
    QTableWidgetItem, QSplitter, QHBoxLayout, QTabWidget, QSizePolicy,
    QComboBox, QDateEdit, QInputDialog, QErrorMessage, QMessageBox)

#Данные для авторизации которые считываються с файла
os.environ['NLS_LANG'] = 'UKRAINIAN_UKRAINE.CL8MSWIN1251'
file = open('Authorization informations.txt')
data_from_file = file.readlines()
MCPS_Authorization_Data = str(data_from_file[1])
DBM_Authorization_Data = str(data_from_file[4])
print(DBM_Authorization_Data)
file.close()
#класс для просмотра большого текста в таблицах
class BigTextFromTable(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):

        self.TextFromTable = QTextEdit()
        self.CloseTextFromTableButton = QPushButton('Ok')

        grid = QGridLayout()
        grid.setSpacing(10)
        grid.addWidget(self.TextFromTable, 0, 0, 2, 2)
        grid.addWidget(self.CloseTextFromTableButton, 2, 1)
        self.setLayout(grid)

        self.setGeometry(300, 300, 400, 200)
        self.setWindowTitle('Редактор больших данных')
        self.show()

#класс для окна удаления статусов
class DeletedStatyses(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        self.DeletAllStatyses = QPushButton('&Удалить все похожие статусы')
        self.DeletOneStatys = QPushButton('&Удалить только такой статус')
        self.StatysText = QLineEdit()
        self.StatysText.setText("Source info")

        grid = QGridLayout()
        grid.setSpacing(10)
        grid.addWidget(self.StatysText, 0, 0, 2, 2)
        grid.addWidget(self.DeletAllStatyses, 1, 0)
        grid.addWidget(self.DeletOneStatys, 1, 1)
        self.setLayout(grid)

        self.setGeometry(300, 300, 400, 90)
        self.setWindowTitle('Удаление статусов')
        self.show()


class PrimaryWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        self.bottomLeftTabWidget = QTabWidget()
        self.bottomLeftTabWidget.setSizePolicy(QSizePolicy.Preferred,
                                               QSizePolicy.Ignored)

        #формирование первой вкдадки-------------------
        tab1 = QWidget()
        #кнопка вывода результата с БД МСПС
        self.MCPSGoButton = QPushButton(
            '&Показать статусы и заявки по абоненту:')
        self.MCPSGoButton.setFixedHeight(30)

        #кнопка удаления статусов
        self.DelStatyses = QPushButton('&Удалить статусы')
        self.DelStatyses.setFixedWidth(400)

        #поле ввода юникс даты
        self.ConvertIn = QLineEdit()
        self.ConvertIn.setText("1490572799000")
        #self.ConvertIn.setMaxLength(13)
        self.ConvertIn.setInputMask('D000000000000')

        #поле вывода юникс даты
        self.ConvertOut = QLineEdit()
        self.ConvertOut.setFixedHeight(30)

        #кнопка конвертации с unix date в обычный формат
        self.ConvertButton = QPushButton(
            '&Преобразование UNIX даты в обычный DATE')
        self.ConvertButton.setFixedHeight(30)

        #поле ввода номера
        self.InfoEdit = QLineEdit()
        self.InfoEdit.setText("678450249")
        self.InfoEdit.setMaxLength(9)
        self.InfoEdit.setInputMask('D00000000')

        #поле вывода статусов и COS_Namе с mcps.spsdb_pps_subscriber
        self.StatysVievTitle = QLabel('Статусы по абоненту в статусе:')
        self.StatysVievTitle.setFixedHeight(30)
        self.COSVievTitle = QPushButton(
            'Проверка актуальности статусов по ТП:')
        self.COSVievTitle.setFixedWidth(400)
        self.StatysVievStatys = QComboBox()
        self.StatysVievStatys.setFixedWidth(30)
        self.StatysVievStatys.addItem('1')
        self.StatysVievStatys.addItem('2')

        self.tableWidgetForStatys = QTableWidget()
        #self.tableWidgetForStatys.setGridStyle(5)
        self.tableWidgetForStatys.setAlternatingRowColors(1)
        self.tableWidgetForStatys.setRowCount(1)
        self.tableWidgetForStatys.setColumnCount(6)


        #поле вывода параметров заявок
        self.MCPSOrdersParamButton = QPushButton(
            '&Показать параметры заявки №:')
        self.MCPSOrdersParamButton.setFixedHeight(30)
        self.OrdersNumberEdit = QLineEdit()
        self.OrdersNumberEdit.setText("939693858")
        self.tableWidgetForOrdersParam = QTableWidget()
        self.tableWidgetForOrdersParam.setRowCount(1)
        self.tableWidgetForOrdersParam.setColumnCount(2)
        self.tableWidgetForOrdersParam.setAlternatingRowColors(1)

        #поле вывода заявок по абоненту
        self.OrdersVievTitle = QLabel('Заявки по абоненту:')
        self.OrdersVievTitle.setFixedHeight(20)
        self.tableWidgetForOrders = QTableWidget()
        self.tableWidgetForOrders.setRowCount(1)
        self.tableWidgetForOrders.setColumnCount(17)
        self.tableWidgetForOrders.setAlternatingRowColors(1)




        #поле вывода активности за последниие дни
        self.DBMGoButton = QPushButton('&Активность в промежуток времени:')
        self.DBMGoButton.setFixedHeight(30)

        self.FromDataEdit = QLabel('From:')
        self.FromDataEdit.setFixedWidth(100)

        self.ToDataEdit = QLabel('To:')
        self.ToDataEdit.setFixedWidth(50)
        self.ToDataEdit.setFixedHeight(30)

        #поле с датой "c"
        self.FromDate = QDateEdit()

        #установка даты From по умолчанию
        day = datetime.datetime.today().day
        #day = day - 7
        month = datetime.datetime.today().month
        year = datetime.datetime.today().year
        someDT = datetime.datetime(year, month, day)
        self.FromDate.setDate(someDT)

        #поле с датой "до"
        self.ToDate = QDateEdit()
        self.ToDate.setDate(QDate.currentDate())

        #таблица для вывода результата запроса
        self.tableWidgetForActive = QTableWidget()
        self.tableWidgetForActive.setRowCount(1)
        self.tableWidgetForActive.setColumnCount(59)
        self.tableWidgetForActive.setAlternatingRowColors(1)

        #определение стиля сплитторов

        # stylesheet = "QSplitter::handle{background: #DCDCDC}"
        #сплиттер статусов и вывода текущего ТП
        self.splitter30 = QSplitter(Qt.Horizontal)
        self.splitter30.addWidget(self.COSVievTitle)
        self.splitter30.addWidget(self.StatysVievTitle)

        self.splitter1 = QSplitter(Qt.Horizontal)
        self.splitter1.addWidget(self.splitter30)
        self.splitter1.addWidget(self.StatysVievStatys)


        self.splitter20 = QSplitter(Qt.Vertical)
        self.splitter20.addWidget(self.splitter1)
        self.splitter20.addWidget(self.tableWidgetForStatys)


        #сплиттер заявок
        self.splitter3 = QSplitter(Qt.Vertical)
        self.splitter3.addWidget(self.OrdersVievTitle)
        self.splitter3.addWidget(self.tableWidgetForOrders)

        #сплиттер активности за последнее время
        self.splitter22 = QSplitter(Qt.Horizontal)
        self.splitter22.addWidget(self.FromDataEdit)
        self.splitter22.addWidget(self.FromDate)

        self.splitter23 = QSplitter(Qt.Horizontal)
        self.splitter23.addWidget(self.ToDataEdit)
        self.splitter23.addWidget(self.ToDate)

        self.splitter24 = QSplitter(Qt.Horizontal)
        self.splitter24.addWidget(self.splitter22)
        self.splitter24.addWidget(self.splitter23)

        self.splitter21 = QSplitter(Qt.Vertical)
        self.splitter21.addWidget(self.DBMGoButton)
        self.splitter21.addWidget(self.splitter24)

        self.splitter4 = QSplitter(Qt.Vertical)
        self.splitter4.addWidget(self.splitter21)
        self.splitter4.addWidget(self.tableWidgetForActive)

        #сплиттер статусов
        self.splitter5 = QSplitter(Qt.Horizontal)
        self.splitter5.addWidget(self.MCPSOrdersParamButton)
        self.splitter5.addWidget(self.OrdersNumberEdit)

        self.splitter6 = QSplitter(Qt.Vertical)
        self.splitter6.addWidget(self.splitter5)
        self.splitter6.addWidget(self.tableWidgetForOrdersParam)

        #сплиттер конвертации даты
        self.splitter17 = QSplitter(Qt.Horizontal)
        self.splitter17.addWidget(self.ConvertButton)
        self.splitter17.addWidget(self.ConvertIn)

        self.splitter18 = QSplitter(Qt.Vertical)
        self.splitter18.addWidget(self.splitter17)
        self.splitter18.addWidget(self.ConvertOut)

        #сплиттер между конвертацией даты и параметрами заявки
        self.splitter19 = QSplitter(Qt.Vertical)
        self.splitter19.addWidget(self.splitter6)
        self.splitter19.addWidget(self.splitter18)

        #сплиттер между параметрами заявок\конвертация даты и активностью
        self.splitter7 = QSplitter(Qt.Vertical)
        self.splitter7.addWidget(self.splitter19)
        self.splitter7.addWidget(self.splitter4)

        self.splitter7.setObjectName('splitter7')
        self.splitter7.setStyleSheet("""QSplitter#splitter7::handle {background: #DCDCDC; border-style: outset; border-width: 1px; border-color: #828282; border-radius: 3px;}""")

        #сплиттер между статусами и заявками#######
        self.splitter8 = QSplitter(Qt.Vertical)
        self.splitter8.addWidget(self.splitter20)
        self.splitter8.addWidget(self.splitter3)

        self.splitter8.setObjectName('splitter8')
        self.splitter8.setStyleSheet("""QSplitter#splitter8::handle {background: #DCDCDC; border-style: outset; border-width: 1px; border-color: #828282; border-radius: 3px;}""")

        #общий сплиттер всех окон
        self.splitter9 = QSplitter(Qt.Horizontal)
        self.splitter9.addWidget(self.splitter8)
        self.splitter9.addWidget(self.splitter7)

        self.splitter9.setObjectName('splitter9')
        self.splitter9.setStyleSheet("""QSplitter#splitter9::handle {background: #DCDCDC; border-style: outset; border-width: 1px; border-color: #828282; border-radius: 3px;}""")

        #сплиттер между кнопкой удаления статусов и кнопкой поиска
        self.splitter25 = QSplitter(Qt.Horizontal)
        self.splitter25.addWidget(self.DelStatyses)
        self.splitter25.addWidget(self.MCPSGoButton)

        #сплиттер между кнопкой удаления статусов\поиска и строкой ввода номера
        self.splitter10 = QSplitter(Qt.Horizontal)
        self.splitter10.addWidget(self.splitter25)
        self.splitter10.addWidget(self.InfoEdit)

        #сплиттер между окнами для вывода и кнопкой поиска\ввода номера
        self.splitter11 = QSplitter(Qt.Vertical)
        self.splitter11.addWidget(self.splitter10)
        self.splitter11.addWidget(self.splitter9)

        tab1hbox = QHBoxLayout()
        tab1hbox.setContentsMargins(5, 5, 5, 5)
        tab1hbox.addWidget(self.splitter11)
        tab1.setLayout(tab1hbox)

        #формирование 2 вкладки#####################################

        tab2 = QWidget()

        #поле вывода Active Offers
        self.OffersGoButton = QPushButton('&Start seach in Offers\RC')
        self.OffersGoButton.setFixedHeight(30)
        self.tableWidgetForActiveOffers = QTableWidget()
        self.tableWidgetForActiveOffers.setRowCount(1)
        self.tableWidgetForActiveOffers.setColumnCount(5)
        self.tableWidgetForActiveOffers.setAlternatingRowColors(1)

        #поле вывода RC
        self.TitleForRC = QLabel('RC:')
        self.tableWidgetForRC = QTableWidget()
        self.tableWidgetForRC.setRowCount(1)
        self.tableWidgetForRC.setColumnCount(20)
        self.tableWidgetForRC.setAlternatingRowColors(1)

        #поле вывода Inactive Offers
        self.InacteveOffersGoButton = QPushButton('&Seach Inactive offers')
        self.tableWidgetForInactiveOffers = QTableWidget()
        self.tableWidgetForInactiveOffers.setRowCount(1)
        self.tableWidgetForInactiveOffers.setColumnCount(6)
        self.tableWidgetForInactiveOffers.setAlternatingRowColors(1)

        # сплиттер активных оферов
        self.splitter12 = QSplitter(Qt.Vertical)
        self.splitter12.addWidget(self.OffersGoButton)
        self.splitter12.addWidget(self.tableWidgetForActiveOffers)

        # сплиттер периодик
        self.splitter13 = QSplitter(Qt.Vertical)
        self.splitter13.addWidget(self.TitleForRC)
        self.splitter13.addWidget(self.tableWidgetForRC)

        # сплиттер не активных оферов
        self.splitter14 = QSplitter(Qt.Vertical)
        self.splitter14.addWidget(self.InacteveOffersGoButton)
        self.splitter14.addWidget(self.tableWidgetForInactiveOffers)

        # сплиттер между активными офферами и периодиками
        self.splitter15 = QSplitter(Qt.Vertical)
        self.splitter15.addWidget(self.splitter12)
        self.splitter15.addWidget(self.splitter13)

        # сплиттер между активными офферами\периодиками и не актывными офферами
        self.splitter16 = QSplitter(Qt.Vertical)
        self.splitter16.addWidget(self.splitter15)
        self.splitter16.addWidget(self.splitter14)

        tab2hbox = QHBoxLayout()
        tab2hbox.setContentsMargins(5, 5, 5, 5)
        tab2hbox.addWidget(self.splitter16)
        tab2.setLayout(tab2hbox)

        #формирование 3 вкладки
        tab3 = QWidget()
        #вывод заявок в хипе
        self.HipInfoButton = QPushButton('&Заявки в хипе')

        #вывод заявок в хипе  с определенным действием
        self.HipByNameInfoButton = QPushButton('&Заявки в хипе с действием:')

        #Ввод имени действия
        self.HipByNameInfoEdit = QLineEdit()
        self.HipByNameInfoEdit.setText("SUBS_ACTIVATE")

        #Таблица вывода заявок в хипе
        self.HipInfoTableWidget = QTableWidget()
        self.HipInfoTableWidget.setRowCount(1)
        self.HipInfoTableWidget.setColumnCount(3)
        self.HipInfoTableWidget.setAlternatingRowColors(1)

        #Таблица вывода заявок в хипе по действиям
        self.HipByNameInfoTableWidget = QTableWidget()
        self.HipByNameInfoTableWidget.setRowCount(1)
        self.HipByNameInfoTableWidget.setColumnCount(26)
        self.HipByNameInfoTableWidget.setAlternatingRowColors(1)

        #сплиттер между кнопкой вывода заявок в хипе и таблицей вывода
        self.splitter26 = QSplitter(Qt.Vertical)
        self.splitter26.addWidget(self.HipInfoButton)
        self.splitter26.addWidget(self.HipInfoTableWidget)

        #сплиттер между кнопкой вывода заявок по имени и полем ввода имени
        self.splitter27 = QSplitter(Qt.Vertical)
        self.splitter27.addWidget(self.HipByNameInfoButton)
        self.splitter27.addWidget(self.HipByNameInfoEdit)

        #сплиттер между к предидущему добавляем таблицу
        self.splitter28 = QSplitter(Qt.Vertical)
        self.splitter28.addWidget(self.splitter27)
        self.splitter28.addWidget(self.HipByNameInfoTableWidget)

        #финальный сплиттер вкладки
        self.splitter29 = QSplitter(Qt.Horizontal)
        self.splitter29.addWidget(self.splitter26)
        self.splitter29.addWidget(self.splitter28)


        tab3hbox = QHBoxLayout()
        tab3hbox.setContentsMargins(5, 5, 5, 5)
        tab3hbox.addWidget(self.splitter29)
        tab3.setLayout(tab3hbox)

        self.bottomLeftTabWidget.addTab(tab1, "&MCPS")
        self.bottomLeftTabWidget.addTab(tab2, "DBM")
        self.bottomLeftTabWidget.addTab(tab3, "Others")

        grid = QGridLayout()
        grid.setSpacing(10)
        grid.addWidget(self.bottomLeftTabWidget, 1, 0)
        self.setLayout(grid)

        self.setGeometry(30, 30, 1800, 900)
        self.setWindowTitle('Subscrider info')
        self.show()

        #мапинг функций на кнопки
        self.MCPSGoButton.clicked.connect(self.MCPS_BD_INFO)
        self.MCPSGoButton.setShortcut("Enter")
        self.DBMGoButton.clicked.connect(self.DBM_DB_INFO)
        self.InacteveOffersGoButton.clicked.connect(self.INACTIVEOFFERS_INFO)
        self.OffersGoButton.clicked.connect(self.OFFERS_INFO)
        self.MCPSOrdersParamButton.clicked.connect(self.ORDERS_PARAM_INFO)
        self.ConvertButton.clicked.connect(self.CONVERT_DATE)
        self.DelStatyses.clicked.connect(self.Window_For_Deleting_Statyses)
        self.HipByNameInfoButton.clicked.connect(self.HipDyNameInfo)
        self.HipInfoButton.clicked.connect(self.HipInfo)
        self.COSVievTitle.clicked.connect(self.COSCheck)

#функция авторизации в БД МСПС

    def MCPS_Authorization_Func(self):
        print("Start connect to MCPS_DB")
        os.environ['NLS_LANG'] = 'UKRAINIAN_UKRAINE.CL8MSWIN1251'
        try:
            mcps_connection = cx_Oracle.connect(MCPS_Authorization_Data)
            return mcps_connection
        except cx_Oracle.DatabaseError as e:
            error, = e.args
            if error.code == 1017:
                print(
                    'Error in connect to MCPS_DB. Please check your credentials.'
                )
                sys.exit()
            else:
                if error.code == 12170:
                    print(
                        'Error in connect to MCPS_DB. Please check DB adress.')
                    print(e)
                    sys.exit()
                else:
                    print(e)
                    sys.exit()
                    raise
            raise
        print("Connecting to MCPS_DB done")


#функция авторизации в БД DBM

    def DBM_Authorization_Func(self):
        print("Start connect to DBM_DB")
        try:
            dbm_connection = cx_Oracle.connect(DBM_Authorization_Data)
        except cx_Oracle.DatabaseError as e:
            error, = e.args
            if error.code == 1017:
                print(
                    'Error in connect to DBM. Please check your credentials...')
                print(e)
                sys.exit()
            else:
                if error.code == 12170:
                    print('Error in connect to DBM. Please check DB adress...')
                    print(e)
                    sys.exit()
                else:
                    print(e)
                    sys.exit()
                    raise
            raise
        print("Connecting to DBM done")
        return dbm_connection

#функция получения prov_key

    def Get_Prov_Key(self):
        prov_key_str = self.InfoEdit.text()

        if prov_key_str[0] != '0':
            prov_key = int(prov_key_str)
            return prov_key
        else:
            QMessageBox.information(self, "Ошибка!",
                                    'Введенный номер начинается с "0"')
            self.InfoEdit.setText("678450249")

    def BigTextRedactor(self, BigText):
        self.BigTextRedactorForm = BigTextFromTable()
        #print(BigText)
        self.BigTextRedactorForm.show()
        self.BigTextRedactorForm.TextFromTable.setText(BigText)
        self.BigTextRedactorForm.CloseTextFromTableButton.clicked.connect(self.CloseBigTextRedactor)
    def CloseBigTextRedactor(self):
        self.BigTextRedactorForm.close()

#Блок работы с окном удаления статусов
    def Window_For_Deleting_Statyses(self):
        self.new_form = DeletedStatyses()
        self.new_form.show()
        #self.new_form.StatysText.setText("Улюблена країна")

        self.new_form.DeletOneStatys.clicked.connect(self.deletOneStatyses)
        self.new_form.DeletAllStatyses.clicked.connect(self.deletAllStatyses)

    def deletAllStatyses(self):
        prov_key = self.Get_Prov_Key()

        if prov_key is not None:

            connection = self.MCPS_Authorization_Func()

            for_deleting_statys_cursor = connection.cursor()
            Stat = str(self.new_form.StatysText.text())
            Staty = """%""" + Stat + """%"""
            Statys = str(Staty)
            #print(Statys)
            print("Start deleting info from DB")
            try:
                for_deleting_statys_cursor.execute(
                    """update mcps.data_states oo set oo.state_id = 2, oo.date_modify = sysdate
                                                       where     
                                                        oo.prov_key =:prov_key
                                                        and oo.state_id = 1
                                                       and oo.object_id in 
                                                       (select cp.id from mcps.conf_products cp where 
                                                            cp.version_id = (select s.id from mcps.versions s 
                                                                                where s.activation_date = (select max(d.activation_date) 
                                                                                from mcps.versions d where d.activation_date < sysdate)) 
                                                        and cp.name like :Statys)""",
                    prov_key=prov_key,
                    Statys=Statys)

                for_deleting_statys_cursor.execute("""commit""")

            except cx_Oracle.DatabaseError as e:
                error, = e.args
                print(e)
                for_deleting_statys_cursor.close()
                connection.close()
                sys.exit()
            print('Deleting done')
            for_deleting_statys_cursor.close()
            connection.close()
            self.new_form.close()
            self.MCPS_BD_INFO()

    def deletOneStatyses(self):
        prov_key = self.Get_Prov_Key()
        print(prov_key)

        if prov_key is not None:

            connection = self.MCPS_Authorization_Func()

            for_deleting_statys_cursor = connection.cursor()

            Statys = str(self.new_form.StatysText.text())

            print(Statys)

            print("Start deleting info from DB\n")

            try:
                for_deleting_statys_cursor.execute(
                    """update mcps.data_states oo set oo.state_id = 2, oo.date_modify = sysdate
                                                       where     
                                                        oo.prov_key =:prov_key
                                                        and oo.state_id = 1
                                                       and oo.object_id in 
                                                       (select cp.id from mcps.conf_products cp 
                                                       where cp.version_id = (select s.id from mcps.versions s where s.activation_date = (
                                                                                select max(d.activation_date) from mcps.versions d where d.activation_date < sysdate)) 
                                                        and cp.name = :Statys)""",
                    prov_key=prov_key,
                    Statys=Statys)
                for_deleting_statys_cursor.execute("""commit""")

            except cx_Oracle.DatabaseError as e:
                error, = e.args
                print(e)
                for_deleting_statys_cursor.close()
                connection.close()
                sys.exit()
            print('Deleting done')
            for_deleting_statys_cursor.close()
            connection.close()
            self.new_form.close()
            self.MCPS_BD_INFO()

    def MCPS_BD_INFO(self):

        prov_key = self.Get_Prov_Key()
        print(prov_key)

        if prov_key is not None:
            Statys = int(self.StatysVievStatys.currentText())

            con = self.MCPS_Authorization_Func()

            statys_cursor = con.cursor()

            print("Start retriave info from DB\n")

            try:
                statys_cursor.execute(
                    """select oo.object_id, cp.name,  vv.name as Instance,  oo.date_modify,  oo.last_date_active, oo.date_to
                                     from mcps.data_states oo, mcps.conf_products cp, mcps.conf_prod_instance vv
                                     where cp.id = oo.object_id and
                                           oo.prov_key=:prov_key and oo.state_id = :Statys
                                           and cp.version_id = (select s.id from mcps.versions s 
                                                                where s.activation_date = (select max(d.activation_date) 
                                                                                            from mcps.versions d where d.activation_date < sysdate))
                                           and vv.version_id = (select s.id from mcps.versions s 
                                                                where s.activation_date = (select max(d.activation_date) 
                                                                                            from mcps.versions d where d.activation_date < sysdate))
                                           and  oo.instance_id = vv.id
                                           order by cp.name, oo.date_modify""",
                    prov_key=prov_key,
                    Statys=Statys)
            except cx_Oracle.DatabaseError as e:
                error, = e.args
                print(e)
                statys_cursor.close()
                con.close()
                sys.exit()

            i = 0
            for column_description in statys_cursor.description:
                self.tableWidgetForStatys.setHorizontalHeaderItem(
                    i,
                    QTableWidgetItem(
                        str("{:<13}".format(*column_description))))
                i = i + 1

            self.tableWidgetForStatys.setRowCount(0)

            for row, form in enumerate(statys_cursor):
                self.tableWidgetForStatys.insertRow(row)
                for column, item in enumerate(form):
                    self.tableWidgetForStatys.setItem(
                        row, column, QTableWidgetItem(str(item)))
            self.tableWidgetForStatys.resizeColumnsToContents()
            self.tableWidgetForStatys.setColumnWidth(1, 300)



            #select for orders
            orders_cursor = con.cursor()

            try:
                orders_cursor.execute(
                    """select o.id, o.prov_key, o.object_id, q.name, to_char(o.order_time, 'yyyy-mm-dd hh24:mi:ss') as ORDER_TIME, 
                                            to_char(o.exec_time, 'yyyy-mm-dd hh24:mi:ss') as EXEC_TIME, o.result_code, o.result_message,
                                            to_char(o.accept_time, 'yyyy-mm-dd hh24:mi:ss') as ACCEPT_TIME, o.action_name, s.name as kanal, 
                                            o.id_in_source, t.name as typ_obrabotki , o.order_type,
                                            c.name as source, o.retry_count, to_char(o.retry_time, 'yyyy-mm-dd hh24:mi:ss') as RETRY_TIME
                                     from mcps.data_orders o,
                                         (select * from mcps.conf_products cp
                                          where cp.version_id = (SELECT MAX(v.id) FROM mcps.versions v WHERE v.activation_date IS NOT NULL)
                                          and cp.is_top = 1) q,
                                          mcps.dict_channels c,
                                          mcps.dict_sources s,
                                          mcps.dict_processing_types t
                                     where
                                          o.prov_key=:prov_key
                                          and o.object_id = q.id
                                          and o.source = s.id
                                          and o.channel = c.id
                                          and o.processing_type = t.id
                                          ORDER BY o.order_time""",
                    prov_key=prov_key)
            except cx_Oracle.DatabaseError as e:
                error, = e.args
                print(e)
                orders_cursor.close()
                con.close()
                sys.exit()
            print("Retreaving done")
            i = 0
            for column_description in orders_cursor.description:
                self.tableWidgetForOrders.setHorizontalHeaderItem(
                    i,
                    QTableWidgetItem(
                        str("{:<13}".format(*column_description))))
                i = i + 1

            self.tableWidgetForOrders.setRowCount(0)
            for row, form in enumerate(orders_cursor):
                self.tableWidgetForOrders.insertRow(row)
                for column, item in enumerate(form):
                    self.tableWidgetForOrders.setItem(
                        row, column, QTableWidgetItem(str(item)))
            self.tableWidgetForOrders.resizeColumnsToContents()
            self.tableWidgetForOrders.setColumnWidth(3, 250)
            self.tableWidgetForOrders.setColumnWidth(7, 300)
            statys_cursor.close()
            orders_cursor.close()
            con.close()

            #Параметры заявки в окне вывода по двойному клику
            self.tableWidgetForOrders.cellDoubleClicked.connect(self.ExecuteOrdersNumber)
            self.tableWidgetForStatys.cellDoubleClicked.connect(self.StatysBigText)
    def ExecuteOrdersNumber(self):
        OrdersColumn=self.tableWidgetForOrders.currentColumn()
        OrdersRow=self.tableWidgetForOrders.currentRow()
        OrderItem=self.tableWidgetForOrders.item(OrdersRow,OrdersColumn)
        OrderItemText=OrderItem.text()

        if OrdersColumn == 0:
            self.OrdersNumberEdit.setText(OrderItemText)
            self.ORDERS_PARAM_INFO()
        else:
            print('Order number is not selected')
            self.BigTextRedactor(OrderItemText)

    def StatysBigText(self):
        OrdersColumnForStatys = self.tableWidgetForStatys.currentColumn()
        OrdersRowForStatys = self.tableWidgetForStatys.currentRow()
        OrderItemForStatys = self.tableWidgetForStatys.item(OrdersRowForStatys, OrdersColumnForStatys)
        OrderItemText = OrderItemForStatys.text()
        self.BigTextRedactor(OrderItemText)

    def Error_mesage(self, message):
        QMessageBox.information(self, "Ошибка!", message)

    def COSCheck(self):
        prov_key = self.Get_Prov_Key()
        print(prov_key)

        if prov_key is not None:

            con = self.MCPS_Authorization_Func()
            orders_cursor = con.cursor()
            #Вытягивание кос нейма
            prov_key = str(prov_key)

            print("Start check COS info")
            try:
                orders_cursor.execute(
                    """select TTT.COS_NAME
                                        from mcps.spsdb_pps_subscriber TTT
                                       where TTT.snap_date = trunc(sysdate)
                                       and TTT.SUBSCRIBER_ID=:prov_key""",
                    prov_key=prov_key)
            except cx_Oracle.DatabaseError as e:
                error, = e.args
                print(e)
                orders_cursor.close()
                con.close()
                sys.exit()

            print("Check info done")

            Orders_cursor_fatchall = orders_cursor.fetchall()
            #print(Orders_cursor_fatchall)

            if not Orders_cursor_fatchall:
                print('empty cos')
                self.Error_mesage(
                    'Отсутствует статус по ТП в from mcps.spsdb_pps_subscriber')
            else:
                try:
                    pps_cos = str(
                        """{:<20}""".format(*Orders_cursor_fatchall[0]))
                    pps_cos = pps_cos.strip()

                except Exception as e:
                    error, = e.args
                    pps_cos = 'Empty'
                    print(e)

            #блок сравнения ТП из разных источников
            prov_key = self.Get_Prov_Key()
            Statys = int(self.StatysVievStatys.currentText())
            statys_cursor = con.cursor()
            try:
                statys_cursor.execute(
                    """select cp.description
                                     from mcps.data_states oo, mcps.conf_products cp, mcps.conf_prod_instance vv
                                     where cp.id = oo.object_id and
                                           oo.prov_key=:prov_key and oo.state_id = :Statys
                                           and cp.version_id = (select s.id from mcps.versions s 
                                                                where s.activation_date = (select max(d.activation_date) 
                                                                                            from mcps.versions d where d.activation_date < sysdate))
                                           and vv.version_id = (select s.id from mcps.versions s 
                                                                where s.activation_date = (select max(d.activation_date) 
                                                                                            from mcps.versions d where d.activation_date < sysdate))
                                           and  oo.instance_id = vv.id
                                            and cp.description is not null""",
                    prov_key=prov_key,
                    Statys=Statys)
            except cx_Oracle.DatabaseError as e:
                error, = e.args
                print(e)
                statys_cursor.close()
                con.close()
                sys.exit()

            r = statys_cursor.fetchall()

            if not r:
                print('empty cos')
                self.COSVievTitle.setText(
                    'Отсутствует статус по ТП в mcps.conf_products' + '\n' +
                    'В spsdb_pps_subscriber ТП: ' + pps_cos)
                self.Error_mesage(
                    'Отсутствует статус по ТП в mcps.conf_products')
                statys_cursor.close()
                orders_cursor.close()
                con.close()

            else:

                try:
                    mcps_cos = str("""{:<20}""".format(*r[0]))
                    mcps_cos = mcps_cos.strip()

                except Exception as e:
                    error, = e.args
                    print(e)


                try:
                    if str(mcps_cos) == str(pps_cos):
                        print('ТП из в разных источниках совпадают')
                        pps_cos = str(pps_cos)
                        TP_cursor = con.cursor()
                        TP_cursor.execute(
                            """select cp.name
                                     from mcps.conf_products cp                                            
                                     where   cp.version_id = (select s.id from mcps.versions s 
                                                                where s.activation_date = (select max(d.activation_date) 
                                                                                            from mcps.versions d where d.activation_date < sysdate))      
                                     and cp.description is not null    
                                     and cp.description =:pps_cos""",
                            pps_cos=pps_cos)
                        name = TP_cursor.fetchall()

                        name = str("""{:<20}""".format(*name[0]))

                        self.COSVievTitle.setText(
                            'Текущий: ' + name + '\n' +
                            'ТП в разных источниках совпадают.')
                        TP_cursor.close()
                    else:
                        self.Error_mesage(
                            'ТП в mcps.conf_products и mcps.spsdb_pps_subscriber не совпадают!'
                        )

                        self.COSVievTitle.setText(
                            'ТП в mcps.conf_products и mcps.spsdb_pps_subscriber не совпадают!' + '\n' +
                            'В spsdb_pps_subscriber ТП: ' + pps_cos)

                except Exception as e:
                    error, = e.args
                    print(e)


                statys_cursor.close()
                orders_cursor.close()
                con.close()

    def DBM_DB_INFO(self):
        prov_key = self.Get_Prov_Key()
        print(prov_key)

        if prov_key is not None:

            dbm_con = self.DBM_Authorization_Func()

            Activity_cursor = dbm_con.cursor()

            print("Cursor created")

            prov_key = str(prov_key)

            print("Prov key: " + prov_key)

            year = self.FromDate.date().year()
            month = self.FromDate.date().month()
            day = self.FromDate.date().day()
            from_date = datetime.datetime(year, month, day)
            from_date = str(
                """{:%d.%m.%y}""".format(datetime.datetime(year, month, day)))
            print("From date: " + from_date)

            year = self.ToDate.date().year()
            month = self.ToDate.date().month()
            day = self.ToDate.date().day()
            to_date = str(
                """{:%d.%m.%y}""".format(datetime.datetime(year, month, day)))
            print("To date: " + to_date)

            print("Start check info")
            try:
                Activity_cursor.execute(
                    """select   tt.src_table_name, tt.start_call_date_time, tt.end_call_date_time, tt.called_number, tt.called_number_norm, tt.call_type,tt.change_amount_1 as CORE_BALANCE, tt.balance1,
                                                tt.comments,  tt.term_name, tt.change_amount_2 as FREE_SEC_BALANCE, tt.balance2, tt.change_amount_3 as FREE_SEC_BALANCE2, tt.balance3,  
                                                tt.change_amount_4 as MMS_BALANCE, tt.balance4, tt.change_amount_5,  tt.balance5, tt.change_amount_25 as GPRS_BALANCE6, tt.balance25,
                                                tt.change_amount_6 as SMS_BALANCE2, tt.balance6, tt.change_amount_7 as FREE_MONEY, tt.balance7, tt.change_amount_8 as FREE_MONEY2,
                                                tt.balance8, tt.change_amount_9 as FREE_SEC_BALANCE3, tt.balance9, tt.change_amount_10 as FREE_MONEY3,
                                                tt.balance10, tt.change_amount_11 as GPRS_BALANCE, tt.balance11, tt.change_amount_12 as FREE_MONEY4, tt.balance12, tt.change_amount_13 as GPRS_BALANCE2, tt.balance13, tt.change_amount_14 as CONTENT_BALANCE, tt.balance14,
                                                tt.change_amount_15 as SMS_BALANCE3, tt.balance15, tt.change_amount_16 as MMS_BALANCE2, tt.balance16, tt.change_amount_17 as FREE_SEC_BALANCE5, tt.balance17,
                                                tt.change_amount_18 as GPRS_BALANCE3, tt.balance18, tt.change_amount_19 as FREE_SEC_BALANCE6,
                                                tt.balance19, tt.change_amount_20 as GPRS_BALANCE4, tt.balance20, tt.change_amount_21 as SMS_BALANCE4, tt.balance21, tt.change_amount_22 as FREE_SEC_BALANCE7, tt.balance22, tt.change_amount_23 as FREE_MONEY5, tt.balance23,
                                                tt.change_amount_30 as FREE_SEC_BALANCE4, tt.balance30, tt.change_amount_210 as SMS_BALANCE4, tt.balance210
                                       from c1.call_history$$ tt
                                       where tt.external_id=:prov_key and end_call_date_time between :from_date  and :to_date and  (change_amount_1 != 0.000000
                                             or change_amount_2 != 0.000000 or change_amount_3 != 0.000000 or change_amount_4 != 0.000000 or change_amount_5 != 0.000000
                                             or change_amount_6 != 0.000000 or change_amount_7 != 0.000000 or change_amount_8 != 0.000000 or change_amount_9 != 0.000000
                                             or change_amount_10 != 0.000000 or change_amount_11 != 0.000000 or change_amount_12 != 0.000000 or change_amount_13 != 0.000000
                                             or change_amount_14 != 0.000000 or change_amount_15 != 0.000000 or change_amount_16 != 0.000000 or change_amount_17 != 0.000000
                                             or change_amount_19 != 0.000000 or change_amount_18 != 0.000000 or change_amount_20 != 0.000000 or change_amount_21 != 0.000000
                                            or change_amount_22 != 0.000000 or change_amount_23 != 0.000000 or change_amount_30 != 0.000000 or change_amount_210 != 0.000000 or change_amount_25 != 0.000000)
                                            order by tt.end_call_date_time""",
                    prov_key=prov_key,
                    from_date=from_date,
                    to_date=to_date)
            except cx_Oracle.DatabaseError as e:
                error, = e.args
                print(e)
                Activity_cursor.close()
                dbm_con.close()
                sys.exit()

            print("Check info done")
            i = 0
            print("Start print table header")
            for column_description in Activity_cursor.description:
                self.tableWidgetForActive.setHorizontalHeaderItem(
                    i,
                    QTableWidgetItem(
                        str("{:<13}".format(*column_description))))
                i = i + 1
            print("Start print table")
            self.tableWidgetForActive.setRowCount(0)
            for row, form in enumerate(Activity_cursor):
                self.tableWidgetForActive.insertRow(row)
                for column, item in enumerate(form):
                    self.tableWidgetForActive.setItem(
                        row, column, QTableWidgetItem(str(item)))
            print("Printing table done")
            Activity_cursor.close()
            dbm_con.close()

    def INACTIVEOFFERS_INFO(self):
        prov_key = self.Get_Prov_Key()
        print(prov_key)

        if prov_key is not None:

            InactiveOffers_con = self.DBM_Authorization_Func()

            InactiveOffers_cursor = InactiveOffers_con.cursor()
            prov_key = str(prov_key)
            print("Start check info for inactive offers")

            try:
                InactiveOffers_cursor.execute(
                    """select g.offer_id, cff.display_value, g.active_dt, g.inactive_dt, cff.description, g.disconnect_reason
                                               from c1.offer_inst g,
                                                    C1.ACCOUNT_SUBSCRIBER tt,
                                                    c1.offer cff
                                               where tt.range_map_external_id=:prov_key
                                                    and g.subscr_no = tt.subscr_no
                                                    and g.disconnect_reason is not null
                                                    and tt.snap_date = to_date(sysdate-1)
                                                    and cff.offer_type = 3
                                                    and cff.reseller_version_id =
                                                           (SELECT MAX(v.reseller_version_id)
                                                            FROM c1.reseller_version v
                                                            WHERE v.active_date IS NOT NULL)
                                                    and trunc(cff.dwh_load_dt) = trunc(sysdate)
                                                    and cff.language_code = 1
                                                    and g.offer_id = cff.offer_id""",
                    prov_key=prov_key)
            except cx_Oracle.DatabaseError as e:
                error, = e.args
                print(e)
                InactiveOffers_cursor.close()
                InactiveOffers_con.close()
                sys.exit()
            print("Check info done")

            i = 0
            print("Start print table header")
            for column_description in InactiveOffers_cursor.description:
                self.tableWidgetForInactiveOffers.setHorizontalHeaderItem(
                    i,
                    QTableWidgetItem(
                        str("{:<13}".format(*column_description))))
                i = i + 1

            print("Start print table")
            self.tableWidgetForInactiveOffers.setRowCount(0)
            for row, form in enumerate(InactiveOffers_cursor):
                self.tableWidgetForInactiveOffers.insertRow(row)
                for column, item in enumerate(form):
                    self.tableWidgetForInactiveOffers.setItem(
                        row, column, QTableWidgetItem(str(item)))
            print("Printing table done")
            InactiveOffers_cursor.close()
            InactiveOffers_con.close()

    def OFFERS_INFO(self):
        prov_key = self.Get_Prov_Key()
        print(prov_key)

        if prov_key is not None:
            self.tableWidgetForActiveOffers.setItem(
                0, 0,
                QTableWidgetItem(
                    "Info retreaving from INACTIVEOFFERS_INFO...."))
            print("Start connect to dbm for OFFERS_INFO")

            ActiveOffers_con = self.DBM_Authorization_Func()

            ActiveOffers_cursor = ActiveOffers_con.cursor()

            print("Start check info for active offers")

            try:
                ActiveOffers_cursor.execute(
                    """select   g.offer_id, cff.display_value, g.active_dt, g.inactive_dt, cff.description
                                             from c1.offer_inst g, C1.ACCOUNT_SUBSCRIBER tt,    c1.offer cff
                                             where tt.range_map_external_id=:prov_key
                                                   and g.subscr_no = tt.subscr_no
                                                   and g.offer_state = 1
                                                   and g.disconnect_reason is null
                                                   and tt.snap_date = to_date(sysdate-1)
                                                   and cff.offer_type = 3
                                                   and cff.reseller_version_id =
                                                  (SELECT MAX(v.reseller_version_id)
                                                     FROM c1.reseller_version v
                                                     WHERE v.active_date IS NOT NULL)
                                                     and trunc(cff.dwh_load_dt) = trunc(sysdate)
                                                   and cff.language_code = 1
                                                   and g.offer_id = cff.offer_id""",
                    prov_key=prov_key)
            except cx_Oracle.DatabaseError as e:
                error, = e.args
                print(e)
                ActiveOffers_cursor.close()
                ActiveOffers_con.close()
                sys.exit()

            print("Check info done")

            i = 0
            print("Start print table header")
            for column_description in ActiveOffers_cursor.description:
                self.tableWidgetForActiveOffers.setHorizontalHeaderItem(
                    i,
                    QTableWidgetItem(
                        str("{:<13}".format(*column_description))))
                i = i + 1

            print("Start print table")
            self.tableWidgetForActiveOffers.setRowCount(0)

            for row, form in enumerate(ActiveOffers_cursor):
                self.tableWidgetForActiveOffers.insertRow(row)
                for column, item in enumerate(form):
                    self.tableWidgetForActiveOffers.setItem(
                        row, column, QTableWidgetItem(str(item)))

            print("Printing table done")
            ActiveOffers_cursor.close()

            ActiveRC_cursor = ActiveOffers_con.cursor()

            print("Start check info for active RC")

            try:
                ActiveRC_cursor.execute(
                    """ select   t2.display_value, t2.description, t1.apply_day, t1.rc_term_inst_active_dt as ACTIVE_DATE, t1.rc_term_inst_inactive_dt as INACTIVE_DATE,  t1.last_processed_dt,
                                                 t1.last_charged_dt, t1.chg_who, t1.last_generate_dt, t1.last_apply_dt, t1.processing_status, t1.retry_count, t1.prev_rc_term_inst_inactive_dt,
                                                 t1.status, t1.period_frequency, t1.process_error_code,  t1.chg_dt, t1.offer_inst_id,  t1.disconnect_reason, t1.total_charged_amount
                                    from
                                           (select *
                                              from c1.rc_term tt where
                                              tt.reseller_version_id = (SELECT MAX(v.reseller_version_id) FROM c1.reseller_version v WHERE v.active_date IS NOT NULL)
                                              and tt.language_code = 1) t2,
                                           (select *
                                              from  c1.rc_term_inst rci
                                              where rci.PARENT_SUBSCR_NO= (select t.subscr_no from C1.ACCOUNT_SUBSCRIBER t
                                              where t.snap_date >= trunc(sysdate)
                                              and t.range_map_external_id=:prov_key)
                                              and trunc(rci.DWH_LOAD_DT)>=trunc(sysdate)) t1
                                   where t1.rc_term_id = t2.rc_term_id and t1.disconnect_reason is null order by t2.rc_term_id""",
                    prov_key=prov_key)
            except cx_Oracle.DatabaseError as e:
                error, = e.args
                print(e)
                ActiveRC_cursor.close()
                ActiveOffers_con.close()
                sys.exit()
            print("Check RC info done")

            i = 0
            print("Start print table header")
            for column_description in ActiveRC_cursor.description:
                self.tableWidgetForRC.setHorizontalHeaderItem(
                    i,
                    QTableWidgetItem(
                        str("{:<13}".format(*column_description))))
                i = i + 1

            print("Start print table")
            self.tableWidgetForRC.setRowCount(0)

            for row, form in enumerate(ActiveRC_cursor):
                self.tableWidgetForRC.insertRow(row)
                for column, item in enumerate(form):
                    self.tableWidgetForRC.setItem(row, column,
                                                  QTableWidgetItem(str(item)))

            print("Printing table done")

            ActiveRC_cursor.close()

            ActiveOffers_con.close()

    def ORDERS_PARAM_INFO(self):
        print("Start connect to MCPS_DB")
        os.environ['NLS_LANG'] = 'UKRAINIAN_UKRAINE.CL8MSWIN1251'

        con = self.MCPS_Authorization_Func()
        #select for orders params
        order_params_cursor = con.cursor()

        try:
            order_id = int(self.OrdersNumberEdit.text())
        except ValueError as e:
            error, = e.args
            print(
                "Введено некорректное число (Скорее всего присутствует буква)")
            print(e)
            con.close()
            order_params_cursor.close()
            con.close()
            sys.exit()

        print("Start retriave info from MCPS_DB\n")

        try:
            order_params_cursor.execute(
                """ SELECT T1.Param_Name, T1.Param_Value
                                            FROM  MCPS.DATA_ORDERS_PARAMS T1
                                            where T1.ORDER_ID =:order_id
                                            order by T1.Param_Name""",
                order_id=order_id)
        except cx_Oracle.DatabaseError as e:
            error, = e.args
            print(e)
            order_params_cursor.close()
            con.close()
            sys.exit()

        i = 0
        for column_description in order_params_cursor.description:
            self.tableWidgetForOrdersParam.setHorizontalHeaderItem(
                i, QTableWidgetItem(str("{:<13}".format(*column_description))))
            i = i + 1

        self.tableWidgetForOrdersParam.setRowCount(0)
        for row, form in enumerate(order_params_cursor):
            self.tableWidgetForOrdersParam.insertRow(row)
            for column, item in enumerate(form):
                self.tableWidgetForOrdersParam.setItem(
                    row, column, QTableWidgetItem(str(item)))

        order_params_cursor.close()
        con.close()
        print('Params successfully retrieving')
        self.tableWidgetForOrdersParam.cellDoubleClicked.connect(self.ConvertDateFronOrdersParam)

    def ConvertDateFronOrdersParam(self):
            OrdersColumn = self.tableWidgetForOrdersParam.currentColumn()
            OrdersRow = self.tableWidgetForOrdersParam.currentRow()
            OrderItem = self.tableWidgetForOrdersParam.item(OrdersRow, OrdersColumn)
            OrderItemText = OrderItem.text()
            IsStringOnlyFromDigit=OrderItemText.isdigit()
            if IsStringOnlyFromDigit:
                self.ConvertIn.setText(OrderItemText)
                self.CONVERT_DATE()
            else:
                print('Not only digit')
                self.BigTextRedactor(OrderItemText)

    def CONVERT_DATE(self):

        os.environ['NLS_LANG'] = 'UKRAINIAN_UKRAINE.CL8MSWIN1251'
        con = self.MCPS_Authorization_Func()
        convert_cursor = con.cursor()

        if self.ConvertIn.text() != '':
            time = int(self.ConvertIn.text())
        else:
            time=9999999999999

        print("Start retriave info from MCPS_DB\n")


        try:
            convert_cursor.execute(
                """ select  NUMTODSINTERVAL(:time  / 1000,
                                                  'SECOND')+TO_DATE('1970-01-01 02:00:00',
                                                  'yyyy-mm-dd hh24:mi:ss')
                                           from dual """,
                    time=time)
        except cx_Oracle.DatabaseError as e:
            error, = e.args
            print(e)
            convert_cursor.close()
            con.close()
            convert_cursor=None
            #sys.exit()
        if convert_cursor is not None:
            for row in convert_cursor:
                self.ConvertOut.setText(str("""{:%Y.%m.%d      %H:%M:%S}""".format(*row) + "\n"))
            convert_cursor.close()
            con.close()
            print('Date converted')
        else:
            self.ConvertOut.setText('Введено некорректное число')

    def HipInfo(self):

        con = self.MCPS_Authorization_Func()

        hip_cursor = con.cursor()

        try:
            hip_cursor.execute(
                """  SELECT count(pp.action_name) as KOLICHESTVO,pp.action_name, pp.priority
                                        FROM mcps.data_orders_heap pp
                                        where 
                                        pp.order_time > trunc(sysdate)
                                        group by pp.action_name, pp.priority
                                        order by pp.priority""")
        except cx_Oracle.DatabaseError as e:
            error, = e.args
            print(e)
            hip_cursor.close()
            con.close()
            sys.exit()

        i = 0
        for column_description in hip_cursor.description:
            self.HipInfoTableWidget.setHorizontalHeaderItem(
                i, QTableWidgetItem(str("{:<13}".format(*column_description))))
            i = i + 1

        self.HipInfoTableWidget.setRowCount(0)
        for row, form in enumerate(hip_cursor):
            self.HipInfoTableWidget.insertRow(row)
            for column, item in enumerate(form):
                self.HipInfoTableWidget.setItem(row, column,
                                                QTableWidgetItem(str(item)))
        hip_cursor.close()
        con.close()

    def HipDyNameInfo(self):
        con = self.MCPS_Authorization_Func()

        hip_cursor = con.cursor()

        name = self.HipByNameInfoEdit.text()
        name = name.strip()

        try:
            hip_cursor.execute(
                """select * from mcps.data_orders_heap pp,
                                    (select * from mcps.conf_products cp
                                    where cp.version_id = (select s.id from mcps.versions s 
                                                                where s.activation_date = (select max(d.activation_date) 
                                                                                            from mcps.versions d where d.activation_date < sysdate))
                                    and cp.is_top = 1) q
                                    where pp.action_name = :name
                                    and pp.object_id = q.id""",
                name=name)
        except cx_Oracle.DatabaseError as e:
            error, = e.args
            print(e)
            hip_cursor.close()
            con.close()
            sys.exit()

        i = 0
        for column_description in hip_cursor.description:
            self.HipByNameInfoTableWidget.setHorizontalHeaderItem(
                i, QTableWidgetItem(str("{:<13}".format(*column_description))))
            i = i + 1

        self.HipByNameInfoTableWidget.setRowCount(0)
        for row, form in enumerate(hip_cursor):
            self.HipByNameInfoTableWidget.insertRow(row)
            for column, item in enumerate(form):
                self.HipByNameInfoTableWidget.setItem(
                    row, column, QTableWidgetItem(str(item)))
        hip_cursor.close()
        con.close()

if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = PrimaryWindow()
    sys.exit(app.exec_())

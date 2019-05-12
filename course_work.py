import sys
import os
import xlrd
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QIcon
from PyQt5 import QtWidgets, QtCore, uic


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.ui = uic.loadUi('university.ui')
        self.ui.setStyleSheet('background-image: url(888.jpg)')
        self.ui.setWindowTitle('Помічник абітурієнта')
        self.ui.setWindowIcon(QIcon('sova_1.png'))

        self.ui.pushButton.clicked.connect(self.university_DONNU)
        self.ui.pushButton_2.clicked.connect(self.university_VNTU)
        self.ui.pushButton_3.clicked.connect(self.university_VSPU)
        self.ui.pushButton_4.clicked.connect(self.university_VSAU)
        self.ui.pushButton_5.clicked.connect(self.university_VNMU)
        self.ui.pushButton_6.clicked.connect(self.university_VFEU)

        self.ui.push_Button_1.clicked.connect(self.top)
        self.push_Button_1 = self.ui.push_Button_1
        self.push_Button_1.setIcon(QIcon('2.png'))
        self.push_Button_1.setIconSize(QSize(80, 80))

        self.ui.push_Button_2.clicked.connect(self.info_1)
        self.push_Button_2 = self.ui.push_Button_2
        self.push_Button_2.setIcon(QIcon('4.png'))
        self.push_Button_2.setIconSize(QSize(80, 80))

        self.ui.push_Button_3.clicked.connect(self.info_2)
        self.push_Button_3 = self.ui.push_Button_3
        self.push_Button_3.setIcon(QIcon('3.png'))
        self.push_Button_3.setIconSize(QSize(80, 80))

        self.ui.cl_Button.clicked.connect(self.clock)
        self.cl_Button = self.ui.cl_Button
        self.cl_Button.setIcon(QIcon('5.png'))
        self.cl_Button.setIconSize(QSize(80, 80))

        self.ui.date_Button.clicked.connect(self.calendar)
        self.date_Button = self.ui.date_Button
        self.date_Button.setIcon(QIcon('6.png'))
        self.date_Button.setIconSize(QSize(80, 80))

        self.path = os.getcwd()
        self.ui.show()

    def top(self):
        self.mw = TOP()
        self.mw.show()

    def clock(self):
        self.cl = Clock()


    def calendar(self):
        self.cal = Calendar()

    def university_VSPU(self):

            self.vspu = uic.loadUi('university_VSPU.ui')
            self.vspu.setWindowTitle('Помічник абітурієнта')
            self.vspu.setWindowIcon(QIcon('sova_1.png'))
            self.vspu.setStyleSheet('background-color: #051721')
            self.work_ui = self.vspu

            self.vspu.pushButton.clicked.connect(self.open_fac)
            self.vspu.pushButton_2.clicked.connect(self.open_fac)
            self.vspu.pushButton_3.clicked.connect(self.open_fac)
            self.vspu.pushButton_4.clicked.connect(self.open_fac)
            self.vspu.pushButton_5.clicked.connect(self.open_fac)
            self.vspu.pushButton_6.clicked.connect(self.open_fac)
            self.vspu.pushButton_7.clicked.connect(self.open_fac)
            self.vspu.pushButton_8.clicked.connect(self.open_fac)
            self.vspu.show()

    def university_DONNU(self):

        self.donnu = uic.loadUi('university_DONNU.ui')
        self.donnu.setWindowTitle('Помічник абітурієнта')
        self.donnu.setWindowIcon(QIcon('sova_1.png'))
        self.donnu.setStyleSheet('background-color: #051721')
        self.work_ui = self.donnu

        self.donnu.pushButton.clicked.connect(self.open_fac)
        self.donnu.pushButton_2.clicked.connect(self.open_fac)
        self.donnu.pushButton_3.clicked.connect(self.open_fac)
        self.donnu.pushButton_4.clicked.connect(self.open_fac)
        self.donnu.pushButton_5.clicked.connect(self.open_fac)
        self.donnu.pushButton_6.clicked.connect(self.open_fac)
        self.donnu.pushButton_7.clicked.connect(self.open_fac)
        self.donnu.pushButton_8.clicked.connect(self.open_fac)
        self.donnu.pushButton_9.clicked.connect(self.open_fac)

        self.donnu.show()

    def university_VFEU(self):

        self.vfeu = uic.loadUi('university_VFEU.ui')
        self.vfeu.setWindowTitle('Помічник абітурієнта')
        self.vfeu.setWindowIcon(QIcon('sova_1.png'))
        self.vfeu.setStyleSheet('background-color: #051721')
        self.work_ui = self.vfeu

        self.vfeu.pushButton.clicked.connect(self.open_fac)
        self.vfeu.pushButton_2.clicked.connect(self.open_fac)
        self.vfeu.pushButton_3.clicked.connect(self.open_fac)
        self.vfeu.pushButton_4.clicked.connect(self.open_fac)
        self.vfeu.show()

    def university_VNMU(self):

        self.vnmu = uic.loadUi('university_VNMU.ui')
        self.vnmu.setWindowTitle('Помічник абітурієнта')
        self.vnmu.setWindowIcon(QIcon('sova_1.png'))
        self.vnmu.setStyleSheet('background-color: #051721')
        self.work_ui = self.vnmu

        self.vnmu.pushButton.clicked.connect(self.open_fac)
        self.vnmu.pushButton_2.clicked.connect(self.open_fac)
        self.vnmu.pushButton_3.clicked.connect(self.open_fac)
        self.vnmu.pushButton_4.clicked.connect(self.open_fac)
        self.vnmu.pushButton_5.clicked.connect(self.open_fac)
        self.vnmu.pushButton_6.clicked.connect(self.open_fac)

        self.vnmu.show()

    def university_VNTU(self):

        self.vntu = uic.loadUi('university_VNTU.ui')
        self.vntu.setWindowTitle('Помічник абітурієнта')
        self.vntu.setStyleSheet('background-color: #051721')
        self.vntu.setWindowIcon(QIcon('sova_1.png'))
        self.work_ui = self.vntu

        self.vntu.pushButton.clicked.connect(self.open_fac)
        self.vntu.pushButton_2.clicked.connect(self.open_fac)
        self.vntu.pushButton_3.clicked.connect(self.open_fac)
        self.vntu.pushButton_4.clicked.connect(self.open_fac)
        self.vntu.pushButton_5.clicked.connect(self.open_fac)
        self.vntu.pushButton_6.clicked.connect(self.open_fac)
        self.vntu.pushButton_7.clicked.connect(self.open_fac)
        self.vntu.pushButton_8.clicked.connect(self.open_fac)
        self.vntu.show()

    def university_VSAU(self):

        self.vsau = uic.loadUi('university_VSAU.ui')
        self.vsau.setWindowTitle('Помічник абітурієнта')
        self.vsau.setWindowIcon(QIcon('sova_1.png'))
        self.vsau.setStyleSheet('background-color: #051721')
        self.work_ui = self.vsau

        self.vsau.pushButton.clicked.connect(self.open_fac)
        self.vsau.pushButton_2.clicked.connect(self.open_fac)
        self.vsau.pushButton_3.clicked.connect(self.open_fac)
        self.vsau.pushButton_4.clicked.connect(self.open_fac)
        self.vsau.pushButton_5.clicked.connect(self.open_fac)
        self.vsau.pushButton_6.clicked.connect(self.open_fac)
        self.vsau.pushButton_7.clicked.connect(self.open_fac)
        self.vsau.show()

    def info_1(self):

        self.info_1 = uic.loadUi('INFO_1.ui')
        self.info_1.setWindowTitle('Помічник абітурієнта')
        self.info_1.setWindowIcon(QIcon('sova_1.png'))
        self.info_1.setStyleSheet('background-color: #051721')
        self.info_1.show()

    def info_2(self):

        self.info_2 = uic.loadUi('INFO_2.ui')
        self.info_2.setWindowTitle('Помічник абітурієнта')
        self.info_2.setStyleSheet('background-color: #051721')
        self.info_2.setWindowIcon(QIcon('sova_1.png'))

        self.info_2.show()

    def open_fac(self):

        self.unpath = self.path + self.sender().statusTip()
        file = open(self.unpath)
        text = ''
        for row in file:
            text += row
        self.work_ui.plainTextEdit.setPlainText(text)

class TOP(QMainWindow):

    def __init__(self):
        QMainWindow.__init__(self)

        self.setMinimumSize(QSize(1500, 600))

        self.setWindowTitle("Консолідований рейтинг вищів України 2018")
        self.setWindowIcon(QIcon('sova_1.png'))
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        grid_layout = QGridLayout()
        central_widget.setLayout(grid_layout)

        table = QTableWidget(self)
        table.setColumnCount(6)
        table.setRowCount(273)


        table.setHorizontalHeaderLabels(["Назва навчального закладу", "Місце у загальному рейтингу", "ТОП 200 України",
                                         "Scopus", "Бал ЗНО на контракт", "Підсумковий бал"])

        rb = xlrd.open_workbook('топ университетов украины.xlsx')
        sheet = rb.sheet_by_index(0)

        val = sheet.col_values(0)
        for i, entry in enumerate(val):
            table.setRowCount(i + 1)
            item = QTableWidgetItem(0)
            item.setText(str(entry))
            table.setItem(i, 0, item)

        val2 = sheet.col_values(1)
        for q, entry in enumerate(val2):
            table.setRowCount(i + 1)
            item = QTableWidgetItem(0)
            item.setText(str(entry))
            table.setItem(q, 1, item)

        val3 = sheet.col_values(2)
        for w, entry in enumerate(val3):
            table.setRowCount(i + 1)
            item = QTableWidgetItem(0)
            item.setText(str(entry))
            table.setItem(w, 2, item)

        val4 = sheet.col_values(3)
        for a, entry in enumerate(val4):
            table.setRowCount(i + 1)
            item = QTableWidgetItem(0)
            item.setText(str(entry))
            table.setItem(a, 3, item)

        val5 = sheet.col_values(4)
        for s, entry in enumerate(val5):
            table.setRowCount(i + 1)
            item = QTableWidgetItem(0)
            item.setText(str(entry))
            table.setItem(s, 4, item)

        val6 = sheet.col_values(5)
        for v, entry in enumerate(val6):
            table.setRowCount(i + 1)
            item = QTableWidgetItem(0)
            item.setText(str(entry))
            table.setItem(v, 5, item)

        table.resizeColumnsToContents()
        grid_layout.addWidget(table, 0, 0)
        self.show()


class Clock(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(Clock, self).__init__(parent)

        self.ui = uic.loadUi('clock.ui')
        self.ui.setWindowTitle('Помічник абітурієнта')
        self.ui.setWindowIcon(QIcon('sova_1.png'))
        timer = QtCore.QTimer(self)
        timer.start(1000)
        timer.timeout.connect(self.showTime)
        self.showTime()

        self.ui.show()

    def showTime(self):
        time = QtCore.QTime.currentTime()
        text = time.toString('hh:mm')
        self.ui.lcdNumber.display(text)

        
class Calendar(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(Calendar, self).__init__(parent)

        self.ui = uic.loadUi('calendar.ui')
        self.ui.setWindowTitle('Помічник абітурієнта')
        self.ui.setWindowIcon(QIcon('sova_1.png'))

        self.ui.calendarWidget.clicked.connect(self.show_date)
        self.ui.show()

    def show_date(self, date):
        self.ui.label.setText(date.toString())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec_())

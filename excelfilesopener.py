import sys

import openpyxl
from PySide2 import QtCore, QtGui, QtWidgets
from PySide2.QtCore import Qt, QSize
from PySide2.QtGui import QPixmap, QPainter, QIcon
from PySide2.QtWidgets import QHBoxLayout, QPushButton, QMainWindow, QApplication, QVBoxLayout, QWidget, QFileDialog, \
    QSpinBox, QLabel, QMessageBox, QListWidget, QComboBox, QToolBar, QAction, QAbstractButton
from PySide2.examples.widgets.itemviews.addressbook.tablemodel import TableModel

class PicButton(QAbstractButton):
    def __init__(self, pixmap, parent=None):
        super(PicButton, self).__init__(parent)
        self.pixmap = pixmap

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.drawPixmap(event.rect(), self.pixmap)

    def sizeHint(self):
        return QSize(10, 50)

class TableModel(QtCore.QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data  # worksheet
        #self.filter = filter
    def flags(self, index):
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

    def data(self, index, role):
        if role == Qt.DisplayRole:
            _value = self._data[index.row()][index.column()]
            return _value

    def setData(self, index, value, role):
        if role == Qt.EditRole:
            self._data[index.row()][index.column()] = value
            return True
        return False

    def rowCount(self, index):
        return len(self._data)

    def columnCount(self, index):
        return len(self._data[0])


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Opener V Sashka")
        self.setGeometry(300, 300, 800, 400)

        # self.ui = Ui_MainWindow()
        # self.ui.setupUi(self)
        self.file = None
        self.data = []
        self.book = None
        self.sheet = None

        self.mainToolbar = QToolBar("Main toolbar")
        self.addToolBar(self.mainToolbar)

        self.open_action = QAction("Open")
        self.open_action.setStatusTip("Open from computer")
        self.open_action.triggered.connect(self.open_button_clicked)

        self.save_action = QAction("Save")
        self.save_action.setStatusTip("Save document")
        self.save_action.triggered.connect(self.save_button_clicked)

        self.close_action = QAction("Close")
        self.close_action.setStatusTip("Close current document")
        self.close_action.triggered.connect(self.close_button_clicked)

        self.filter_action = QAction("Filters")
        self.filter_action.setStatusTip("Set some filters over your data")
        self.filter_action.triggered.connect(self.filter_button_clicked)

        self.mainToolbar.addAction(self.open_action)
        self.mainToolbar.addSeparator()
        self.mainToolbar.addAction(self.save_action)
        self.mainToolbar.addSeparator()
        self.mainToolbar.addAction(self.close_action)
        self.mainToolbar.addSeparator()
        self.mainToolbar.addAction(self.filter_action)

        self.sheetCount = QVBoxLayout()
        self.sheetCountLabel = QLabel("Worksheet")
        self.sheetListBox = QComboBox()
        self.sheetListBox.currentTextChanged.connect(self.sheetListBox_text_changed)
        self.sheetCount.addWidget(self.sheetCountLabel)
        self.sheetCount.addWidget(self.sheetListBox)

        self.fontChanger = QComboBox()
        self.fontChanger.addItems(["Calibri", "Arial", "Times New Roman"])

        self.fontSize = QSpinBox()
        self.fontSize.setMinimum(6)
        self.fontSize.setMaximum(20)

        #self.boldLetter = PicButton(QPixmap("boldLetter.png"))
        self.boldLetter = QPushButton()
        self.boldLetter.setIcon(QIcon("boldLetter.png"))

        self.curseLetter = QPushButton()
        self.curseLetter.setIcon(QIcon("cursevLetter.png"))

        self.hLayout = QHBoxLayout()
        self.hLayout.addWidget(self.fontChanger)
        self.hLayout.addWidget(self.fontSize)
        self.hLayout.addWidget(self.boldLetter)
        self.hLayout.addWidget(self.curseLetter)
        self.hLayout.addLayout(self.sheetCount)

        self.vLayout = QVBoxLayout()
        self.vLayout.addLayout(self.hLayout)

        self.table = QtWidgets.QTableView()
        #
        self.vLayout.addWidget(self.table)

        self.widget = QWidget()
        self.widget.setLayout(self.vLayout)
        self.setCentralWidget(self.widget)



    def open_button_clicked(self):
        self.file_name = QFileDialog.getOpenFileName(self)
        if self.file_name[0][-5:] == ".xlsx" or self.file_name[0][-4:] == ".xls":
            self.book = openpyxl.load_workbook(self.file_name[0])
            self.sheet = self.book.active
            self.data = [
                [[self.book[sheet][row][column].value for column in range(0, self.book[sheet].max_column)] for row in
                 range(1, self.book[sheet].max_row + 1)] for sheet in self.book.sheetnames]
            self.data.append([[[0, 0, 0]]])  # 0 - numbers, 1 - words, 2 - dates
            self.sheetListBox.addItems(self.book.sheetnames)

            # self.model = TableModel(self.data[0])
            # self.table.setModel(self.model)

    def close_button_clicked(self):
        qMessage = QMessageBox()
        qMessage.setWindowTitle("Question")
        qMessage.setText("Are you sure you want to close the file?")
        qMessage.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        qMessage.setIcon(QMessageBox.Question)

        if self.book is not None:
            button = qMessage.exec_()

            if button == QMessageBox.Yes:
                self.book.close()
                self.data = []
                self.sheetListBox.clear()
                self.table.setModel(None)

    def save_button_clicked(self):
        sheet_names = self.book.sheetnames
        sheet_name = self.sheet.title
        sheet_number = sheet_names.index(sheet_name)
        self.book.remove(self.book[sheet_name])
        self.book.create_sheet(sheet_name, sheet_number)
        current_sheet = self.book[sheet_name]

        for row in self.data[self.book.sheetnames.index(sheet_name)]:
            current_sheet.append(row)

        try:
            self.book.save(self.file_name[0])
        except Exception:
            falseMessage = QMessageBox.critical(self, "Error", "Unreal to save this file")
            falseMessage.exec_()
        else:
            okMessage = QMessageBox()
            okMessage.setWindowTitle("Ok")
            okMessage.setText("The Excel file was saved successfully")
            okMessage.setIcon(QMessageBox.Information)
            okMessage.exec_()

        self.book.close()

    def filter_button_clicked(self):
        self.w = QBoxWindow()
        self.w.show()
        #self.table.model().filter =

    def sheetListBox_text_changed(self, str):
        try:
            self.sheet = self.book[str]
            # self.data = [[self.sheet[row][column].value for column in range(0, self.sheet.max_column)] for row in
            #             range(1, self.sheet.max_row + 1)]
            self.model = TableModel(self.data[self.book.sheetnames.index(str)])
            self.table.setModel(self.model)
        except KeyError:
            o = 1

class QBoxWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Filters")
        self.setGeometry(500, 400, 300, 300)

        self.label = QLabel("Filters")

app = QApplication(sys.argv)
window = MainWindow()
window.show()
app.exec_()

import xlrd
import openpyxl
from excelSearch import Ui_MainWindow
from PyQt5.QtWidgets import QMainWindow, QApplication, QTableWidgetItem,QLabel,QPushButton, QFileDialog
# import pandas as pd
import openpyxl
# from openpyxl_image_loader import SheetImageLoader
import sys
# from Utils import read_excel
import time
from Utils import excel_search, sheet_selected


class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)
        self.fileAButton.clicked.connect(self.buttonA_pressed)
        self.fileBButton.clicked.connect(self.buttonB_pressed)
        self.startButton.clicked.connect(self.buttonStart_pressed)
        self.sheetNameACBox.currentIndexChanged.connect(self.sheet_selected)
    #     self.pushButton = QPushButton(self.centralwidget)
    #     self.pushButton.setText("Hello~")
    #     self.pushButton.clicked.connect(self.button_pressed)
    #


    def buttonA_pressed(self):
        # newLabel = QLabel()
        # newLabel.setText("hello ~~")
        # self.gridLayout.addWidget(newLabel)
        file, check = QFileDialog.getOpenFileName(None,
                                                       'Select file',
                                                       './',
                                                       'Excel Files (*.xlsx)')
        if check:
            print(file)
            file_name = file.split('/')[-1]
            self.fileNameA.setText(file_name)
            self.sheetNameACBox.addItems(excel_search(file))





    def buttonB_pressed(self):
        file, check = QFileDialog.getOpenFileName(None,
                                                  'Select file',
                                                  './',
                                                  'Excel Files (*.xlsx)')
        if check:
            print(file)
            file_name = file.split('/')[-1]
            self.fileNameB.setText(file_name)
            self.sheetNameBCBox.addItems(excel_search(file))

    def buttonStart_pressed():
        return

    def sheet_selected(self):

        return

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyMainWindow()
    ex.show()
    sys.exit(app.exec_())


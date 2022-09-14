import xlrd
import openpyxl
from excelSearch import Ui_MainWindow
from PyQt5.QtWidgets import QMainWindow, QApplication, QTableWidgetItem,QLabel,QPushButton, QFileDialog, QMessageBox, QErrorMessage
# import pandas as pd
import openpyxl
# from openpyxl_image_loader import SheetImageLoader
import sys
# from Utils import read_excel
import time
from Utils import sheet_search, buttonStart_pressed2, result_file_list
from openpyxl.utils import get_column_letter

class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)
        self.fileAButton.clicked.connect(self.buttonA_pressed)
        self.fileBButton.clicked.connect(self.buttonB_pressed)
        self.startButton.clicked.connect(self.buttonStart_pressed)

        self.row1CBoxA.addItems([str(r1a+1) for r1a in range(100)])
        self.row2CBoxA.addItems([str(r2a+1) for r2a in range(100)])
        self.column1CBoxA.addItems([get_column_letter(c1a+1) for c1a in range(100)])
        self.column2CBoxA.addItems([get_column_letter(c2a+1) for c2a in range(100)])

        self.row1CBoxB.addItems([str(r2b+1) for r2b in range(100)])
        self.row2CBoxB.addItems([str(r2b+1) for r2b in range(100)])
        self.column1CBoxB.addItems([get_column_letter(r2b+1) for r2b in range(100)])
        self.column2CBoxB.addItems([get_column_letter(r2b+1) for r2b in range(100)])

        # 테스트용
        self.row1CBoxA.setCurrentText("2")
        self.row2CBoxA.setCurrentText("13")
        self.column1CBoxA.setCurrentText("B")
        self.column2CBoxA.setCurrentText("I")
        self.row1CBoxB.setCurrentText("8")
        self.row2CBoxB.setCurrentText("19")
        self.column1CBoxB.setCurrentText("E")
        self.column2CBoxB.setCurrentText("O")

        # self.row1CBoxA.currentIndexChanged.connect(self.selectionChanged)
        # self.row2CBoxA.currentIndexChanged.connect(self.selectionChanged)
        # self.column1CBoxA.currentIndexChanged.connect(self.selectionChanged)
        # self.column2CBoxA.currentIndexChanged.connect(self.selectionChanged)
        #
        # self.row1CBoxB.currentIndexChanged.connect(self.selectionChanged)
        # self.row2CBoxB.currentIndexChanged.connect(self.selectionChanged)
        # self.column1CBoxB.currentIndexChanged.connect(self.selectionChanged)
        # self.column2CBoxB.currentIndexChanged.connect(self.selectionChanged)

    def buttonA_pressed(self):
        self.file1, check = QFileDialog.getOpenFileName(None,
                                                       'Select file',
                                                       './',
                                                       'Excel Files (*.xlsx)')
        if check:
            # print(self.file1)
            file_name = self.file1.split('/')[-1]
            self.fileNameA.setText(file_name)
            self.sheetNameACBox.addItems(sheet_search(self.file1))

    def buttonB_pressed(self):
        self.file2, check = QFileDialog.getOpenFileName(None,
                                                  'Select file',
                                                  './',
                                                  'Excel Files (*.xlsx)')
        if check:
            # print(self.file2)
            file_name = self.file2.split('/')[-1]
            self.fileNameB.setText(file_name)
            self.sheetNameBCBox.addItems(sheet_search(self.file2))


    # def selectionChanged(self):
    #     txt = self.cbo.currentText()
    #     idx = self.cbo.currentIndex()
    def buttonStart_pressed(self):
        try:
            # print("file1: ",self.file1)
            # print("file2: ",self.file2)
            print("실행 중..")
            buttonStart_pressed2(self.file1, self.file2, self.sheetNameACBox.currentText(),self.sheetNameBCBox.currentText(), self.row1CBoxA.currentText(), self.row2CBoxA.currentText(), self.column1CBoxA.currentText(), self.column2CBoxA.currentText(),
                                 self.row1CBoxB.currentText(), self.row2CBoxB.currentText(), self.column1CBoxB.currentText(), self.column2CBoxB.currentText())
            QMessageBox.about(self, "완료", "실행이 완료되었습니다.\n" + result_file_list[0])
        except:
            print("start error")
            # alert

            QErrorMessage.showMessage(self,'실행에 실패하였습니다.')


        return




if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyMainWindow()
    ex.show()
    sys.exit(app.exec_())


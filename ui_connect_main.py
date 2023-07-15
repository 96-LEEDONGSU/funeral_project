import sys
import typing
from PyQt5 import QtCore, uic
from PyQt5.QtWidgets import QFileDialog, QApplication, QMainWindow
import excel_analysis_package


form_class = uic.loadUiType("funeral_project/MainWindow.ui")[0]


class WindowClass(QMainWindow, form_class):
    select_dir_name = ''
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        
        self.btn_excel_analysis.clicked.connect(self.excel_analysis)
        self.btn_select_dir.clicked.connect(self.select_dir)
    
    def select_dir(self):
        self.select_dir_name = QFileDialog.getExistingDirectory(self, '폴더 선택', '')
        self.lbl_select_dir.setText(self.select_dir_name)
        self.select_dir_name += '/'
        print(self.select_dir_name)
        
    def excel_analysis(self):
        print(self.select_dir_name)
        if self.select_dir_name == '':
            print('선택안됨')
        else:
            print('선택됨')
            excel_analysis_package.excel_analysis(self.select_dir_name)
        #self.setEnabled(False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()
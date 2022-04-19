from ctypes import resize
import sys

from PyPDF2 import PdfFileReader, PdfFileWriter

from PySide6.QtSerialPort import *
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('PDF Merger')
        self.setMinimumWidth(700)
        self.init_user_interface()
        self.init_signal_slot()

    # Initialize User Interace
    def init_user_interface(self):
        self.title_label = QLabel('Merge PDF Files')
        self.title_label.setFont(QFont('Sans', 15, QFont.Bold))
        self.description_label = QLabel('Combine PDFs in the order you want with the easiest PDF merger')

        self.vertical_spacer = QSpacerItem(150, 10, QSizePolicy.Expanding)

        self.open_button = QPushButton('Select PDF Files')
        self.open_button.setMinimumHeight(50)
        self.open_button.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))

        self.file_list = QListWidget()
        self.file_list.setVisible(False)
        self.file_list.setMinimumHeight(200)
        self.file_list.setStyleSheet('QListWidget::item { height: 40px; }')
        self.file_list.setVisible(False)

        self.up_button = QPushButton('Up')
        self.up_button.setMinimumHeight(30)
        self.up_button.setVisible(False)

        self.down_button = QPushButton('Down')
        self.down_button.setMinimumHeight(30)
        self.down_button.setVisible(False)

        self.add_button = QPushButton('Add')
        self.add_button.setMinimumHeight(30)
        self.add_button.setVisible(False)

        self.remove_button = QPushButton('Remove')
        self.remove_button.setMinimumHeight(30)
        self.remove_button.setVisible(False)

        self.merge_button = QPushButton('Merge PDF')
        self.merge_button.setMinimumHeight(50)
        self.merge_button.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        self.merge_button.setVisible(False)

        self.back_button = QPushButton('Back')
        self.back_button.setMinimumHeight(50)
        self.back_button.setVisible(False)

        self.central_layout = QGridLayout()
        self.central_layout.setContentsMargins(50, 50, 50, 50)
        self.central_layout.addWidget(self.title_label, 0, 0, 1, 4)
        self.central_layout.addWidget(self.description_label, 1, 0, 1, 4)
        self.central_layout.addItem(self.vertical_spacer, 2, 0, 1, 4)
        self.central_layout.addWidget(self.open_button, 3, 0, 1, 2)
        self.central_layout.addWidget(self.file_list, 4, 0, 1, 4)
        self.central_layout.addWidget(self.up_button, 5, 0, 1, 1)
        self.central_layout.addWidget(self.down_button, 5, 1, 1, 1)
        self.central_layout.addWidget(self.add_button, 5, 2, 1, 1)
        self.central_layout.addWidget(self.remove_button, 5, 3, 1, 1)
        self.central_layout.addWidget(self.back_button, 6, 0, 1, 1)
        self.central_layout.addWidget(self.merge_button, 6, 1, 1, 3)

        self.central_widget = QWidget()
        self.central_widget.setLayout(self.central_layout)
        self.setCentralWidget(self.central_widget)
    
    # Initialize Signal and Slot
    def init_signal_slot(self):
        self.open_button.clicked.connect(self.open_files)
        self.merge_button.clicked.connect(self.merge_pdf)
        self.back_button.clicked.connect(self.back)

        self.up_button.clicked.connect(self.move_up)
        self.down_button.clicked.connect(self.move_down)
        self.add_button.clicked.connect(self.add_item)
        self.remove_button.clicked.connect(self.remove_item)

    # Function to Get Path of the Files
    def open_files(self):
        filenames, *_ = QFileDialog.getOpenFileNames(self, 'Open Files', QDir.currentPath(), 'PDF Files (*.pdf)')
        if filenames:
            self.open_button.setVisible(False)
            self.file_list.setVisible(True)
            self.file_list.addItems(filenames)
            self.up_button.setVisible(True)
            self.down_button.setVisible(True)
            self.add_button.setVisible(True)
            self.remove_button.setVisible(True)
            self.back_button.setVisible(True)
            self.merge_button.setVisible(True)
            self.setFixedSize(self.central_layout.sizeHint())
            self.setMinimumWidth(700)
    
    # Function to Merge PDFs
    def merge_pdf(self):
        paths = []
        for index in range(self.file_list.count()):
            paths.append(self.file_list.item(index).text())
        
        pdf_writer = PdfFileWriter()

        for path in paths:
            pdf_reader = PdfFileReader(path)
            for page in range(pdf_reader.getNumPages()):
                pdf_writer.addPage(pdf_reader.getPage(page))
        
        filename, *_ = QFileDialog.getSaveFileName(self, 'Save File', QDir.currentPath(), 'PDF File (*.pdf)')
        
        with open(filename, 'wb') as out:
            pdf_writer.write(out)
    
    # Move Up Item in File List
    def move_up(self):
        row_index = self.file_list.currentRow()
        current_item = self.file_list.takeItem(row_index)
        self.file_list.insertItem(row_index - 1, current_item)
        self.file_list.setCurrentRow(row_index - 1)

    # Move Down Item in File List
    def move_down(self):
        row_index = self.file_list.currentRow()
        current_item = self.file_list.takeItem(row_index)
        self.file_list.insertItem(row_index + 1, current_item)
        self.file_list.setCurrentRow(row_index + 1)

    # Add Item to File List
    def add_item(self):
        filenames, *_ = QFileDialog.getOpenFileNames(self, 'Open Files', QDir.currentPath(), 'PDF Files (*.pdf)')
        self.file_list.addItems(filenames)
    
    # Remove Item in the File List
    def remove_item(self):
        self.file_list.takeItem(self.file_list.currentRow())
    
    def back(self):
        self.open_button.setVisible(True)
        self.file_list.setVisible(False)
        self.file_list.clear()
        self.up_button.setVisible(False)
        self.down_button.setVisible(False)
        self.add_button.setVisible(False)
        self.remove_button.setVisible(False)
        self.back_button.setVisible(False)
        self.merge_button.setVisible(False)
        self.setFixedSize(self.central_layout.sizeHint())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = MainWindow()
    main.show()
    app.exec()
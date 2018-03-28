# ===================================
__software__ = 'Pdf-file splitter for suppliers'
__author__ = 'Michael Belyansky'
__copyright__ = 'Metro Cash and Carry Ukraine, 2018' + chr(169)
__license__ = 'GNU GPL v3'
__version__ = '1.0.1'
__maintainer__ = 'Michael Belyansky'
__email__ = 'mykhaylo.belyanskiy@metro.ua'
__status__ = 'Production'
# ===================================


import sys
import PyPDF2
import datetime
import act_pdf_splitter_rc
from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_mainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
    def setupUi(self):
        self.setObjectName('mainWindow')
        self.setFixedSize(450, 165)
        self.setWindowTitle('Reconciliation act (pdf-file splitter for suppliers)')
        self.setWindowFlags(QtCore.Qt.MSWindowsFixedSizeDialogHint)
        font_bold = QtGui.QFont()
        font_bold.setBold(True)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(':/Icons/icon.ico'), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName('centralwidget')
        # ===================================
        self.pdf_name = QtWidgets.QLineEdit(self.centralwidget)
        self.pdf_name.setGeometry(QtCore.QRect(10, 10, 300, 20))
        self.pdf_name.setReadOnly(True)
        self.pdf_name.setObjectName('pdf_name')
        self.pdf_name.setFont(font_bold)
        # ===================================
        self.open_file_button = QtWidgets.QPushButton(self.centralwidget)
        self.open_file_button.setGeometry(QtCore.QRect(320, 10, 120, 20))
        self.open_file_button.setText('Pdf-file')
        self.open_file_button.setObjectName('open_file_button')
        self.open_file_button.setFont(font_bold)
        self.open_file_button.clicked.connect(self.take_pdf_file)
        # ===================================
        self.export_directory = QtWidgets.QLineEdit(self.centralwidget)
        self.export_directory.setGeometry(QtCore.QRect(10, 40, 300, 20))
        self.export_directory.setReadOnly(True)
        self.export_directory.setObjectName('export_directory')
        self.export_directory.setFont(font_bold)
        self.export_directory.setText('\\\kiv11appp07.ua.r4.madm.net\\Accounting\\In')
        # ===================================
        self.open_directory_button = QtWidgets.QPushButton(self.centralwidget)
        self.open_directory_button.setGeometry(QtCore.QRect(320, 40, 120, 20))
        self.open_directory_button.setText('Export directory')
        self.open_directory_button.setObjectName('open_directory_button')
        self.open_directory_button.setFont(font_bold)
        self.open_directory_button.clicked.connect(self.take_export_directory)
        # ===================================
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(10, 75, 430, 23))
        self.progressBar.setFormat('processed pdf-file %p%')
        self.progressBar.setProperty('value', 0)
        self.progressBar.setMaximum(100)
        self.progressBar.setAlignment(QtCore.Qt.AlignCenter)
        self.progressBar.setObjectName('progressBar')
        self.progressBar.setFont(font_bold)
        self.progressBar.hide()
        # ===================================
        self.label_help = QtWidgets.QLabel(self.centralwidget)
        self.label_help.setGeometry(QtCore.QRect(10, 75, 381, 23))
        self.label_help.setText('Please choose pdf-file, export directory and press "Split" button')
        self.label_help.setFont(font_bold)
        # ===================================
        self.split_button = QtWidgets.QPushButton(self.centralwidget)
        self.split_button.setGeometry(QtCore.QRect(10, 110, 75, 23))
        self.split_button.setText('&Split')
        self.split_button.setObjectName('split_button')
        self.split_button.setFont(font_bold)
        self.split_button.clicked.connect(self.split_pdf_file)
        # ===================================
        self.close_button = QtWidgets.QPushButton(self.centralwidget)
        self.close_button.setGeometry(QtCore.QRect(100, 110, 75, 23))
        self.close_button.setText('&Exit')
        self.close_button.setObjectName('close_button')
        self.close_button.setFont(font_bold)
        self.close_button.clicked.connect(self.close)
        # ===================================
        self.label_qty = QtWidgets.QLabel(self.centralwidget)
        self.label_qty.setGeometry(QtCore.QRect(190, 110, 250, 23))
        self.label_qty.setText('Processed: 1 supplier(s)')
        self.label_qty.setFont(font_bold)
        self.label_qty.hide()
        # ===================================
        self.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(self)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 400, 21))
        self.menubar.setObjectName('menubar')
        self.menubar.setFont(font_bold)
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName('menuFile')
        self.menuFile.setTitle('&File')
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName('menuHelp')
        self.menuHelp.setTitle('&Help')
        self.setMenuBar(self.menubar)
        
        self.actionAbout = QtWidgets.QAction(self, shortcut='Ctrl+H', triggered=self.about)
        self.actionAbout.setObjectName('actionAbout')
        self.actionAbout.setText('&About')
        self.actionAbout.setShortcut('Ctrl+H')

        self.actionExit = QtWidgets.QAction(self, shortcut='Ctrl+E', triggered=self.close)
        self.actionExit.setObjectName('actionExit')
        self.actionExit.setText('&Exit')
        self.actionExit.setShortcut('Ctrl+E')
        
        self.menuFile.addAction(self.actionExit)
        self.menuHelp.addAction(self.actionAbout)
        
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        # ===================================
        QtCore.QMetaObject.connectSlotsByName(self)
        QtWidgets.QApplication.setStyle(QtWidgets.QStyleFactory.create('Fusion'))

    def closeEvent(self, event):
        close = QtWidgets.QMessageBox(self)
        close.setText('You are sure to exit program?')
        close.setWindowTitle('Exit program')
        close.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        close = close.exec()
        if close == QtWidgets.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def about(self):
        about_txt = '<HTML>'\
            '<p style="font-size:16px; color:#000099"><b>Reconciliation act (pdf-file splitter for suppliers)</b></p>'\
            '<b>This program is provided to split pdf file from SAP</b><br>'\
            '<b>and copy pdf-file in processed folder for METRO-link for each supplier</b><br>'\
            '<b>---------------------------------------------------------</b><br>' +\
            __copyright__ + '<br>'\
            'author: ' + __author__ + '<br>'\
            'e-mail: <a href="mailto:' + __email__ + '?subject=Reconciliation act splitter">' + __email__ + '</a><br>'\
            'licence: ' + __license__ + '<br>'\
            'version: '+ __version__ + ' <br>'\
            '<b>---------------------------------------------------------</b><br>'\
            'Reconciliation act (pdf-file splitter for suppliers) developed in Python v. 3.6.3<br>'\
            '<a href="https://www.python.org/">https://www.python.org/</a><br><br>'\
            '<u><b>Used components and resources:</b></u><br>'\
            'PyQt GPL v 5.10.1 for Python v3.6.3 (x32):<br>'\
            'Compile with pyinstaller module<br>'\
            'Silk icon set 1.3 by Mark James:<br>'\
            '<a href="http://www.famfamfam.com/lab/icons/silk/">http://www.famfamfam.com/lab/icons/silk</a><br>'\
            '<b>---------------------------------------------------------</b><br>'\
            '</HTML>'
        about_box = QtWidgets.QMessageBox(self)
        about_box.setWindowTitle('Pdf-file splitter')
        about_box.setTextFormat(1)
        about_box.setText(about_txt)
        about_box.exec_()

    # ===================================
    def take_pdf_file(self):
        open_pdf_file = QtWidgets.QFileDialog.getOpenFileName(self.centralwidget, 'Open pdf-file:', self.pdf_name.text(), 'Adobe pdf-files (*.pdf)')
        if open_pdf_file:
            self.pdf_name.setText(open_pdf_file[0].replace('/', '\\'))

    def take_export_directory(self):
        export_directory = QtWidgets.QFileDialog.getExistingDirectory(self.centralwidget, 'Open export directory:', self.export_directory.text())
        if export_directory:
            self.export_directory.setText(export_directory.replace('/', '\\'))
    
    def split_pdf_file(self):
        if not self.pdf_name.text():
            error_msg = QtWidgets.QMessageBox(self)
            error_msg.setIcon(QtWidgets.QMessageBox.Warning)
            error_msg.setText('For processing please choose pdf-file!')
            error_msg.setWindowTitle('Error split pdf-file')
            error_msg.exec_()
            return
        if not self.export_directory.text():
            error_msg = QtWidgets.QMessageBox(self)
            error_msg.setIcon(QtWidgets.QMessageBox.Warning)
            error_msg.setText('For processing please choose export directory!')
            error_msg.setWindowTitle('Error split pdf-file')
            error_msg.exec_()
            return
        self.progressBar.show()
        self.label_help.hide()
        self.label_qty.show()
        try:
            pdfFileObj = open(self.pdf_name.text(), 'rb')
        except Exception:
            error_msg = QtWidgets.QMessageBox(self.centralwidget)
            error_msg.setIcon(QtWidgets.QMessageBox.Warning)
            error_msg.setText('Error open pdf-file!')
            error_msg.setWindowTitle('Error split pdf-file')
            error_msg.exec_()
            return
        
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        self.progressBar.setRange(0, pdfReader.numPages - 1)
        maxVal = self.progressBar.maximum()   
        writer = PyPDF2.PdfFileWriter()
        qty_suppl = 0
        for i in range (0, pdfReader.numPages):
           self.progressBar.setValue(i + (maxVal - i) / 100)
           QtWidgets.QApplication.processEvents()
           pageObj = pdfReader.getPage(i)
           check_text = pageObj.extractText().split()
           if len(check_text) > 5 and check_text[1] == '10.11.98':
               if i == 0:
                   suppl_no = list(filter(lambda f: len(f) == 10 and f[:2] == '26', check_text))[0][-5:]
                   writer.addPage(pdfReader.getPage(i))
               else:
                   output_file = '\\Reconciliation_act~Acc_act~60~' + suppl_no + '~' + datetime.date.today().strftime('%Y%m%d') + '.pdf'
                   try:
                       with open(self.export_directory.text() + output_file, 'wb') as outfile:
                           writer.write(outfile)
                           qty_suppl += 1
                           self.label_qty.setText('Processed: ' + str(qty_suppl) + ' supplier(s)')
                           QtWidgets.QApplication.processEvents()
                   except Exception:
                       error_msg = QtWidgets.QMessageBox(self.centralwidget)
                       error_msg.setIcon(QtWidgets.QMessageBox.Warning)
                       error_msg.setText('Error save pdf-file!')
                       error_msg.setWindowTitle('Error split pdf-file')
                       error_msg.exec_()
                       return
                   outfile.close()
                   suppl_no = list(filter(lambda f: len(f) == 10 and f[:2] == '26', check_text))[0][-5:]
                   writer = PyPDF2.PdfFileWriter()
                   writer.addPage(pdfReader.getPage(i))
           else:
               if len(check_text) > 5:
                   writer.addPage(pdfReader.getPage(i))
        output_file = '\\Reconciliation_act~Acc_act~60~' + suppl_no + '~' + datetime.date.today().strftime('%Y%m%d') + '.pdf'
        try:
            with open(self.export_directory.text() + output_file, 'wb') as outfile:
                writer.write(outfile)
                outfile.close()
                qty_suppl += 1
                self.label_qty.setText('Processed: ' + str(qty_suppl) + ' supplier(s)')                                  
        except Exception:
            error_msg = QtWidgets.QMessageBox(self.centralwidget)
            error_msg.setIcon(QtWidgets.QMessageBox.Warning)
            error_msg.setText('Error save pdf-file!')
            error_msg.setWindowTitle('Error split pdf-file')
            error_msg.exec_()
            return
        success_msg = QtWidgets.QMessageBox(self)
        success_msg.setIcon(QtWidgets.QMessageBox.Information)
        success_msg.setText('Processed: ' + str(qty_suppl) + ' supplier(s)')
        QtWidgets.QApplication.processEvents()
        success_msg.setWindowTitle('Success split pdf-file')
        success_msg.exec_()
        self.progressBar.hide()
        self.label_help.show()
        

if __name__ == '__main__':

    app = QtWidgets.QApplication(sys.argv)
    mainWindow = Ui_mainWindow()
    mainWindow.setupUi()
    mainWindow.show()
    sys.exit(app.exec_())


import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from os import listdir
import os, codecs, openpyxl
import speech_recognition as sr
from pydub import AudioSegment


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")

        MainWindow.resize(640, 480)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 110, 131, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(20, 170, 131, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")

        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(160, 120, 331, 31))
        self.lineEdit.setObjectName("lineEdit")

        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(160, 180, 331, 31))
        self.lineEdit_2.setObjectName("lineEdit_2")

        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(160, 230, 131, 41))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_2")

        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(160, 280, 121, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.submit)

        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(350, 280, 141, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(QtCore.QCoreApplication.instance().quit)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 640, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "Path to Media: "))
        self.label_2.setText(_translate("MainWindow", "Path to Rawdata: "))
        self.pushButton.setText(_translate("MainWindow", "Submit"))
        self.pushButton_2.setText(_translate("MainWindow", "Exit"))

    def submit(self, MainWindow):
        mp3_path = os.path.normpath(self.lineEdit.text())
        excel_path = os.path.normpath(self.lineEdit_2.text())
        wit_ai_key = "W4V5LK7TMXE7RQQ3CU42P2FBY5SY23CH"

        files = os.listdir(mp3_path)
        mpeg4_files = [x for x in files if '.mpeg4' in x]

        wb1 = openpyxl.load_workbook(excel_path, keep_vba=True)
        wb2 = openpyxl.Workbook()
        sheet1 = wb1["Sheet1"]
        sheet2 = wb2.active
        sheet2.cell(row=1, column=1).value = "Respondent Serial No"
        sheet2.cell(row=1, column=2).value = "Centre"
        sheet2.cell(row=1, column=3).value = "Respondent Name"
        sheet2.cell(row=1, column=4).value = "Language"

        row_num = 2
        j = True
        while (j):
            intnr = int(sheet1.cell(row=row_num, column=1).value)
            sheet2.cell(row=row_num, column=1).value = intnr
            sheet2.cell(row=row_num, column=2).value = sheet1.cell(row=row_num, column=2).value
            sheet2.cell(row=row_num, column=3).value = sheet1.cell(row=row_num, column=5).value

            file_name = [x for x in mpeg4_files if intnr == int(x[:8])]
            for u, v in enumerate(file_name):
                file = v.replace(".mpeg4", ".wav")
                song = AudioSegment.from_file(os.path.join(mp3_path,v))
                song.export(file, format="wav")

                r = sr.Recognizer()
                with sr.AudioFile(file) as source:
                    # r.adjust_for_ambient_noise(source)
                    audio_text = r.listen(source)

                #try:
                    text = r.recognize_wit(audio_text, key=wit_ai_key)
                #except sr.UnknownValueError:
                    text = "Could not understand audio"
                #except sr.RequestError as e:
                    text = "Could not request results from Wit.ai service; {0}".format(e)

                if row_num == 2:
                    sheet2.cell(row=1, column=u * 2 + 5).value = "File Name"
                    sheet2.cell(row=1, column=u * 2 + 6).value = "Text"
                    sheet2.cell(row=row_num, column=u * 2 + 5).value = v
                    sheet2.cell(row=row_num, column=u * 2 + 5).hyperlink = os.path.join(mp3_path, v)
                    sheet2.cell(row=row_num, column=u * 2 + 6).value = text
                else:
                    sheet2.cell(row=row_num, column=u * 2 + 5).value = v
                    sheet2.cell(row=row_num, column=u * 2 + 5).hyperlink = os.path.join(mp3_path, v)
                    sheet2.cell(row=row_num, column=u * 2 + 6).value = text

                os.remove(file)

            row_num += 1
            if sheet1.cell(row=row_num, column=1).value == None:
                j = 0

        message = "Completed"

        wb2.save(os.path.join(mp3_path,"Output.xlsx"))
        wb1.close()

        self.msg = QtWidgets.QMessageBox(self.centralwidget)
        self.msg.setIcon(QtWidgets.QMessageBox.Information)

        self.msg.setText(message)
        self.msg.setWindowTitle("Status")
        self.msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        self.msg.exec_()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

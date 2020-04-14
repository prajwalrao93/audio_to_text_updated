"# audio_to_text_updated" 
# -*- coding: utf-8 -*-
# Form implementation generated from reading ui file 'speech_to_text.ui'
#
# Created by: PyQt5 UI code generator 5.9.2

import sys
from PyQt5 import QtCore, QtGui, QtWidgets
import os, langid
from os import listdir
import speech_recognition as sr
from pydub import AudioSegment
from wit import WIT

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        
        MainWindow.resize(640, 480)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 110, 131, 41))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.label.setFont(font)
        self.label.setObjectName("label")
        
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(20, 170, 131, 41))
        font = QtGui.QFont()
        font.setPointSize(16)
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
        self.label_2.setText(_translate("MainWindow", "Path to Excel: "))
        self.pushButton.setText(_translate("MainWindow", "Submit"))
        self.pushButton_2.setText(_translate("MainWindow", "Exit"))



    def submit(self,MainWindow):
        
        mp3_path = os.path.normpath(self.lineEdit.text())
        excel_path = os.path.normpath(self.lineEdit_2.text())
        files = os.listdir(mp3_path)

        mp3_files = [x for x in files if ".mp3" in x]

        for file_name in mp3_files:
            file = os.path.join(mp3_path, file_name)
            song = AudioSegment.from_mp3(file)
            file = file.replace(".mp3", ".wav")
            song.export(file, format="wav")
            
            r = sr.Recognizer()
            file_audio = sr.AudioFile(file)
            with file_audio as source:
                r.adjust_for_ambient_noise(source)
                audio_text = r.record(source)
                
            WIT_AI_KEY = "JS3HNTZH22MFTKO24UA5QRMELMPDMAUO" # Wit.ai keys are 32-character uppercase alphanumeric strings
            try:
                text = r.recognize_sphinx(audio_text)
                lang = langid.classify(text)[0]
                print(lang)
                text = r.recognize_wit(audio_text, key=WIT_AI_KEY)

                txt_file = file.replace(".wav", ".txt")
                with open(txt_file, 'w') as out:
                    out.write(text)

                os.remove(file)
            except sr.UnknownValueError:
                print("Wit.ai could not understand audio")
            except sr.RequestError as e:
                print("Could not request results from Wit.ai service; {0}".format(e))
                

        self.msg = QtWidgets.QMessageBox(self.centralwidget)
        self.msg.setIcon(QtWidgets.QMessageBox.Information)

        self.msg.setText("Completed")
        self.msg.setWindowTitle("Completed")
        self.msg.setStandardButtons(QtWidgets.QMessageBox.Ok)       
        self.msg.exec_()
        


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

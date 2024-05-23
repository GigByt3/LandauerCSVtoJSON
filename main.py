import json
import sys
import os
import logging

import pandas as pd
import tabula
import jpype
from tabulate import tabulate

from PyQt6 import QtGui, QtCore
from PyQt6.QtWidgets import QMessageBox, QApplication, QLabel, QMainWindow, QPushButton, QLineEdit, QVBoxLayout, QWidget, QFileDialog

from PyQt6.QtGui import QGuiApplication, QWindow
from PyQt6.QtQml import QQmlApplicationEngine
from PyQt6.QtQuick import QQuickWindow
from PyQt6.QtCore import Qt


class Window(QMainWindow):

    line_edit=""

    def __init__(self):
        super().__init__()
        self.setMinimumSize(100,100)

        msgBox = QMessageBox();
        msgBox.setText(f"tabula java connection {tabula.environment_info()}");
        msgBox.exec();
        
        btn = QPushButton("Select a Landauer Report PDF")
        aside = QLabel("'Landauer Report PDFs' in this case refer to \n documents generated by logging into your Landauer Portal, \n selecting a dose report, and selecting print.")
        saveFileInstructions = QLabel("Please Choose a Save File Name: \n default is sampleLandauer, do not include a file extension or \n special charecters.")
        saveFileName = QLineEdit(parent=self)
        Window.line_edit=saveFileName

        layout = QVBoxLayout()
        layout.addWidget(saveFileInstructions)
        layout.addWidget(saveFileName)
        layout.addWidget(btn)
        layout.addWidget(aside)

        guiDisplay = QWidget()
        guiDisplay.setLayout(layout)

        self.setCentralWidget(guiDisplay)
        self.setMinimumSize(400,100)
        self.setWindowTitle("Landauer Report Processor")

        btn.clicked.connect(self.clickHandler)

        # set app icon    
        base_dir = os.path.dirname(__file__)
        file_path = os.path.join(base_dir, '.GUI/icons/16x16.png')
        app_icon = QtGui.QIcon()
        app_icon.addFile(os.path.join(base_dir, '.GUI/icons/16x16.png'), QtCore.QSize(16,16))
        app_icon.addFile(os.path.join(base_dir, '.GUI/icons/24x24.png'), QtCore.QSize(24,24))
        app_icon.addFile(os.path.join(base_dir, '.GUI/icons/32x32.png'), QtCore.QSize(32,32))
        app_icon.addFile(os.path.join(base_dir, '.GUI/icons/48x48.png'), QtCore.QSize(48,48))
        app_icon.addFile(os.path.join(base_dir, '.GUI/icons/256x256.png'), QtCore.QSize(256,256))
        app.setWindowIcon(app_icon)

    def clickHandler(self):
        #logger = logging.getLogger("__log__")
        #logging.basicConfig(filename='myapp.log', level=logging.INFO)

        dialog = QFileDialog(self)
        dialog.setNameFilter("PDF Files (*.pdf)")
        dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        dialogSucessful = dialog.exec()

        if dialogSucessful == 1:

            selectedFiles = dialog.selectedFiles()

            msgBox = QMessageBox();
            msgBox.setText(f"filepath {selectedFiles[0]}");
            msgBox.exec();

            #bottom is ACTUALY 83.07
            try:
                allTables=tabula.read_pdf(selectedFiles[0], pages="all", multiple_tables=True, relative_columns=True, relative_area=True, area=(29.41,2.66,87.07,97.59), columns=[2.66, 6.17, 20.79, 24.92, 30.58, 32.94, 34.80, 38.57, 42.60, 47.15, 51.10, 55.09, 59.64, 63.48, 67.48, 71.96, 76.45, 81.58, 86.29, 90.26, 97.59])
            except Exception as e:
                error = e
                msgBox = QMessageBox();
                msgBox.setText(f"Failed Making All Tables! {e}");
                msgBox.exec();

            #recording set to waiting as we iterate down the sheet, primed once we hit the first monitering period, recording once we hit the second, and stopped once we hit the third.
            # 0 -- waiting
            # 1 -- primed
            # 2 -- recording
            # 3 -- stopped

            try:
                lastPeriod=tabula.read_pdf(selectedFiles[0], pages=1, relative_area=True, area=[29.35, 34.72, 78.00, 47.11])
                oldPeriod = lastPeriod[0].iat[2,0]
                recording=0
                sheetNumber=1
                allPersonel = []
            except Exception as e:
                error = e
                msgBox = QMessageBox();
                msgBox.setText(f"Failed Making All Tables! {e}");
                msgBox.exec();

            msgBox = QMessageBox();
            msgBox.setText("Sorting File...");
            msgBox.exec();

            for iterator in range(0,len(allTables)):
                
                msgBox = QMessageBox();
                msgBox.setText(f"In Loop... {iterator+1}/13");
                msgBox.exec();

                p = allTables[iterator]

                msgBox = QMessageBox();
                msgBox.setText("Set iterator...");
                msgBox.exec();

                try:
                    checkPeriod=tabula.read_pdf(selectedFiles[0], pages=sheetNumber, relative_area=True, area=[29.35, 34.72, 78.00, 47.11])
                except FileNotFoundError as fE:
                    oops = "oops"
                    msgBox = QMessageBox();
                    msgBox.setText("Error: sheetNumber overflow on line 110.");
                    msgBox.exec();

                mDLP = 0;
                monitoringDateList=[]
                for r in range(checkPeriod[0].shape[0]):
                    if " to " in str(checkPeriod[0].iat[r,0]):
                        monitoringDateList.append(str(checkPeriod[0].iat[r,0]).strip().replace("D", "").replace("E", "").replace("r", "").replace("\\", "").replace("L", "").replace("S", ""))
                
                #print(monitoringDateList)
                sheetNumber = sheetNumber + 1
                tableShape = p.shape

                #print(str(tableShape))
                for r in range(tableShape[0]):
                    #print(str(p.iat[r, 7]) + ", 7 - " + str(p.iat[r, 8]) + ", 8 - " + str(p.iat[r, 9]) + ", 9 - " + str(p.iat[r, 10]) + ", 10 - " + str(p.iat[r, 11]) + ", 11 - " + str(p.iat[r, 12]) + ", 12")
                    for c in range(tableShape[1]):
                        if "For M" in str(p.iat[r, c]):
                            if str(monitoringDateList[mDLP]).strip() != str(oldPeriod).strip():
                                #print("New Date: " + str(monitoringDateList[mDLP]).strip() + "\n Old Date: " + str(oldPeriod).strip())
                                recording = recording+1
                                oldPeriod = monitoringDateList[mDLP]
                            else:
                                #print("Check Date: " + str(monitoringDateList[mDLP]).strip() + " same as\n Old Date: " + str(oldPeriod).strip())
                                doNothing = "donothing"
                            mDLP=mDLP+1

                        #Check if your in the right reporting period
                        if recording == 1:
                            #Check if you've found a personel line.
                            if str(p.iat[r,c]) != "nan" and str(p.iat[r,c]) != "For M" and str(p.iat[r,c]).lstrip("0") != "" and c == 1:
                                #print("found personel number: " + str(p.iat[r,c]).lstrip("0"))
                                personelNumber = str(p.iat[r,c]).lstrip("0")
                                personelDict = {"number": personelNumber}
                                pCnt = 0
                                #Record Data!
                                while pCnt == 0 or str(p.iat[r+pCnt,c]) == "nan" or str(p.iat[r+pCnt,c]) == "For M" or str(p.iat[r+pCnt,c]).lstrip("0") == "":
                                    if r+pCnt+1 < p.shape[0]:
                                        #print("pCnt: " + str(pCnt+1) + ", bellow item: " + str(p.iat[r+pCnt+1,c]))
                                        doNothing="do nothing"
                                    #Check for dataframe overflow
                                    if r+pCnt+1 < tableShape[0]:
                                        #Check if Absent?
                                        if p.iat[r+pCnt+1,4] == "NOTE":
                                            wasUsed = True
                                            for i in range(10):
                                                    if str(p.iat[r+pCnt+1,i]).strip() in "ABSENT" or str(p.iat[r+pCnt+1,i]).strip() in "Unused":
                                                        wasUsed = False
                                                        if str(p.iat[r+pCnt,4]).strip() != "nan":
                                                            personelDict[str(p.iat[r+pCnt,4]).strip()] = {
                                                                "Type": str(p.iat[r+pCnt,3]).strip(),
                                                                "DDE": "UNUSED",
                                                                "LDE": "UNUSED",
                                                                "SDE": "UNUSED"
                                                            }
                                                        else:
                                                            personelDict["??"] = {
                                                                "Type": str(p.iat[r+pCnt,3]).strip(),
                                                                "DDE": "UNUSED",
                                                                "LDE": "UNUSED",
                                                                "SDE": "UNUSED"
                                                            }
                                            if wasUsed == False:
                                                pCnt = pCnt + 2
                                                #print("Unused or Absent Dosimeter! at " + str(i))
                                            else:
                                                pCnt = pCnt + 1
                                        else:
                                            if str(p.iat[r+pCnt,4]).strip() != "nan":
                                                personelDict[str(p.iat[r+pCnt,4]).strip()] = {
                                                    "Type": str(p.iat[r+pCnt,3]).strip(),
                                                    "DDE": p.iat[r,7],
                                                    "LDE": p.iat[r,8],
                                                    "SDE": p.iat[r,9]
                                                }
                                            else:
                                                personelDict["??"] = {
                                                    "Type": str(p.iat[r+pCnt,3]).strip(),
                                                    "DDE": p.iat[r,7],
                                                    "LDE": p.iat[r,8],
                                                    "SDE": p.iat[r,9]
                                                }
                                            pCnt = pCnt + 1
                                    else:
                                        if str(p.iat[r+pCnt,4]).strip() != "nan":
                                            personelDict[str(p.iat[r+pCnt,4]).strip()] = {
                                                "Type": str(p.iat[r+pCnt,3]).strip(),
                                                "DDE": p.iat[r,7],
                                                "LDE": p.iat[r,8],
                                                "SDE": p.iat[r,9]
                                            }
                                        else:
                                            personelDict["??"] = {
                                                "Type": str(p.iat[r+pCnt,3]).strip(),
                                                "DDE": p.iat[r,7],
                                                "LDE": p.iat[r,8],
                                                "SDE": p.iat[r,9]
                                            }
                                        pCnt = pCnt + 1
                                        break 
                                    if r+pCnt >= p.shape[0]:
                                        break
                                allPersonel.append(personelDict)
                        elif recording > 1:
                            #print("stop!")
                            recording=1000
            #print(allPersonel)

            msgBox = QMessageBox();
            msgBox.setText("Writing File...");
            msgBox.exec();

            jsonOutDict = {"personel": allPersonel}
            jsonOutCorrectedDict = ""
            if Window.line_edit.text() != "":
                jsonOutName = Window.line_edit.text() + ".json"
            else:
                jsonOutName = "sampleLandauer.json"
            
            with open(jsonOutName, "w") as outfile: 
                json.dump(jsonOutDict, outfile)
            with open(jsonOutName, "r") as checkfile:
                jsonOutCorrectedString = checkfile.read().replace(": NaN", ": '??'")
                jsonOutCorrectedDict = eval(jsonOutCorrectedString)
            with open(jsonOutName, "w") as outfile: 
                json.dump(jsonOutCorrectedDict, outfile)
        else:
            doNothing = True

app = QApplication([])
app.setApplicationName("Landauer Report processor")
window = Window()

window.show()   
sys.exit(app.exec())
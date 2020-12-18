# !/usr/bin/env python3
# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import ui_form
import sqlite3
import sys
import os
import pyperclip

class Pos:
    def __init__(self, start, itemNumber):
        self.itemNumber = itemNumber
        self.start = start

class DataSearcher:
    excelItems = ["item_number","item_category","item_name","SILO (BRAND)","unit_of_measure","quantity","manufacturer_1","manufacturer_2","current_prod_cost","bom_notes"]
    excelItemsIndex = {}
    levelPosMap = []
    title = None
    fp = None

    def __init__(self, fileName, signal = None):
        try:
            print('Load file', fileName)
            self.fp = open(fileName, mode='r', encoding='gb18030')

            while True:
                start = self.fp.tell()
                line = self.fp.readline().replace('\n', '')
                if line == "":
                    break
                lineData = line.split(',')
                if lineData[0] == "level":
                    self.title = Pos(start, 'item_number')
                    for item in self.excelItems:
                        self.excelItemsIndex[item] = lineData.index(item)
                    print(self.excelItemsIndex)
                    print(line)
                    break

            if self.title:
                self.levelPosMap = []
                while True:
                    start = self.fp.tell()
                    line = self.fp.readline().replace('\n', '')
                    lineData = line.split(',')
                    if lineData[0] == "0":
                        self.levelPosMap.append(Pos(start, lineData[self.excelItemsIndex['item_number']]))
                    if line == "":
                        break

                print('Found', len(self.levelPosMap), 'items')
                if signal:
                    signal.emit(len(self.levelPosMap))

        except Exception as err:
            print(err)
            return None

    def GetLevelItems(self):
        if self.title:
            items = []
            for levelPos in self.levelPosMap:
                self.fp.seek(levelPos.start)
                itemLineData = self.fp.readline().replace('\n', '').split(',')
                line = []
                for k,v in self.excelItemsIndex.items():
                    line.append(itemLineData[v])
                items.append(line)
            return items
        return []
        pass

    def SearchItem(self, itemNumber):
        if self.title:
            print('Searing', itemNumber, '...')
            result = []
            for levelPos in self.levelPosMap:
                if levelPos.itemNumber == itemNumber:
                    print("Found item(%s):"%(levelPos.itemNumber))

                    #title
                    # line = []
                    # for k,v in self.excelItemsIndex.items():
                    #     line.append(k)
                    # result.append(line)

                    self.fp.seek(levelPos.start)

                    #itemNumber
                    itemLineData = self.fp.readline().replace('\n', '').split(',')
                    line = []
                    for k,v in self.excelItemsIndex.items():
                        line.append(itemLineData[v])
                    result.append(line)

                    while True:
                        line = self.fp.readline().replace('\n', '')
                        if line == "":
                            break
                        lineData = line.split(',')
                        if lineData[0] == "0":
                            break

                        #sub item
                        line = []
                        for k,v in self.excelItemsIndex.items():
                            line.append(lineData[v])
                        result.append(line)
                    return result
        return []

    def __del__(self):
        if self.fp:
            self.fp.close()


class MainWindow(QMainWindow):
    foundItemSignal = pyqtSignal(int)

    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = ui_form.Ui_Form()
        self.ui.setupUi(self)
        self.ui.tableWidget.setColumnCount(0)
        self.ui.tableWidget.setRowCount(0)
        self.ui.tableWidget.clear()
        self.ui.lineEdit.clear()
        self.ui.FoundLable.clear()
        self.ui.SaveResultButton.clicked.connect(self.SaveResult)
        self.ui.OpenFileButton.clicked.connect(self.OpenFileButton)
        self.ui.lineEdit.returnPressed.connect(self.SearchItem)
        self.ui.CopyButton.clicked.connect(self.CopyResultToClip)
        self.dataSearcher = None
        self.searchResult = []
        self._translate = QCoreApplication.translate
        self.foundItemSignal.connect(self.UpdateFoundLable)
        QMessageBox.about(self,"","Welcome to Yuke Sports BOM Search System")

    def OpenFileButton(self):
        fileName, filetype = QFileDialog.getOpenFileName(self, "选取文件", '', "Text Files (*.csv)")
        if fileName == "":
            return
        self.dataSearcher = DataSearcher(fileName, self.foundItemSignal)
        self.ui.tableWidget.clear()
        self.ui.tableWidget.setColumnCount(len(self.dataSearcher.excelItems))
        for i, val in enumerate(self.dataSearcher.excelItems):
            self.ui.tableWidget.setHorizontalHeaderItem(i, QTableWidgetItem())
            item = self.ui.tableWidget.horizontalHeaderItem(i)
            item.setText(self._translate("Form", self.dataSearcher.excelItems[i]))
        self.ShowItems()
        pass

    def ShowItems(self):
        if self.dataSearcher:
            self.searchResult = self.dataSearcher.GetLevelItems()
            self.ShowResult()
        pass

    def ShowResult(self):
        if len(self.searchResult) > 0:
            self.ui.tableWidget.clearContents()
            self.ui.tableWidget.setRowCount(len(self.searchResult))
            for col, colVal in enumerate(self.searchResult):
                for row, rowVal in enumerate(colVal):
                    self.ui.tableWidget.setItem(col, row, QTableWidgetItem())
                    item = self.ui.tableWidget.item(col, row)
                    item.setText(self._translate("Form", rowVal))
        pass

    def CopyResultToClip(self):
        if len(self.searchResult) > 0:
            result = ''
            for item in self.dataSearcher.excelItems:
                result = result + item + '\t'
            result += '\n'

            for col, colVal in enumerate(self.searchResult):
                for row, rowVal in enumerate(colVal):
                    result = result + rowVal + '\t'
                result += '\n'
            pyperclip.copy(result)
            QMessageBox.about(self,"","复制成功")
            print("Copy result to clip")
        pass

    def SaveResult(self):
        fileName, filetype = QFileDialog.getSaveFileName(self, "文件保存", '', "Text Files (*.csv)")  

        if fileName == "":
            return

        print('Save to file %s'%(fileName))
        fp = open(fileName, mode='w+')

        for item in self.dataSearcher.excelItems:
            fp.write(item + ',')
        fp.write('\n')

        for col, colVal in enumerate(self.searchResult):
            for row, rowVal in enumerate(colVal):
                fp.write(rowVal + ',')
            fp.write('\n')
        fp.close()
        pass

    def SearchItem(self):
        searchItemNumber = self.ui.lineEdit.text()
        if searchItemNumber != "":
            self.searchResult = self.dataSearcher.SearchItem(searchItemNumber)
            self.ShowResult()
        else:
            self.ShowItems()
        pass

    def UpdateFoundLable(self, size):
        self.ui.FoundLable.setText(self._translate("Form", "共" + str(size) + "个型号"))
        pass

if __name__ == "__main__":
    QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    app = QApplication([])
    application = MainWindow()
    application.show()
    sys.exit(app.exec())

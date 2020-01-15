import sys
import os
if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']

from PyQt5 import QtCore, QtWidgets, QtGui
from Ui_Main import Ui_MainWindow
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import pyqtgraph as pg

import numpy as np
import xlsxwriter


class LoadS2P(Ui_MainWindow, QMainWindow):
    def __init__(self, parent=None):
        super(LoadS2P, self).__init__(parent)
        self.setupUi(self)
        self.registerEvent()
        self.tabelNum = 0
        self.DataNull = False

        pg.setConfigOptions(antialias=True)
        self.plotgraph = pg.PlotWidget(title="S-Parameter")
        self.Graphics_Layout.addWidget(self.plotgraph)

    def initGraph(self):
        # pg.setConfigOption('background','w')
        # pg.setConfigOption('foreground','b')
        if self.DataNull:
            self.plotgraph.clear()
            self.plotgraph.plot(self._freq,self._mag_s11,pen=(242,242,0),name="S11")
            self.plotgraph.plot(self._freq,self._mag_s22,pen=(236,5,236))           
            self.plotgraph.plot(self._freq,self._mag_s21,pen=(3,245,245))
        
    def registerEvent(self):
        self.ImportButton.clicked.connect(self.import_click)
        self.ExportButton.clicked.connect(self.export_click)
        self.SaveButton.clicked.connect(self.save_click)
        self.OpenButton.clicked.connect(self.open_click)
        self.AddButton.clicked.connect(self.add_click)
        self.DeleteButton.clicked.connect(self.delete_click)

        self.MaxIL_Table.clicked.connect(self.maxiltabel_click)
        self.AvgIL_Table.clicked.connect(self.avgtabel_click)
        self.Att_Table.clicked.connect(self.atttabel_click)
        self.RL_Table.clicked.connect(self.rltable_click)
        self.Ripple_Table.clicked.connect(self.ripple_click)

        self.MaxIL_Table.horizontalHeader().sectionClicked.connect(self.maxiltabel_click)
        self.AvgIL_Table.horizontalHeader().sectionClicked.connect(self.avgtabel_click)
        self.Att_Table.horizontalHeader().sectionClicked.connect(self.atttabel_click)
        self.RL_Table.horizontalHeader().sectionClicked.connect(self.rltable_click)
        self.Ripple_Table.horizontalHeader().sectionClicked.connect(self.ripple_click)

        self.CalculateButton.clicked.connect(self.calculate_click)
        self.ApplyButton.clicked.connect(self.apply_click)
        self.HelpButton.clicked.connect(self.help_click)

    def setData(self, file_name):
        self.DataNull = True
        self.load_file_name = file_name
        self.file_data = np.loadtxt(
            file_name, dtype=float, comments=['!', '#'])
        self._freq = self.file_data[..., 0]
        self._mag_s11 = self.file_data[..., 1]
        self._ang_s11 = self.file_data[..., 2]
        self._mag_s21 = self.file_data[..., 3]
        self._ang_s21 = self.file_data[..., 4]
        self._mag_s12 = self.file_data[..., 5]
        self._ang_s12 = self.file_data[..., 6]
        self._mag_s22 = self.file_data[..., 7]
        self._ang_s22 = self.file_data[..., 8]

    def get_freq(self):
        return self._freq

    def get_s11(self):
        return self._mag_s11, self._ang_s11

    def get_s12(self):
        return self._mag_s12, self._ang_s12

    def get_s21(self):
        return self._mag_s21, self._ang_s21

    def get_s22(self):
        return self._mag_s22, self._ang_s22

    def search_withpoint(self, point):

        if self.DataNull:
            point = point * 1000000.0
            f0 = self._freq[0]
            kbw = self._freq[1] - self._freq[0]
            _value = []
            if ((point-f0) % kbw == 0):
                freq_low = self._freq.tolist().index(point)
                return self._mag_s11[freq_low], self._mag_s21[freq_low], self._mag_s12[freq_low], self._mag_s22[freq_low]
            else:
                freq_low = self._freq.tolist().index(f0 + ((point-f0)//kbw)*kbw)
                freq_high = self._freq.tolist().index(f0 + ((point-f0)//kbw + 1.0)*kbw)
                value_low = [self._mag_s11[freq_low], self._mag_s21[freq_low],
                             self._mag_s12[freq_low], self._mag_s22[freq_low]]
                value_high = [self._mag_s11[freq_high], self._mag_s21[freq_high],
                              self._mag_s12[freq_high], self._mag_s22[freq_high]]
            for i in range(len(value_low)):
                _value.append(
                    value_low[i] + (value_high[i] - value_low[i]) / (self._freq[freq_high] - self._freq[freq_low]) * (point - self._freq[freq_low]))
            return _value

    def search_withrange(self, point1, point2, Max_AVG='Max'):
        if self.DataNull:
            point1 = point1 * 1000000
            point2 = point2 * 1000000
            data = []
            f0 = self._freq[0]
            kbw = self._freq[1] - self._freq[0]
            if point1 in self._freq:
                point3 = point1
            else:
                point3 = f0 + ((point1-f0)//kbw + 1.0)*kbw
            if point2 in self._freq:
                point4 = point2
            else:
                point4 = f0 + ((point2-f0)//kbw)*kbw
            if point3 not in self._freq or point4 not in self._freq:
                data = "over"
            else:
                ILdata = self._mag_s21[self._freq.tolist().index(
                    point3):self._freq.tolist().index(point4)+1]
                ILdata = np.append(
                    ILdata, self.search_withpoint(point1 / 1000000)[1])
                ILdata = np.append(
                    ILdata, self.search_withpoint(point2 / 1000000)[1])
                if Max_AVG == 'Max':
                    data = max(ILdata)
                if Max_AVG == 'AVG':
                    data = np.mean(ILdata)
                if Max_AVG == 'Min':
                    data = min(ILdata)
                if Max_AVG == 'RL':
                    RLdata1 = self._mag_s11[self._freq.tolist().index(
                        point3):self._freq.tolist().index(point4)+1]
                    RLdata2 = self._mag_s22[self._freq.tolist().index(
                        point3):self._freq.tolist().index(point4)+1]
                    RLdata1 = np.append(
                        RLdata1, self.search_withpoint(point1 / 1000000)[0])
                    RLdata1 = np.append(
                        RLdata1, self.search_withpoint(point2 / 1000000)[0])
                    RLdata2 = np.append(
                        RLdata2, self.search_withpoint(point1 / 1000000)[3])
                    RLdata2 = np.append(
                        RLdata2, self.search_withpoint(point2 / 1000000)[3])
                    data = [max(RLdata1), max(RLdata2)]
                if Max_AVG == 'Ripple':
                    data = min(ILdata) - max(ILdata)
            return data

    def import_click(self):
        fileName, fileType = QFileDialog.getOpenFileName(self, "Choose S2P", os.getcwd(),
                                                         "s2p Files(*.s2p)")
        if fileName == "":
            return
        else:
            self.FileName_label.setText('Open File：\n'+fileName)
            self.setData(fileName)

    def save_click(self):
        fileName, fileType = QFileDialog.getSaveFileName(self, "Save File", os.getcwd(),
                                                         "TXT Files(*.txt)")
        if fileName == "":
            return
        else:
            return

    def open_click(self):
        fileName, fileType = QFileDialog.getOpenFileName(self, "Open File", os.getcwd(),
                                                         "TXT Files(*.txt)")
        if fileName == "":
            return
        else:
            return

    def export_click(self):
        fileName, fileType = QFileDialog.getSaveFileName(self, "Export Excel", os.getcwd(),
                                                         "Excel Files(*.xlsx)")
        if fileName == "":
            return
        else:
            if self.DataNull:
                work_book = xlsxwriter.Workbook(fileName)
                book_sheet = work_book.add_worksheet('#1')
                bold = work_book.add_format(
                    {'bold': True, 'font': 'Arial', 'font_size': 10})
                format1 = work_book.add_format()
                format1.set_num_format('0.00')

                book_sheet.write_row('A1', ['Return Loss'], bold)
                book_sheet.write_row(
                    'A2', ['Start/MHz', 'Stop/MHz', 'Value/dB'], bold)
                book_sheet.write_column(
                    'A3', self.getTableData(self.RL_Table)[0])
                book_sheet.write_column(
                    'B3', self.getTableData(self.RL_Table)[1])
                book_sheet.write_column(
                    'C3', self.getTableData(self.RL_Table)[2], format1)

                book_sheet.write_row('A5', ['Max Insertion Loss'], bold)
                book_sheet.write_row(
                    'A6', ['Start/MHz', 'Stop/MHz', 'Value/dB'], bold)
                book_sheet.write_column(
                    'A7', self.getTableData(self.MaxIL_Table)[0])
                book_sheet.write_column(
                    'B7', self.getTableData(self.MaxIL_Table)[1])
                book_sheet.write_column(
                    'C7', self.getTableData(self.MaxIL_Table)[2], format1)

                rowNumber = self.RL_Table.rowCount() + 5 + self.MaxIL_Table.rowCount()
                book_sheet.write_row('A'+str(rowNumber),
                                     ['Avg Insertion Loss'], bold)
                book_sheet.write_row(
                    'A'+str(rowNumber+1), ['Start/MHz', 'Stop/MHz', 'Value/dB'], bold)
                rowNumber = rowNumber + 2
                book_sheet.write_column(
                    'A'+str(rowNumber), self.getTableData(self.AvgIL_Table)[0])
                book_sheet.write_column(
                    'B'+str(rowNumber), self.getTableData(self.AvgIL_Table)[1])
                book_sheet.write_column(
                    'C'+str(rowNumber), self.getTableData(self.AvgIL_Table)[2], format1)

                rowNumber = rowNumber + self.AvgIL_Table.rowCount()
                book_sheet.write_row('A'+str(rowNumber), ['IL Ripple'], bold)
                book_sheet.write_row(
                    'A'+str(rowNumber+1), ['Start/MHz', 'Stop/MHz', 'Value/dB'], bold)
                rowNumber = rowNumber + 2
                book_sheet.write_column(
                    'A'+str(rowNumber), self.getTableData(self.Ripple_Table)[0])
                book_sheet.write_column(
                    'B'+str(rowNumber), self.getTableData(self.Ripple_Table)[1])
                book_sheet.write_column(
                    'C'+str(rowNumber), self.getTableData(self.Ripple_Table)[2], format1)

                rowNumber = rowNumber + self.Ripple_Table.rowCount()
                book_sheet.write_row('A'+str(rowNumber), ['Attenuation'], bold)
                book_sheet.write_row(
                    'A'+str(rowNumber+1), ['Start/MHz', 'Stop/MHz', 'Value/dB'], bold)
                rowNumber = rowNumber + 2
                book_sheet.write_column(
                    'A'+str(rowNumber), self.getTableData(self.Att_Table)[0])
                book_sheet.write_column(
                    'B'+str(rowNumber), self.getTableData(self.Att_Table)[1])
                book_sheet.write_column(
                    'C'+str(rowNumber), self.getTableData(self.Att_Table)[2], format1)

                work_book.close()
            return

    def help_click(self):
        self.helpbox = QMessageBox(
            QMessageBox.Question, 'Help', 'All rights reserved @ Luoxiaotao')
        ruturnBtn = self.helpbox.addButton('返回', QMessageBox.YesRole)
        self.helpbox.show()

    def apply_click(self):
        if self.DataNull:
            self.addTable_Data(self.Att_Table)
            self.addTable_Data(self.MaxIL_Table)
            self.addTable_Data(self.AvgIL_Table)
            self.addTable_Data(self.RL_Table)
            self.addTable_Data(self.Ripple_Table)
            self.initGraph()

    def addTable_Data(self, Table):
        attRow = Table.rowCount()
        attColum = Table.columnCount()
        attData = []
        IfItemNull = True

        for i in range(attRow):
            if (Table.item(i, 0) == None and Table.item(i, 1) != None) or (Table.item(i, 1) == None and Table.item(i, 0) != None):
                IfItemNull = False

        if IfItemNull:
            for i in range(attRow):
                attData.append([])
                for j in range(attColum-1):
                    if Table.item(i, j) == None:
                        item = QTableWidgetItem('0')
                        Table.setItem(i, j, item)
                    attData[i].append(int(Table.item(i, j).text()))
                if Table == self.AvgIL_Table:
                    item = QTableWidgetItem(
                        str(self.search_withrange(attData[i][0], attData[i][1], 'AVG'))[:7])
                    item.setFlags(Qt.ItemIsEditable)
                    Table.setItem(i, 2, item)
                if Table == self.Att_Table:
                    item = QTableWidgetItem(
                        str(self.search_withrange(attData[i][0], attData[i][1]))[:7])
                    item.setFlags(Qt.ItemIsEditable)
                    Table.setItem(i, 2, item)
                if Table == self.MaxIL_Table:
                    item = QTableWidgetItem(
                        str(self.search_withrange(attData[i][0], attData[i][1], 'Min'))[:7])
                    item.setFlags(Qt.ItemIsEditable)
                    Table.setItem(i, 2, item)
                if Table == self.RL_Table:
                    item = QTableWidgetItem(
                        str(self.search_withrange(attData[i][0], attData[i][1], 'RL')[i])[:7])
                    item.setFlags(Qt.ItemIsEditable)
                    Table.setItem(i, 2, item)
                if Table == self.Ripple_Table:
                    item = QTableWidgetItem(
                        str(self.search_withrange(attData[i][0], attData[i][1], 'Ripple'))[:7])
                    item.setFlags(Qt.ItemIsEditable)
                    Table.setItem(i, 2, item)

    def getTableData(self, Table):
        startFreq = []
        stopFreq = []
        calvalue = []
        for i in range(Table.rowCount()):
            startFreq.append(Table.item(i, 0).text())
            stopFreq.append(Table.item(i, 1).text())
            calvalue.append(Table.item(i, 2).text())
        tabledata = [startFreq, stopFreq, calvalue]
        return tabledata

    def add_click(self):
        if self.tabelNum == 1:
            self.insertTableRow(self.MaxIL_Table)
        if self.tabelNum == 2:
            self.insertTableRow(self.AvgIL_Table)
        if self.tabelNum == 3:
            self.insertTableRow(self.Att_Table)
        if self.tabelNum == 4:
            self.insertTableRow(self.Ripple_Table)

    def delete_click(self):
        if self.tabelNum == 1:
            self.deleteTableRow(self.MaxIL_Table)
        if self.tabelNum == 2:
            self.deleteTableRow(self.AvgIL_Table)
        if self.tabelNum == 3:
            self.deleteTableRow(self.Att_Table)
        if self.tabelNum == 4:
            self.deleteTableRow(self.Ripple_Table)       

    def maxiltabel_click(self):
        self.tabelNum = 1

    def avgtabel_click(self):
        self.tabelNum = 2

    def atttabel_click(self):
        self.tabelNum = 3

    def rltable_click(self):
        self.tabelNum = 0

    def ripple_click(self):
        self.tabelNum = 4

    def calculate_click(self):
        return

    def insertTableRow(self, mtable, value=''):
        rowcount = mtable.rowCount()
        mtable.insertRow(rowcount)
        item = QTableWidgetItem(value)
        item.setFlags(Qt.ItemIsEditable)
        mtable.setItem(rowcount, 2, item)

    def deleteTableRow(self, mtable):
        index = mtable.currentRow()
        if index != -1:
            mtable.removeRow(index)


if __name__ == "__main__":

    app = QApplication(sys.argv)
    load = LoadS2P()
    load.show()

    font = QFont('等线', 9)
    font.setBold(True)

    load.MaxIL_Table.setRowCount(0)
    load.MaxIL_Table.setColumnCount(3)
    load.MaxIL_Table.horizontalHeader().setFont(font)
    load.MaxIL_Table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    load.MaxIL_Table.setHorizontalHeaderLabels(
        ['Start/MHz', 'Stop/MHz', 'Value/dB'])
    load.insertTableRow(load.MaxIL_Table)

    load.RL_Table.setRowCount(2)
    load.RL_Table.setColumnCount(3)
    load.RL_Table.horizontalHeader().setFont(font)
    load.RL_Table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    load.RL_Table.setHorizontalHeaderLabels(
        ['Start/MHz', 'Stop/MHz', 'Value/dB'])

    load.AvgIL_Table.setRowCount(0)
    load.AvgIL_Table.setColumnCount(3)
    load.AvgIL_Table.horizontalHeader().setFont(font)
    load.AvgIL_Table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    load.AvgIL_Table.setHorizontalHeaderLabels(
        ['Start/MHz', 'Stop/MHz', 'Value/dB'])
    load.insertTableRow(load.AvgIL_Table)

    load.Ripple_Table.setRowCount(0)
    load.Ripple_Table.setColumnCount(3)
    load.Ripple_Table.horizontalHeader().setFont(font)
    load.Ripple_Table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    load.Ripple_Table.setHorizontalHeaderLabels(
        ['Start/MHz', 'Stop/MHz', 'Value/dB'])
    load.insertTableRow(load.Ripple_Table)

    load.Att_Table.setRowCount(0)
    load.Att_Table.setColumnCount(3)
    load.Att_Table.horizontalHeader().setFont(font)
    load.Att_Table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    load.Att_Table.setHorizontalHeaderLabels(
        ['Start/MHz', 'Stop/MHz', 'Value/dB'])
    load.insertTableRow(load.Att_Table)

    sys.exit(app.exec_())

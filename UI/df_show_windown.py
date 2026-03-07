import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QTableView, QMessageBox
)
from PyQt5.QtCore import QAbstractTableModel, Qt, QVariant
import pandas as pd
import os
import re

# Custom model để hiển thị DataFrame trong QTableView
class PandasModel(QAbstractTableModel):
    def __init__(self, df=pd.DataFrame(), parent=None):
        super(PandasModel, self).__init__(parent)
        self._df = df

    def rowCount(self, parent=None):
        return self._df.shape[0]

    def columnCount(self, parent=None):
        return self._df.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid() and role == Qt.DisplayRole:
            value = self._df.iloc[index.row(), index.column()]
            return str(value)
        return QVariant()

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return self._df.columns[section]
            else:
                return str(self._df.index[section])
        return QVariant()

class df_show(object):
    def setupUi(self, windown,data):
        windown.setObjectName("Data file")
        windown.resize(800, 600)
        self.central_widget = QWidget(windown)
        windown.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        self.table_view = QTableView(self.central_widget)
        self.layout.addWidget(self.table_view)
        self.show_data(data)
        windown.show()

    def show_data(self,data):
        model = PandasModel(data)
        self.table_view.setModel(model)

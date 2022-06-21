from PyQt5 import QtCore, QtGui, QtWidgets
import cv2
import numpy as np
import threading
import time
import matplotlib.pyplot as plt
from numpy import trapz
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import serial
import serial.tools.list_ports
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QWidget, QPushButton, QLineEdit, QInputDialog, QApplication, QLCDNumber
from PyQt5.QtGui import QPixmap, QColor, QIcon
from PyQt5.QtCore import pyqtSignal, pyqtSlot, Qt, QObject, QThread
from pyqtgraph import PlotWidget, plot
import pyqtgraph as pg
import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
from matplotlib.figure import Figure

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from pyvcam import pvc
from pyvcam.camera import Camera

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(850, 680)
        MainWindow.setMinimumSize(QtCore.QSize(800, 600))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_5 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_5.setObjectName("gridLayout_5")
        self.frame_3 = QtWidgets.QFrame(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame_3.sizePolicy().hasHeightForWidth())
        self.frame_3.setSizePolicy(sizePolicy)
        self.frame_3.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_3.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_3.setObjectName("frame_3")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.frame_3)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.camera_frame = QtWidgets.QLabel(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.camera_frame.sizePolicy().hasHeightForWidth())
        self.camera_frame.setSizePolicy(sizePolicy)
        self.camera_frame.setMinimumSize(QtCore.QSize(341, 171))
        self.camera_frame.setStyleSheet("background-color: rgb(0, 0, 0);")
        self.camera_frame.setText("")
        self.camera_frame.setAlignment(QtCore.Qt.AlignCenter)
        self.camera_frame.setObjectName("camera_frame")
        self.gridLayout_2.addWidget(self.camera_frame, 0, 0, 1, 1)
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_2.addItem(spacerItem, 3, 0, 1, 1)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_4 = QtWidgets.QLabel(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_4.sizePolicy().hasHeightForWidth())
        self.label_4.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout.addWidget(self.label_4)
        self.amperage_value = QtWidgets.QLCDNumber(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.amperage_value.sizePolicy().hasHeightForWidth())
        self.amperage_value.setSizePolicy(sizePolicy)
        self.amperage_value.setStyleSheet("color: rgb(255, 0, 0);\n"
"")
        self.amperage_value.setObjectName("amperage_value")
        self.horizontalLayout.addWidget(self.amperage_value)
        self.camera_rec = QtWidgets.QPushButton(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.camera_rec.sizePolicy().hasHeightForWidth())
        self.camera_rec.setSizePolicy(sizePolicy)
        self.camera_rec.setStyleSheet("background-color: rgb(222, 222, 222);\n"
"font: 75 10pt \"Times New Roman\";")
        self.camera_rec.setObjectName("camera_rec")
        self.horizontalLayout.addWidget(self.camera_rec)
        self.camera_setting = QtWidgets.QPushButton(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.camera_setting.sizePolicy().hasHeightForWidth())
        self.camera_setting.setSizePolicy(sizePolicy)
        self.camera_setting.setStyleSheet("background-color: rgb(222, 222, 222);\n"
"font: 75 10pt \"Times New Roman\";")
        self.camera_setting.setObjectName("camera_setting")
        self.horizontalLayout.addWidget(self.camera_setting)
        self.gridLayout_2.addLayout(self.horizontalLayout, 1, 0, 1, 1)
        self.gridLayout_4 = QtWidgets.QGridLayout()
        self.gridLayout_4.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.gridLayout_4.setHorizontalSpacing(0)
        self.gridLayout_4.setVerticalSpacing(4)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.textEdit_69 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_69.sizePolicy().hasHeightForWidth())
        self.textEdit_69.setSizePolicy(sizePolicy)
        self.textEdit_69.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_69.setObjectName("textEdit_69")
        self.gridLayout_4.addWidget(self.textEdit_69, 2, 2, 1, 1)
        self.src1_line0 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.src1_line0.sizePolicy().hasHeightForWidth())
        self.src1_line0.setSizePolicy(sizePolicy)
        self.src1_line0.setMaximumSize(QtCore.QSize(41, 31))
        self.src1_line0.setObjectName("src1_line0")
        self.gridLayout_4.addWidget(self.src1_line0, 1, 1, 1, 1)
        self.label_13 = QtWidgets.QLabel(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_13.sizePolicy().hasHeightForWidth())
        self.label_13.setSizePolicy(sizePolicy)
        self.label_13.setAlignment(QtCore.Qt.AlignCenter)
        self.label_13.setObjectName("label_13")
        self.gridLayout_4.addWidget(self.label_13, 0, 1, 1, 1)
        self.label_15 = QtWidgets.QLabel(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_15.sizePolicy().hasHeightForWidth())
        self.label_15.setSizePolicy(sizePolicy)
        self.label_15.setAlignment(QtCore.Qt.AlignCenter)
        self.label_15.setObjectName("label_15")
        self.gridLayout_4.addWidget(self.label_15, 0, 3, 1, 1)
        self.label_16 = QtWidgets.QLabel(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_16.sizePolicy().hasHeightForWidth())
        self.label_16.setSizePolicy(sizePolicy)
        self.label_16.setAlignment(QtCore.Qt.AlignCenter)
        self.label_16.setObjectName("label_16")
        self.gridLayout_4.addWidget(self.label_16, 0, 4, 1, 1)
        self.label_12 = QtWidgets.QLabel(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_12.sizePolicy().hasHeightForWidth())
        self.label_12.setSizePolicy(sizePolicy)
        self.label_12.setAlignment(QtCore.Qt.AlignCenter)
        self.label_12.setObjectName("label_12")
        self.gridLayout_4.addWidget(self.label_12, 0, 0, 1, 1)
        self.label_14 = QtWidgets.QLabel(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_14.sizePolicy().hasHeightForWidth())
        self.label_14.setSizePolicy(sizePolicy)
        self.label_14.setAlignment(QtCore.Qt.AlignCenter)
        self.label_14.setObjectName("label_14")
        self.gridLayout_4.addWidget(self.label_14, 0, 2, 1, 1)
        self.textEdit_64 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_64.sizePolicy().hasHeightForWidth())
        self.textEdit_64.setSizePolicy(sizePolicy)
        self.textEdit_64.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_64.setObjectName("textEdit_64")
        self.gridLayout_4.addWidget(self.textEdit_64, 1, 2, 1, 1)
        self.textEdit_63 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_63.sizePolicy().hasHeightForWidth())
        self.textEdit_63.setSizePolicy(sizePolicy)
        self.textEdit_63.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_63.setObjectName("textEdit_63")
        self.gridLayout_4.addWidget(self.textEdit_63, 1, 3, 1, 1)
        self.time_line1 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.time_line1.sizePolicy().hasHeightForWidth())
        self.time_line1.setSizePolicy(sizePolicy)
        self.time_line1.setMaximumSize(QtCore.QSize(41, 31))
        self.time_line1.setObjectName("time_line1")
        self.gridLayout_4.addWidget(self.time_line1, 2, 0, 1, 1)
        self.textEdit_68 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_68.sizePolicy().hasHeightForWidth())
        self.textEdit_68.setSizePolicy(sizePolicy)
        self.textEdit_68.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_68.setObjectName("textEdit_68")
        self.gridLayout_4.addWidget(self.textEdit_68, 2, 1, 1, 1)
        self.time_line0 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.time_line0.sizePolicy().hasHeightForWidth())
        self.time_line0.setSizePolicy(sizePolicy)
        self.time_line0.setMaximumSize(QtCore.QSize(41, 31))
        self.time_line0.setObjectName("time_line0")
        self.gridLayout_4.addWidget(self.time_line0, 1, 0, 1, 1)
        self.textEdit_52 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_52.sizePolicy().hasHeightForWidth())
        self.textEdit_52.setSizePolicy(sizePolicy)
        self.textEdit_52.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_52.setObjectName("textEdit_52")
        self.gridLayout_4.addWidget(self.textEdit_52, 3, 4, 1, 1)
        self.time_line3 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.time_line3.sizePolicy().hasHeightForWidth())
        self.time_line3.setSizePolicy(sizePolicy)
        self.time_line3.setMaximumSize(QtCore.QSize(41, 31))
        self.time_line3.setObjectName("time_line3")
        self.gridLayout_4.addWidget(self.time_line3, 4, 0, 1, 1)
        self.textEdit_70 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_70.sizePolicy().hasHeightForWidth())
        self.textEdit_70.setSizePolicy(sizePolicy)
        self.textEdit_70.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_70.setObjectName("textEdit_70")
        self.gridLayout_4.addWidget(self.textEdit_70, 4, 2, 1, 1)
        self.textEdit_46 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_46.sizePolicy().hasHeightForWidth())
        self.textEdit_46.setSizePolicy(sizePolicy)
        self.textEdit_46.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_46.setObjectName("textEdit_46")
        self.gridLayout_4.addWidget(self.textEdit_46, 6, 1, 1, 1)
        self.time_line2 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.time_line2.sizePolicy().hasHeightForWidth())
        self.time_line2.setSizePolicy(sizePolicy)
        self.time_line2.setMaximumSize(QtCore.QSize(41, 31))
        self.time_line2.setObjectName("time_line2")
        self.gridLayout_4.addWidget(self.time_line2, 3, 0, 1, 1)
        self.textEdit_42 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_42.sizePolicy().hasHeightForWidth())
        self.textEdit_42.setSizePolicy(sizePolicy)
        self.textEdit_42.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_42.setObjectName("textEdit_42")
        self.gridLayout_4.addWidget(self.textEdit_42, 5, 2, 1, 1)
        self.textEdit_43 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_43.sizePolicy().hasHeightForWidth())
        self.textEdit_43.setSizePolicy(sizePolicy)
        self.textEdit_43.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_43.setObjectName("textEdit_43")
        self.gridLayout_4.addWidget(self.textEdit_43, 5, 1, 1, 1)
        self.textEdit_57 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_57.sizePolicy().hasHeightForWidth())
        self.textEdit_57.setSizePolicy(sizePolicy)
        self.textEdit_57.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_57.setObjectName("textEdit_57")
        self.gridLayout_4.addWidget(self.textEdit_57, 2, 3, 1, 1)
        self.textEdit_67 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_67.sizePolicy().hasHeightForWidth())
        self.textEdit_67.setSizePolicy(sizePolicy)
        self.textEdit_67.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_67.setObjectName("textEdit_67")
        self.gridLayout_4.addWidget(self.textEdit_67, 1, 4, 1, 1)
        self.textEdit_60 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_60.sizePolicy().hasHeightForWidth())
        self.textEdit_60.setSizePolicy(sizePolicy)
        self.textEdit_60.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_60.setObjectName("textEdit_60")
        self.gridLayout_4.addWidget(self.textEdit_60, 3, 1, 1, 1)
        self.textEdit_61 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_61.sizePolicy().hasHeightForWidth())
        self.textEdit_61.setSizePolicy(sizePolicy)
        self.textEdit_61.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_61.setObjectName("textEdit_61")
        self.gridLayout_4.addWidget(self.textEdit_61, 3, 2, 1, 1)
        self.textEdit_50 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_50.sizePolicy().hasHeightForWidth())
        self.textEdit_50.setSizePolicy(sizePolicy)
        self.textEdit_50.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_50.setObjectName("textEdit_50")
        self.gridLayout_4.addWidget(self.textEdit_50, 3, 3, 1, 1)
        self.textEdit_54 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_54.sizePolicy().hasHeightForWidth())
        self.textEdit_54.setSizePolicy(sizePolicy)
        self.textEdit_54.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_54.setObjectName("textEdit_54")
        self.gridLayout_4.addWidget(self.textEdit_54, 4, 3, 1, 1)
        self.textEdit_53 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_53.sizePolicy().hasHeightForWidth())
        self.textEdit_53.setSizePolicy(sizePolicy)
        self.textEdit_53.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_53.setObjectName("textEdit_53")
        self.gridLayout_4.addWidget(self.textEdit_53, 4, 1, 1, 1)
        self.textEdit_59 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_59.sizePolicy().hasHeightForWidth())
        self.textEdit_59.setSizePolicy(sizePolicy)
        self.textEdit_59.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_59.setObjectName("textEdit_59")
        self.gridLayout_4.addWidget(self.textEdit_59, 2, 4, 1, 1)
        self.textEdit_56 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_56.sizePolicy().hasHeightForWidth())
        self.textEdit_56.setSizePolicy(sizePolicy)
        self.textEdit_56.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_56.setObjectName("textEdit_56")
        self.gridLayout_4.addWidget(self.textEdit_56, 4, 4, 1, 1)
        self.textEdit_37 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_37.sizePolicy().hasHeightForWidth())
        self.textEdit_37.setSizePolicy(sizePolicy)
        self.textEdit_37.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_37.setObjectName("textEdit_37")
        self.gridLayout_4.addWidget(self.textEdit_37, 7, 4, 1, 1)
        self.time_line4 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.time_line4.sizePolicy().hasHeightForWidth())
        self.time_line4.setSizePolicy(sizePolicy)
        self.time_line4.setMaximumSize(QtCore.QSize(41, 31))
        self.time_line4.setObjectName("time_line4")
        self.gridLayout_4.addWidget(self.time_line4, 5, 0, 1, 1)
        self.textEdit_48 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_48.sizePolicy().hasHeightForWidth())
        self.textEdit_48.setSizePolicy(sizePolicy)
        self.textEdit_48.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_48.setObjectName("textEdit_48")
        self.gridLayout_4.addWidget(self.textEdit_48, 6, 3, 1, 1)
        self.time_line5 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.time_line5.sizePolicy().hasHeightForWidth())
        self.time_line5.setSizePolicy(sizePolicy)
        self.time_line5.setMaximumSize(QtCore.QSize(41, 31))
        self.time_line5.setObjectName("time_line5")
        self.gridLayout_4.addWidget(self.time_line5, 6, 0, 1, 1)
        self.textEdit_47 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_47.sizePolicy().hasHeightForWidth())
        self.textEdit_47.setSizePolicy(sizePolicy)
        self.textEdit_47.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_47.setObjectName("textEdit_47")
        self.gridLayout_4.addWidget(self.textEdit_47, 6, 2, 1, 1)
        self.textEdit_49 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_49.sizePolicy().hasHeightForWidth())
        self.textEdit_49.setSizePolicy(sizePolicy)
        self.textEdit_49.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_49.setObjectName("textEdit_49")
        self.gridLayout_4.addWidget(self.textEdit_49, 6, 4, 1, 1)
        self.textEdit_40 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_40.sizePolicy().hasHeightForWidth())
        self.textEdit_40.setSizePolicy(sizePolicy)
        self.textEdit_40.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_40.setObjectName("textEdit_40")
        self.gridLayout_4.addWidget(self.textEdit_40, 7, 2, 1, 1)
        self.time_line6 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.time_line6.sizePolicy().hasHeightForWidth())
        self.time_line6.setSizePolicy(sizePolicy)
        self.time_line6.setMaximumSize(QtCore.QSize(41, 31))
        self.time_line6.setObjectName("time_line6")
        self.gridLayout_4.addWidget(self.time_line6, 7, 0, 1, 1)
        self.textEdit_39 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_39.sizePolicy().hasHeightForWidth())
        self.textEdit_39.setSizePolicy(sizePolicy)
        self.textEdit_39.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_39.setObjectName("textEdit_39")
        self.gridLayout_4.addWidget(self.textEdit_39, 7, 1, 1, 1)
        self.textEdit_41 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_41.sizePolicy().hasHeightForWidth())
        self.textEdit_41.setSizePolicy(sizePolicy)
        self.textEdit_41.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_41.setObjectName("textEdit_41")
        self.gridLayout_4.addWidget(self.textEdit_41, 5, 3, 1, 1)
        self.textEdit_44 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_44.sizePolicy().hasHeightForWidth())
        self.textEdit_44.setSizePolicy(sizePolicy)
        self.textEdit_44.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_44.setObjectName("textEdit_44")
        self.gridLayout_4.addWidget(self.textEdit_44, 5, 4, 1, 1)
        self.textEdit_36 = QtWidgets.QTextEdit(self.frame_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textEdit_36.sizePolicy().hasHeightForWidth())
        self.textEdit_36.setSizePolicy(sizePolicy)
        self.textEdit_36.setMaximumSize(QtCore.QSize(41, 31))
        self.textEdit_36.setObjectName("textEdit_36")
        self.gridLayout_4.addWidget(self.textEdit_36, 7, 3, 1, 1)
        self.gridLayout_2.addLayout(self.gridLayout_4, 2, 0, 1, 1)
        self.gridLayout_5.addWidget(self.frame_3, 0, 0, 1, 1)
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame_2.sizePolicy().hasHeightForWidth())
        self.frame_2.setSizePolicy(sizePolicy)
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.gridLayout_10 = QtWidgets.QGridLayout(self.frame_2)
        self.gridLayout_10.setObjectName("gridLayout_10")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.video_frame = QtWidgets.QLabel(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.video_frame.sizePolicy().hasHeightForWidth())
        self.video_frame.setSizePolicy(sizePolicy)
        self.video_frame.setMinimumSize(QtCore.QSize(341, 171))
        self.video_frame.setStyleSheet("background-color: rgb(0, 0, 0);")
        self.video_frame.setText("")
        self.video_frame.setObjectName("video_frame")
        self.verticalLayout_2.addWidget(self.video_frame)
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem1)
        self.video_timeline = QtWidgets.QSlider(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.video_timeline.sizePolicy().hasHeightForWidth())
        self.video_timeline.setSizePolicy(sizePolicy)
        self.video_timeline.setMinimumSize(QtCore.QSize(250, 0))
        self.video_timeline.setOrientation(QtCore.Qt.Horizontal)
        self.video_timeline.setObjectName("video_timeline")
        self.horizontalLayout_4.addWidget(self.video_timeline)
        self.video_timeline_value = QtWidgets.QSpinBox(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.video_timeline_value.sizePolicy().hasHeightForWidth())
        self.video_timeline_value.setSizePolicy(sizePolicy)
        self.video_timeline_value.setObjectName("video_timeline_value")
        self.horizontalLayout_4.addWidget(self.video_timeline_value)
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem2)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.open_video = QtWidgets.QPushButton(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.open_video.sizePolicy().hasHeightForWidth())
        self.open_video.setSizePolicy(sizePolicy)
        self.open_video.setStyleSheet("background-color: rgb(222, 222, 222);\n"
"font: 75 10pt \"Times New Roman\";")
        self.open_video.setObjectName("open_video")
        self.horizontalLayout_3.addWidget(self.open_video)
        self.video_pause = QtWidgets.QPushButton(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.video_pause.sizePolicy().hasHeightForWidth())
        self.video_pause.setSizePolicy(sizePolicy)
        self.video_pause.setStyleSheet("background-color: rgb(222, 222, 222);\n"
"font: 75 10pt \"Times New Roman\";")
        self.video_pause.setObjectName("video_pause")
        self.horizontalLayout_3.addWidget(self.video_pause)
        self.video_create_graph = QtWidgets.QPushButton(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.video_create_graph.sizePolicy().hasHeightForWidth())
        self.video_create_graph.setSizePolicy(sizePolicy)
        self.video_create_graph.setStyleSheet("background-color: rgb(222, 222, 222);\n"
"font: 75 10pt \"Times New Roman\";")
        self.video_create_graph.setObjectName("video_create_graph")
        self.horizontalLayout_3.addWidget(self.video_create_graph)
        self.save_xlsx_file = QtWidgets.QPushButton(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.save_xlsx_file.sizePolicy().hasHeightForWidth())
        self.save_xlsx_file.setSizePolicy(sizePolicy)
        self.save_xlsx_file.setStyleSheet("background-color: rgb(222, 222, 222);\n"
"font: 75 10pt \"Times New Roman\";")
        self.save_xlsx_file.setObjectName("save_xlsx_file")
        self.horizontalLayout_3.addWidget(self.save_xlsx_file)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.video_path = QtWidgets.QLineEdit(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.video_path.sizePolicy().hasHeightForWidth())
        self.video_path.setSizePolicy(sizePolicy)
        self.video_path.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.video_path.setObjectName("video_path")
        self.verticalLayout.addWidget(self.video_path)
        self.gridLayout_8 = QtWidgets.QGridLayout()
        self.gridLayout_8.setObjectName("gridLayout_8")
        self.video_mask = QtWidgets.QGroupBox(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.video_mask.sizePolicy().hasHeightForWidth())
        self.video_mask.setSizePolicy(sizePolicy)
        self.video_mask.setAlignment(QtCore.Qt.AlignCenter)
        self.video_mask.setCheckable(True)
        self.video_mask.setChecked(False)
        self.video_mask.setObjectName("video_mask")
        self.gridLayout = QtWidgets.QGridLayout(self.video_mask)
        self.gridLayout.setObjectName("gridLayout")
        self.label_2 = QtWidgets.QLabel(self.video_mask)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 2, 0, 1, 1)
        self.video_slider_upper_limit = QtWidgets.QSlider(self.video_mask)
        self.video_slider_upper_limit.setMaximum(255)
        self.video_slider_upper_limit.setSingleStep(0)
        self.video_slider_upper_limit.setProperty("value", 255)
        self.video_slider_upper_limit.setOrientation(QtCore.Qt.Horizontal)
        self.video_slider_upper_limit.setObjectName("video_slider_upper_limit")
        self.gridLayout.addWidget(self.video_slider_upper_limit, 3, 0, 1, 2)
        self.video_spin_lower_limit = QtWidgets.QSpinBox(self.video_mask)
        self.video_spin_lower_limit.setMaximum(255)
        self.video_spin_lower_limit.setProperty("value", 0)
        self.video_spin_lower_limit.setObjectName("video_spin_lower_limit")
        self.gridLayout.addWidget(self.video_spin_lower_limit, 0, 1, 1, 1)
        self.label = QtWidgets.QLabel(self.video_mask)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.video_slider_lower_limit = QtWidgets.QSlider(self.video_mask)
        self.video_slider_lower_limit.setMaximum(255)
        self.video_slider_lower_limit.setOrientation(QtCore.Qt.Horizontal)
        self.video_slider_lower_limit.setObjectName("video_slider_lower_limit")
        self.gridLayout.addWidget(self.video_slider_lower_limit, 1, 0, 1, 2)
        self.video_spin_upper_limit = QtWidgets.QSpinBox(self.video_mask)
        self.video_spin_upper_limit.setMaximum(255)
        self.video_spin_upper_limit.setProperty("value", 255)
        self.video_spin_upper_limit.setObjectName("video_spin_upper_limit")
        self.gridLayout.addWidget(self.video_spin_upper_limit, 2, 1, 1, 1)
        self.gridLayout_8.addWidget(self.video_mask, 0, 0, 2, 1)
        self.video_blure = QtWidgets.QGroupBox(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.video_blure.sizePolicy().hasHeightForWidth())
        self.video_blure.setSizePolicy(sizePolicy)
        self.video_blure.setAlignment(QtCore.Qt.AlignCenter)
        self.video_blure.setCheckable(True)
        self.video_blure.setChecked(False)
        self.video_blure.setObjectName("video_blure")
        self.gridLayout_6 = QtWidgets.QGridLayout(self.video_blure)
        self.gridLayout_6.setObjectName("gridLayout_6")
        self.label_3 = QtWidgets.QLabel(self.video_blure)
        self.label_3.setObjectName("label_3")
        self.gridLayout_6.addWidget(self.label_3, 0, 0, 1, 1)
        self.label_6 = QtWidgets.QLabel(self.video_blure)
        self.label_6.setObjectName("label_6")
        self.gridLayout_6.addWidget(self.label_6, 1, 0, 1, 1)
        self.video_blure_sigma = QtWidgets.QSpinBox(self.video_blure)
        self.video_blure_sigma.setProperty("value", 1)
        self.video_blure_sigma.setObjectName("video_blure_sigma")
        self.gridLayout_6.addWidget(self.video_blure_sigma, 1, 1, 1, 1)
        self.video_blure_k = QtWidgets.QSpinBox(self.video_blure)
        self.video_blure_k.setSingleStep(2)
        self.video_blure_k.setProperty("value", 3)
        self.video_blure_k.setObjectName("video_blure_k")
        self.gridLayout_6.addWidget(self.video_blure_k, 0, 1, 1, 1)
        self.gridLayout_8.addWidget(self.video_blure, 0, 1, 1, 1)
        self.video_get_channel = QtWidgets.QGroupBox(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.video_get_channel.sizePolicy().hasHeightForWidth())
        self.video_get_channel.setSizePolicy(sizePolicy)
        self.video_get_channel.setAlignment(QtCore.Qt.AlignCenter)
        self.video_get_channel.setCheckable(True)
        self.video_get_channel.setChecked(False)
        self.video_get_channel.setObjectName("video_get_channel")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.video_get_channel)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.label_8 = QtWidgets.QLabel(self.video_get_channel)
        self.label_8.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_8.setObjectName("label_8")
        self.gridLayout_3.addWidget(self.label_8, 0, 0, 1, 1)
        self.video_channel_left = QtWidgets.QSlider(self.video_get_channel)
        self.video_channel_left.setOrientation(QtCore.Qt.Horizontal)
        self.video_channel_left.setObjectName("video_channel_left")
        self.gridLayout_3.addWidget(self.video_channel_left, 0, 1, 1, 1)
        self.video_channel_right = QtWidgets.QSlider(self.video_get_channel)
        self.video_channel_right.setOrientation(QtCore.Qt.Horizontal)
        self.video_channel_right.setObjectName("video_channel_right")
        self.gridLayout_3.addWidget(self.video_channel_right, 1, 1, 3, 1)
        self.label_9 = QtWidgets.QLabel(self.video_get_channel)
        self.label_9.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_9.setObjectName("label_9")
        self.gridLayout_3.addWidget(self.label_9, 3, 0, 3, 1)
        self.label_10 = QtWidgets.QLabel(self.video_get_channel)
        self.label_10.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_10.setObjectName("label_10")
        self.gridLayout_3.addWidget(self.label_10, 2, 0, 1, 1)
        self.video_channel_top = QtWidgets.QSlider(self.video_get_channel)
        self.video_channel_top.setOrientation(QtCore.Qt.Horizontal)
        self.video_channel_top.setObjectName("video_channel_top")
        self.gridLayout_3.addWidget(self.video_channel_top, 4, 1, 1, 1)
        self.video_channel_bottom = QtWidgets.QSlider(self.video_get_channel)
        self.video_channel_bottom.setOrientation(QtCore.Qt.Horizontal)
        self.video_channel_bottom.setObjectName("video_channel_bottom")
        self.gridLayout_3.addWidget(self.video_channel_bottom, 7, 1, 1, 1)
        self.label_11 = QtWidgets.QLabel(self.video_get_channel)
        self.label_11.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_11.setObjectName("label_11")
        self.gridLayout_3.addWidget(self.label_11, 7, 0, 1, 1)
        self.gridLayout_8.addWidget(self.video_get_channel, 0, 2, 2, 1)
        self.verticalLayout.addLayout(self.gridLayout_8)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.groupBox_9 = QtWidgets.QGroupBox(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox_9.sizePolicy().hasHeightForWidth())
        self.groupBox_9.setSizePolicy(sizePolicy)
        self.groupBox_9.setMinimumSize(QtCore.QSize(109, 0))
        self.groupBox_9.setAlignment(QtCore.Qt.AlignCenter)
        self.groupBox_9.setCheckable(False)
        self.groupBox_9.setChecked(False)
        self.groupBox_9.setObjectName("groupBox_9")
        self.gridLayout_7 = QtWidgets.QGridLayout(self.groupBox_9)
        self.gridLayout_7.setObjectName("gridLayout_7")
        self.label_7 = QtWidgets.QLabel(self.groupBox_9)
        self.label_7.setScaledContents(False)
        self.label_7.setAlignment(QtCore.Qt.AlignCenter)
        self.label_7.setWordWrap(True)
        self.label_7.setObjectName("label_7")
        self.gridLayout_7.addWidget(self.label_7, 0, 0, 1, 1)
        self.min_value_for_areas = QtWidgets.QSpinBox(self.groupBox_9)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.min_value_for_areas.sizePolicy().hasHeightForWidth())
        self.min_value_for_areas.setSizePolicy(sizePolicy)
        self.min_value_for_areas.setMaximum(255)
        self.min_value_for_areas.setObjectName("min_value_for_areas")
        self.gridLayout_7.addWidget(self.min_value_for_areas, 1, 0, 1, 1)
        self.calc_areas = QtWidgets.QPushButton(self.groupBox_9)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.calc_areas.sizePolicy().hasHeightForWidth())
        self.calc_areas.setSizePolicy(sizePolicy)
        self.calc_areas.setObjectName("calc_areas")
        self.gridLayout_7.addWidget(self.calc_areas, 2, 0, 1, 1)
        self.areas_value = QtWidgets.QLabel(self.groupBox_9)
        self.areas_value.setText("")
        self.areas_value.setObjectName("areas_value")
        self.gridLayout_7.addWidget(self.areas_value, 3, 0, 1, 1)
        self.horizontalLayout_2.addWidget(self.groupBox_9)
        self.output = QtWidgets.QWidget(self.frame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.output.sizePolicy().hasHeightForWidth())
        self.output.setSizePolicy(sizePolicy)
        self.output.setMinimumSize(QtCore.QSize(261, 150))
        self.output.setObjectName("output")
        self.gridLayout_9 = QtWidgets.QGridLayout(self.output)
        self.gridLayout_9.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_9.setObjectName("gridLayout_9")
        self.output_layout = QtWidgets.QGridLayout()
        self.output_layout.setContentsMargins(0, 0, 0, 0)
        self.output_layout.setObjectName("output_layout")
        self.gridLayout_9.addLayout(self.output_layout, 0, 0, 1, 1)
        self.horizontalLayout_2.addWidget(self.output)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.verticalLayout_2.addLayout(self.verticalLayout)
        self.gridLayout_10.addLayout(self.verticalLayout_2, 0, 0, 1, 1)
        self.gridLayout_5.addWidget(self.frame_2, 0, 1, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 850, 21))
        self.menubar.setObjectName("menubar")
        self.menu_File = QtWidgets.QMenu(self.menubar)
        self.menu_File.setObjectName("menu_File")
        self.menuConnect = QtWidgets.QMenu(self.menubar)
        self.menuConnect.setObjectName("menuConnect")
        self.menuPCB = QtWidgets.QMenu(self.menuConnect)
        self.menuPCB.setObjectName("menuPCB")
        self.menuCamera = QtWidgets.QMenu(self.menuConnect)
        self.menuCamera.setObjectName("menuCamera")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.action_Open = QtWidgets.QAction(MainWindow)
        self.action_Open.setObjectName("action_Open")
        self.action_Save_as = QtWidgets.QAction(MainWindow)
        self.action_Save_as.setObjectName("action_Save_as")
        self.actionExit = QtWidgets.QAction(MainWindow)
        self.actionExit.setObjectName("actionExit")
        self.open_port = QtWidgets.QAction(MainWindow)
        self.open_port.setObjectName("open_port")
        self.read_port = QtWidgets.QAction(MainWindow)
        self.read_port.setObjectName("read_port")
        self.close_port = QtWidgets.QAction(MainWindow)
        self.close_port.setObjectName("close_port")
        self.open_camera = QtWidgets.QAction(MainWindow)
        self.open_camera.setObjectName("open_camera")
        self.close_camera = QtWidgets.QAction(MainWindow)
        self.close_camera.setObjectName("close_camera")
        self.menu_File.addAction(self.action_Open)
        self.menu_File.addAction(self.action_Save_as)
        self.menu_File.addAction(self.actionExit)
        self.menuPCB.addAction(self.open_port)
        self.menuPCB.addAction(self.read_port)
        self.menuPCB.addAction(self.close_port)
        self.menuCamera.addAction(self.open_camera)
        self.menuCamera.addAction(self.close_camera)
        self.menuConnect.addAction(self.menuPCB.menuAction())
        self.menuConnect.addAction(self.menuCamera.menuAction())
        self.menubar.addAction(self.menu_File.menuAction())
        self.menubar.addAction(self.menuConnect.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        #Camera
        camera_video_layer = QPixmap(self.camera_frame.width(), self.camera_frame.height())
        camera_video_layer.fill(QColor('Black'))
        self.camera_frame.setPixmap(camera_video_layer)
        self.settings_window = CameraSettings()
        self.camera = CameraBSI(self.settings_window)
        self.camera_setting.clicked.connect(self.open_setting_window)
        self.open_camera.triggered.connect(self.camera.initialize_camera)
        self.close_camera.triggered.connect(self.camera.close_camera)
        self.camera_rec.clicked.connect(self.camera.start_recording)

        #Experiment
        self.amperage_value.setSegmentStyle(QLCDNumber.Flat)
        self.experiment = Experiment()

        self.open_port.triggered.connect(self.experiment.open_port)
        self.read_port.triggered.connect(lambda: self.experiment.read_port(self.experiment.serialInst))
        self.close_port.triggered.connect(lambda: self.experiment.close_port(self.experiment.serialInst))

        #Video
        self.open_video.clicked.connect(self.open_video_file)
        video_layer = QPixmap(self.video_frame.width(), self.video_frame.height())
        video_layer.fill(QColor('Black'))
        self.video_frame.setPixmap(video_layer)
        self.video_create_graph.clicked.connect(self.create_graph_instance)

        #Graph
        self.calc_areas.clicked.connect(self.create_graph_instance)

        #Output
        self.save_xlsx_file.clicked.connect(self.create_output_file)

    def open_video_file(self):
        fname = QFileDialog.getOpenFileName() 
        self.video_path.setText(fname[0])
        self.video_player = VideoPlayer()
        self.video_timeline.setMaximum(int(self.video_player.frames))
        self.video_timeline.sliderMoved.connect(lambda:self.video_player.update_sliders_values(self.video_timeline_value, self.video_timeline.value()))
        self.video_timeline_value.setMaximum(self.video_player.frames)
        self.video_timeline_value.setProperty('value', 1)
        self.video_timeline_value.valueChanged.connect(lambda: self.video_player.update_sliders_values(self.video_timeline, self.video_timeline_value.value()))
        self.video_timeline_value.valueChanged.connect(lambda: self.video_player.set_frame(self.video_timeline_value.value()))
        #channel
        self.video_channel_left.setMaximum(int(self.video_player.video.get(cv2.CAP_PROP_FRAME_WIDTH)))
        self.video_channel_right.setMaximum(int(self.video_player.video.get(cv2.CAP_PROP_FRAME_WIDTH)))
        self.video_channel_right.setProperty('value', int(self.video_player.video.get(cv2.CAP_PROP_FRAME_WIDTH)))
        self.video_channel_top.setMaximum(int(self.video_player.video.get(cv2.CAP_PROP_FRAME_HEIGHT)))
        self.video_channel_bottom.setMaximum(int(self.video_player.video.get(cv2.CAP_PROP_FRAME_HEIGHT)))
        self.video_channel_bottom.setProperty('value', int(self.video_player.video.get(cv2.CAP_PROP_FRAME_HEIGHT)))

        self.video_pause.clicked.connect(self.video_player.play_video)
        self.video_mask.toggled.connect(self.video_player.is_mask)
        self.video_blure.toggled.connect(self.video_player.is_blure)
        self.video_slider_lower_limit.sliderMoved.connect(lambda: self.video_player.update_image(self.video_player.frame))
        self.video_slider_upper_limit.sliderMoved.connect(lambda: self.video_player.update_image(self.video_player.frame))
        self.video_slider_lower_limit.valueChanged.connect(lambda: self.video_player.update_image(self.video_player.frame))
        self.video_slider_upper_limit.valueChanged.connect(lambda: self.video_player.update_image(self.video_player.frame))
        self.video_spin_lower_limit.valueChanged.connect(lambda: self.video_player.update_sliders_values(self.video_slider_lower_limit, self.video_spin_lower_limit.value()))
        self.video_spin_upper_limit.valueChanged.connect(lambda: self.video_player.update_sliders_values(self.video_slider_upper_limit, self.video_spin_upper_limit.value()))
        self.video_channel_left.sliderMoved.connect(lambda: self.video_player.update_image(self.video_player.frame))
        self.video_channel_right.sliderMoved.connect(lambda: self.video_player.update_image(self.video_player.frame))
        self.video_channel_top.sliderMoved.connect(lambda: self.video_player.update_image(self.video_player.frame))
        self.video_channel_bottom.sliderMoved.connect(lambda: self.video_player.update_image(self.video_player.frame))
        self.video_get_channel.toggled.connect(self.video_player.is_get_channel)
        self.video_player.change_pixmap_signal.connect(self.video_player.update_image)
        self.video_frame.setAlignment(Qt.AlignCenter)
        self.video_player.start()


    def create_graph_instance(self):
        try:
            self.graph = Graph(self.video_player.frame)
            self.graph.create_output_data()
        except AttributeError:
            print('Select video')

    def create_output_file(self):
        self.output_data = OutputFile(self.graph.area, self.graph.data_table)
        self.output_data.save_output_file()

    def open_setting_window(self):
        self.settings_window.show()


    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label_4.setText(_translate("MainWindow", "Amperage, mA:"))
        self.camera_rec.setText(_translate("MainWindow", "REC"))
        self.camera_setting.setText(_translate("MainWindow", "Settings"))
        self.label_13.setText(_translate("MainWindow", "Src1"))
        self.label_15.setText(_translate("MainWindow", "Src3"))
        self.label_16.setText(_translate("MainWindow", "Src4"))
        self.label_12.setText(_translate("MainWindow", "Time"))
        self.label_14.setText(_translate("MainWindow", "Src2"))
        self.open_video.setText(_translate("MainWindow", "Open"))
        self.video_pause.setText(_translate("MainWindow", "Pause"))
        self.video_create_graph.setText(_translate("MainWindow", "Create graph"))
        self.save_xlsx_file.setText(_translate("MainWindow", "Save as \'.xlsx\'"))
        self.video_mask.setTitle(_translate("MainWindow", "Mask"))
        self.label_2.setText(_translate("MainWindow", "Upper Limit:"))
        self.label.setText(_translate("MainWindow", "Lower Limit:"))
        self.video_blure.setTitle(_translate("MainWindow", "Blure"))
        self.label_3.setText(_translate("MainWindow", "K"))
        self.label_6.setText(_translate("MainWindow", "Sigma"))
        self.video_get_channel.setTitle(_translate("MainWindow", "Get channel"))
        self.label_8.setText(_translate("MainWindow", "Left:"))
        self.label_9.setText(_translate("MainWindow", "Top:"))
        self.label_10.setText(_translate("MainWindow", "Right:"))
        self.label_11.setText(_translate("MainWindow", "Bottom:"))
        self.groupBox_9.setTitle(_translate("MainWindow", "Areas"))
        self.label_7.setText(_translate("MainWindow", "Min intensity value for calc areas:"))
        self.calc_areas.setText(_translate("MainWindow", "Calc"))
        self.menu_File.setTitle(_translate("MainWindow", "&File"))
        self.menuConnect.setTitle(_translate("MainWindow", "Connect"))
        self.menuPCB.setTitle(_translate("MainWindow", "PCB"))
        self.menuCamera.setTitle(_translate("MainWindow", "Camera"))
        self.action_Open.setText(_translate("MainWindow", "&Open"))
        self.action_Save_as.setText(_translate("MainWindow", "&Save as"))
        self.actionExit.setText(_translate("MainWindow", "Exit"))
        self.open_port.setText(_translate("MainWindow", "Open"))
        self.read_port.setText(_translate("MainWindow", "Read"))
        self.close_port.setText(_translate("MainWindow", "Close"))
        self.open_camera.setText(_translate("MainWindow", "Open"))
        self.close_camera.setText(_translate("MainWindow", "Close connection"))

class Ui_Settings(object):
    def setupUi(self, Settings):
        Settings.setObjectName("Settings")
        Settings.resize(300, 400)
        self.centralwidget = QtWidgets.QWidget(Settings)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.camera_port = QtWidgets.QComboBox(self.centralwidget)
        self.camera_port.setObjectName("camera_port")
        self.gridLayout.addWidget(self.camera_port, 0, 1, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 1, 0, 1, 1)
        self.camera_speed = QtWidgets.QLineEdit(self.centralwidget)
        self.camera_speed.setObjectName("camera_speed")
        self.gridLayout.addWidget(self.camera_speed, 1, 1, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 2, 0, 1, 1)
        self.camera_gain = QtWidgets.QLineEdit(self.centralwidget)
        self.camera_gain.setObjectName("camera_gain")
        self.gridLayout.addWidget(self.camera_gain, 2, 1, 1, 1)
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setObjectName("label_4")
        self.gridLayout.addWidget(self.label_4, 3, 0, 1, 1)
        self.camera_exposure_time = QtWidgets.QLineEdit(self.centralwidget)
        self.camera_exposure_time.setObjectName("camera_exposure_time")
        self.gridLayout.addWidget(self.camera_exposure_time, 3, 1, 1, 1)
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setObjectName("label_5")
        self.gridLayout.addWidget(self.label_5, 4, 0, 1, 1)
        self.camera_path_to_save = QtWidgets.QLineEdit(self.centralwidget)
        self.camera_path_to_save.setObjectName("camera_path_to_save")
        self.gridLayout.addWidget(self.camera_path_to_save, 4, 1, 1, 1)
        self.camera_select_path_to_save = QtWidgets.QToolButton(self.centralwidget)
        self.camera_select_path_to_save.setObjectName("camera_select_path_to_save")
        self.gridLayout.addWidget(self.camera_select_path_to_save, 4, 2, 1, 1)
        Settings.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(Settings)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 300, 21))
        self.menubar.setObjectName("menubar")
        Settings.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(Settings)
        self.statusbar.setObjectName("statusbar")
        Settings.setStatusBar(self.statusbar)

        self.retranslateUi(Settings)
        QtCore.QMetaObject.connectSlotsByName(Settings)

        self.camera_path_to_save.setText(os.getcwd())

    def retranslateUi(self, Settings):
        _translate = QtCore.QCoreApplication.translate
        Settings.setWindowTitle(_translate("Settings", "MainWindow"))
        self.label.setText(_translate("Settings", "Port"))
        self.label_2.setText(_translate("Settings", "Speed"))
        self.label_3.setText(_translate("Settings", "Gain"))
        self.label_4.setText(_translate("Settings", "Exposure time"))
        self.label_5.setText(_translate("Settings", "Path to save"))
        self.camera_select_path_to_save.setText(_translate("Settings", "..."))

class CameraSettings(QtWidgets.QMainWindow, Ui_Settings):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.retranslateUi(self)
        self.exp_time = 0.5
        self.clear_mode = "Never"
        self.exp_mode = "Ext Trig Trig First"
        self.readout_port = 0
        self.speed_table_index = 0
        self.gain = 1
        self.path_to_save = os.getcwd()

        self.camera_select_path_to_save.clicked.connect(self.set_path_to_save)

    def set_path_to_save(self):
        path = QFileDialog.getExistingDirectory()
        self.camera_path_to_save.setText(path)

class CameraBSI:
    def __init__(self, settings):
        self.cam = None
        #Settings
        self.settings = settings
        self.exp_time = settings.exp_time
        self.clear_mode = settings.clear_mode
        self.exp_mode = settings.exp_mode
        self.readout_port = settings.readout_port
        self.speed_table_index = settings.speed_table_index
        self.gain = settings.gain

    def initialize_camera(self):
        pvc.init_pvcam()   
        try:
            self.cam = next(Camera.detect_camera())
            self.cam.open()
            self.show_live()
        except:
            print('No available camera found')    

    def close_camera(self):
        self.cam.close()
        pvc.uninit_pvcam()
    ###########################
    def show_live(self):
        self.cam.start_live(exp_time=self.exp_time)
        cnt = 0
        tot = 0
        t1 = time.time()
        start = time.time()
        width = 800
        height = int(self.cam.sensor_size[1] * width / self.cam.sensor_size[0])
        dim = (width, height)
        fps = 0

        while True:
            frame, fps, frame_count = self.cam.poll_frame()
            frame['pixel_data'] = cv2.resize(frame['pixel_data'], dim, interpolation = cv2.INTER_AREA)
            cv2.imshow('Live Mode', frame['pixel_data'])

            self.update_image(frame['pixel_data'])

            low = np.amin(frame['pixel_data'])
            high = np.amax(frame['pixel_data'])
            average = np.average(frame['pixel_data'])

            if cnt == 10:
                t1 = time.time() - t1
                fps = 10/t1
                t1 = time.time()
                cnt = 0
            if cv2.waitKey(10) == 27:
                break
            print('Min:{}\tMax:{}\tAverage:{:.0f}\tFrame Rate: {:.1f}\n'.format(low, high, average, fps))
            cnt += 1
            tot += 1

        self.cam.finish()
        print('Total frames: {}\nAverage fps: {}\n'.format(tot, (tot/(time.time()-start))))
    ###############

    @pyqtSlot(np.ndarray)
    def update_image(self, cv_img):
        qt_img = self.convert_cv_qt(cv_img)
        ui.camera_frame.setPixmap(qt_img)

    def convert_cv_qt(self, cv_img):
        image = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
        h, w = image.shape
        bytes_per_line = w
        convert_to_Qt_format = QtGui.QImage(image.data, w, h, bytes_per_line, QtGui.QImage.Format_Grayscale8)
        p = convert_to_Qt_format.scaled(400, 250, Qt.KeepAspectRatio)
        return QPixmap.fromImage(p)
    
    def get_single_image(self):
        frame = self.cam.get_frame(exp_time=self.exp_time)
        print("First five pixels of frame: {}, {}, {}, {}, {}".format(*frame[:5]))

    # def start_recording(self):
    #     NUM_FRAMES = 200
    #     FRAME_DATA_PATH = self.settings.path_to_save
    #     BUFFER_FRAME_COUNT = 16
    #     WIDTH = 1000
    #     HEIGHT = 1000

    #     self.cam.meta_data_enabled = True
    #     self.cam.set_roi(0, 0, WIDTH, HEIGHT)
    #     self.cam.start_live(exp_time=100, buffer_frame_count=BUFFER_FRAME_COUNT, stream_to_disk_path=FRAME_DATA_PATH)

    #     # Data is streamed to disk in a C++ call-back function invoked directly by PVCAM. To not overburden the system,
    #     # only poll for frames in python at a slow rate, then exit when the frame count indicates all frames have been
    #     # written to disk
    #     while True:
    #         frame, fps, frame_count = self.cam.poll_frame()

    #         if frame_count >= NUM_FRAMES:
    #             low = np.amin(frame['pixel_data'])
    #             high = np.amax(frame['pixel_data'])
    #             average = np.average(frame['pixel_data'])
    #             print('Min:{}\tMax:{}\tAverage:{:.0f}\tFrame Count:{:.0f} Frame Rate: {:.1f}'.format(low, high, average, frame_count, fps))
    #             break

    #         time.sleep(1)

    #     cam.finish()

    #     imageFormat = cam.get_param(const.PARAM_IMAGE_FORMAT)
    #     if imageFormat == const.PL_IMAGE_FORMAT_MONO8:
    #         BYTES_PER_PIXEL = 1
    #     else:
    #         BYTES_PER_PIXEL = 2

    def stop_recording(self):
        pass

    def check_status(self):
        pass

class PCB:
    def __init__(self, port):
        self.port = port
        self.is_on = False
        self.const_voltage = 0.132
        self.const_amperage = 0.1957

    def on_off_power_supply(self):
        if self.is_on:
            payload_value = 0
            self.is_on = False
        else:
            payload_value = 1
            self.is_on = True
        message = [0xAA, 0x55, 0x00, payload_value, 0xAA^0x55^0x00^payload_value]
        tmpBuffer = bytearray(message)
        self.port.write(tmpBuffer)

    def set_voltage(self, src, value):
        if self.is_on:
            operation = src
            message = [0xAA, 0x55, operation, value, 0xAA^0x55^operation^value]
            tmpBuffer = bytearray(message)
            self.port.write(tmpBuffer)
        else:
            print('Power supply turned off')

    def get_status_pcb(self):
        status = self.port.read(15)
        print(status)

class Experiment:

    def __init__(self):
        self.output_data_port = []
        self.serialInst = {}
        self.amperage = 0
        self.voltage = {
            'ch1': 0,
            'ch2': 0,
            'ch3': 0,
            'ch4': 0
        }
        self.PCB = None

    def set_serial_inst(self, port):
        self.serialInst = port

    def close_port(self, port): 
        port.close()

    def open_port(self):
        serialInst = serial.Serial()
        serialInst.port = 'COM3'
        serialInst.baudrate = 9600
        serialInst.open()
        self.serialInst = serialInst
        self.PCB = PCB(serialInst)

    def get_ports_list(self):
        portsList = []
        ports = serial.tools.list_ports.comports()
        for onePort in ports:
            portsList.append(str(onePort))
        return portsList

    def read_port_thread(self, port):
        while port.isOpen():
            packet = port.readline()
            packet = int(packet.rstrip().decode('UTF-8'))
            self.output_data_port.append(packet)
            self.update_amperage_value(self.output_data_port)

    def read_port(self, port):
        threading.Thread(target=self.read_port_thread, args=(port, ), daemon=True).start()

    def update_amperage_value(self, new_data):
        current_value = new_data[-1]
        ui.amperage_value.display(f'{current_value}')


    def start_recording(self):
        pass
    
    def is_limit_for_recording(self, current_value, limit):
        return current_value >= limit

class VideoPlayer(QThread):
    change_pixmap_signal = pyqtSignal(np.ndarray)

    def __init__(self):
        super().__init__()
        self.play = True
        self.mask = False
        self.blure = False
        self.get_channel = False

        self.frame_count = 1

        self.video = cv2.VideoCapture(ui.video_path.text())
        self.frames = int(self.video.get(cv2.CAP_PROP_FRAME_COUNT))
        self.fps = self.video.get(cv2.CAP_PROP_FPS)

    def run(self):
        while self.video.isOpened():
            while self.play:
                time.sleep(1/self.fps)
                ret, cv_img = self.video.read()
                if ret:
                    self.frame = cv_img
                    self.change_pixmap_signal.emit(cv_img)
                    self.frame_count = self.frame_count + 1
                    self.update_frame_timeline(self.frame_count)
                else:
                    self.play = False
                    ui.video_pause.setText('Play')


    @pyqtSlot(np.ndarray)
    def update_image(self, cv_img):
        qt_img = self.convert_cv_qt(cv_img)
        ui.video_frame.setPixmap(qt_img)

    def convert_cv_qt(self, cv_img):
        image = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
        if self.mask:
            image = self.video_mask(image)
        if self.blure:
            image = self.video_blure(image)
        if self.get_channel:
            image = self.video_get_channel(image)
        h, w = image.shape
        bytes_per_line = w
        convert_to_Qt_format = QtGui.QImage(image.data, w, h, bytes_per_line, QtGui.QImage.Format_Grayscale8)
        p = convert_to_Qt_format.scaled(400, 250, Qt.KeepAspectRatio)
        return QPixmap.fromImage(p)

    def play_video(self):
        if self.play:
            self.play = False
            ui.video_pause.setText('Play')
        else: 
            self.play =True
            ui.video_pause.setText('Pause')

    def is_mask(self):
        if self.mask:
            self.mask = False
            if not self.play:
                self.update_image(self.frame)
        else:
            self.mask = True
            if not self.play:
                self.update_image(self.frame)

    def video_mask(self, frame):
        ui.video_spin_lower_limit.setValue(ui.video_slider_lower_limit.value())
        ui.video_spin_upper_limit.setValue(ui.video_slider_upper_limit.value())
        mask = cv2.inRange(frame, ui.video_slider_lower_limit.value(), ui.video_slider_upper_limit.value())
        frame = cv2.bitwise_and(frame, frame, mask=mask)
        return frame

    def is_blure(self):
        if self.blure:
            self.blure = False
            if not self.play:
                self.update_image(self.frame)
        else:
            self.blure = True
            if not self.play:
                self.update_image(self.frame)

    def video_blure(self, frame):
        k = ui.video_blure_k.value()
        sigma = ui.video_blure_sigma.value()
        frame = cv2.GaussianBlur(frame, (k, k), sigma)
        return frame

    def is_get_channel(self):
        if self.get_channel:
            self.get_channel = False
            if not self.play:
                self.update_image(self.frame)
        else:
            self.get_channel = True
            if not self.play:
                self.update_image(self.frame)

    def video_get_channel(self, frame):
        return cv2.rectangle(frame, (ui.video_channel_left.value(), ui.video_channel_top.value()), (
            ui.video_channel_right.value(), ui.video_channel_bottom.value()), 255, 3)

    def set_frame(self, new_frame):
        if self.video:
            if self.play == False:
                ret, frame = self.go_to_frame(new_frame)
                self.frame_count = new_frame
                self.update_frame_timeline(self.frame_count)
                if ret:
                    self.update_image(frame)


    def go_to_frame(self, new_frame):
        if self.video.isOpened():
            self.video.set(cv2.CAP_PROP_POS_FRAMES, new_frame)
            ret, frame = self.video.read()
            self.frame = frame
            if ret:
                return ret, frame
            else:
                return ret, None
        else:
            return 0, None

    def update_sliders_values(self, slider, new_value):
        slider.setValue(new_value)

    def update_frame_timeline(self, value):
        ui.video_timeline_value.setValue(value)

class MplCanvas(FigureCanvasQTAgg):

    def __init__(self):
        self.fig = Figure()
        # self.fig, self.axes = plt.subplots()
        self.axes = self.fig.subplots()
        super(MplCanvas, self).__init__(self.fig)

class Graph:
    def __init__(self, frame):
        self.data_graph = {
            'width': None,
            'intensity_width': []
        }
        self.area = []
        self.array_figure = []
        self.frame = frame
        self.min_intensity_value = ui.min_value_for_areas.value()
        self.sc = MplCanvas()

    def handle_img(self, img):
        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        img = img[
            ui.video_channel_top.value() : ui.video_channel_bottom.value(), 
            ui.video_channel_left.value() : ui.video_channel_right.value()
            ]
        mask = cv2.inRange(img, ui.video_slider_lower_limit.value(), ui.video_slider_upper_limit.value())
        img = cv2.bitwise_and(img, img, mask=mask)
        return img

    def calc_data_graph(self, img):
        self.data_graph['intensity_width'] = np.mean(img, axis=0)
        self.data_graph['width'] = np.arange(1, np.size(img, 1) + 1, 1)
        self.data_table = {
            'x': self.data_graph['width'],
            'y': self.data_graph['intensity_width']
        }   

    def get_figures(self, x1, array_length):
        if x1 > array_length:
            return
        
        for i in range (x1, array_length):
            if self.data_graph['intensity_width'][i] > self.min_intensity_value:
                left_dot = i
                for j in range(i, len(self.data_graph['intensity_width'])):
                    if j == len(self.data_graph['intensity_width']) - 1:
                        right_dot = j
                        self.array_figure.append((left_dot, right_dot))
                        return
                    elif self.data_graph['intensity_width'][j] <= self.min_intensity_value:
                        right_dot = j  - 1
                        self.array_figure.append((left_dot, right_dot))
                        self.get_figures(right_dot + 1, array_length)
                        return

    def calc_area(self, main_array, array_figure):
        area_array = []
        figures =[]
        new_main_array = []
        for i in main_array:
            new_main_array.append(i - self.min_intensity_value)
        for i in array_figure:
            left_dot, right_dot = i
            figures.append(new_main_array[left_dot : right_dot])
        for i in figures:
            area_array.append(round(trapz(i, dx=1), 1))
        return area_array

    def filter_area(self, area):
        if area > 100:
            return True
        else:
            return False

    def update_output_area_element(self):
        self.get_figures(1, len(self.data_graph['intensity_width']))
        self.area = self.calc_area(self.data_graph['intensity_width'], self.array_figure)
        self.area = filter(self.filter_area, self.area)
        self.area = [*self.area]
        area_element = ''
        count = 1
        for i in range(0, len(self.area)):
            area_element += f'Area{count}: {self.area[i]}\n'
            count += 1
        ui.areas_value.setText(f'{area_element}')

    def create_output_data(self):
        img = self.frame
        img = self.handle_img(img)
        self.calc_data_graph(img) 
        self.update_output_area_element()
        self.sc.axes.grid()
        y1 = self.data_graph['intensity_width']
        y2 = np.full(len(self.data_graph['width']), self.min_intensity_value)
        self.sc.axes.fill_between(self.data_graph['width'], y1, y2, where = (y1 > y2))
        self.sc.axes.plot(self.data_graph['width'], self.data_graph['intensity_width'])
        self.sc.axes.plot(np.full(len(self.data_graph['width']), self.min_intensity_value))
        self.sc.fig.savefig('graph.png')
        for i in range(ui.output_layout.count()): ui.output_layout.itemAt(i).widget().close()
        ui.output_layout.addWidget(self.sc)

        OutputFile(self.area, self.data_table)

class OutputFile:
    def __init__(self, output_data, table):
        self.output_data = output_data
        self.table = table
        self. wb = Workbook()

    def create_data_style(self, name, bold, font_size):
        ns = NamedStyle(name=name)
        ns.font = Font(bold=bold, size=font_size)
        border = Side(style='thin', color='000000')
        ns.border = Border(left=border, top=border, right=border, bottom=border)
        ns.alignment = Alignment(horizontal="center", vertical="center")
        self.wb.add_named_style(ns)

    def insert_graph(self):
        self.wb.create_sheet(title = 'Intensity signal', index = 0)

        self.create_data_style('highlight', True, 18)
        self.create_data_style('table', False, 12)

        self.wb['Intensity signal'].column_dimensions['B'].width = 30

        img = openpyxl.drawing.image.Image('graph.png')
        img.anchor = 'D2'

        self.wb['Intensity signal'].add_image(img)

        for i in range(0, len(self.output_data)):
            self.wb['Intensity signal'][f'B{2 + 2 * i}'].style = 'highlight'
            self.wb['Intensity signal'][f'B{3 + 2 * i}'].style = 'highlight'
            self.wb['Intensity signal'][f'B{2 + 2 * i}'] = f'Area{i + 1}'
            self.wb['Intensity signal'][f'B{3 + 2 * i}'] = '{0:,}'.format(self.output_data[i]).replace(',', ' ')

        self.wb['Intensity signal'][f'A{5 + (len(self.output_data) - 1) * 2}'] = 'Distance'
        self.wb['Intensity signal'][f'B{5 + (len(self.output_data) - 1) * 2}'] = 'Signal Intensity'
        self.wb['Intensity signal'][f'A{5 + (len(self.output_data) - 1) * 2}'].style = 'table'
        self.wb['Intensity signal'][f'B{5 + (len(self.output_data) - 1) * 2}'].style = 'table'

        shift = 6 + ((len(self.output_data) - 1) * 2)
        for i in range(0, len(self.table['x'])):
            self.wb['Intensity signal'][f'A{i + shift}'] = self.table['x'][i]
            self.wb['Intensity signal'][f'B{i + shift}'] = self.table['y'][i]
            self.wb['Intensity signal'][f'A{i + shift}'].style = 'table'
            self.wb['Intensity signal'][f'B{i + shift}'].style = 'table'

    def save_output_file(self):
        self.insert_graph()
        path = QFileDialog.getSaveFileName()[0]
        self.wb.save(path+'.xlsx')
        os.remove('graph.png')

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
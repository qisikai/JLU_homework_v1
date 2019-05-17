#encoding=utf-8
import sys
import os
import datetime
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = "D:/python3.7/Lib/site-packages/PyQt5/Qt/plugins"#qt_platform_plugins_path
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog,QPushButton,QTabWidget,QWidget,QHBoxLayout,QMessageBox,QListWidgetItem,QListWidget
from MainWindow import Ui_MainWindow

from PyQt5 import QtCore
from PyQt5.QtCore import QUrl,QThread,QTimer,pyqtSignal,QCoreApplication,QDate,QTime
import time
from PyQt5.QtWebEngineWidgets import QWebEngineView,QWebEngineSettings,QWebEngineProfile
from PyQt5.QtGui import QIcon,QStandardItemModel,QStandardItem,QBrush,QColor
import json
import random
import math

import utils

class MainUI(QMainWindow, Ui_MainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setWindowIcon(QIcon('logo.jpg'))
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.setWindowTitle("JLU 课程设计")
        self.setFixedSize(self.width(), self.height())
        self._translate = QtCore.QCoreApplication.translate

        self.timeEdit.setEnabled(False)
        self.calendarWidget.setEnabled(False)

        self.timeEdit.setTime(QTime(10, 0, 0))

        self.table_model = QStandardItemModel()
        self.table_model.setHorizontalHeaderLabels(
            ["座位ID", "姓名", "专业","座位状态"])
        self.tableView.setModel(self.table_model)
        self.tableView.setColumnWidth(0, 50)
        self.tableView.setColumnWidth(2,140)
        self.tableView.setColumnWidth(3, 140)
        self.time_table={
            "1": "08:00-09:40",
            "2": "08:00-09:40",
            "3": "10:00-11:40",
            "4": "10:00-11:40",
            "5": "13:30-15:10",
            "6": "13:30-15:10",
            "7": "15:30-17:10",
            "8": "15:30-17:10",
            "9": "18:00-19:40",
            "10": "18:00-19:40",
        }
        self.time_table_detail={}
        for class_ind in self.time_table:
            start,end = self.time_table[class_ind].split("-")
            start_h,start_m = start.split(":")
            end_h,end_m = end.split(":")
            self.time_table_detail[class_ind]={
                "start_h":int(start_h),
                "start_m":int(start_m),
                "end_h":int(end_h),
                "end_m":int(end_m)
            }
        print(self.time_table_detail)



        self.classes = None
        self.students = None
        self.infoLabel.setVisible(False)
        self.goBtn.setEnabled(False)
        self.loadConfigBtn.setEnabled(False)
        self.loadCourseBtn.clicked.connect(self.doOpenCourseFile)
        self.loadConfigBtn.clicked.connect(self.doOpenConfig)
        self.goBtn.clicked.connect(self.doClickGoBtn)

    def doOpenConfig(self):
        file_, filetype = QFileDialog.getOpenFileName(self,
                                                      "选取文件",
                                                      ".",
                                                      "Excel xls (*.xls);;Excel xlsx (*.xlsx)")  # 设置文件扩展名过滤,注意用双分号间隔
        if not file_ or len(file_) == 0:
            QMessageBox.information(self, "选择文件",
                                    self.tr("未选择文件"))
            return
        res = utils.parse_config(file_)
        if res != None:
            self.rooms = res
            self.goBtn.setEnabled(True)
            start_date = datetime.datetime(2018,9,3)
            max_week = 0
            for room_id in self.rooms:
                students = self.rooms[room_id]
                for student in students:
                    if len(student['name']) == 0:
                        continue
                    cls  =student['class']
                    if cls in self.classes:
                        for course_name in self.classes[cls]:
                            for course in self.classes[cls][course_name]:
                                max_week = max(max_week,max(course['weeks']))

            end_date = start_date + datetime.timedelta(days = max_week * 7)
            self.calendarWidget.setMinimumDate(QDate(start_date.year, start_date.month, start_date.day))
            self.calendarWidget.setMaximumDate(QDate(end_date.year, end_date.month, end_date.day))
            self.calendarWidget.setEnabled(True)
            self.timeEdit.setEnabled(True)
            self.room_ids.clear()
            for room_id in self.rooms.keys():
                self.room_ids.addItem("阅览室_%s" % room_id)

            QMessageBox.information(self, "提示",
                                    self.tr("提取配置 成功"))
            return
        else:
            QMessageBox.information(self, "提示",
                                    self.tr("提取配置 失败"))
            return


    def doOpenCourseFile(self):
        file_, filetype = QFileDialog.getOpenFileName(self,
                                                      "选取文件",
                                                      ".",
                                                      "Excel xls (*.xls);;Excel xlsx (*.xlsx)")  # 设置文件扩展名过滤,注意用双分号间隔
        if not file_ or len(file_) == 0:
            QMessageBox.information(self, "选择文件",
                                    self.tr("未选择文件"))
            return
        res =  utils.parser_class_all(file_)
        if res != None:
            self.classes = res
            self.loadConfigBtn.setEnabled(True)
            self.calendarWidget.setEnabled(False)
            self.timeEdit.setEnabled(False)
            self.goBtn.setEnabled(False)

            QMessageBox.information(self, "提示",
                                    self.tr("提取课程表文件 成功"))
            return
        else:
            QMessageBox.information(self, "提示",
                                    self.tr("提取课程表文件内容失败"))
            return
    
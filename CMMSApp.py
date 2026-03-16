# -*- coding: utf-8 -*-
from UI.View_result import Ui_View_result
from UI.df_show_windown import df_show
from UI.Setting_Windown import Ui_SettingWindown
from UI.MainWindown import Ui_MainWindow
from UI.Result_chart import Ui_Result_chart
from UI.Machine_detail import Ui_Machine_detail
from UI.Print_select import Ui_print_selector
from UI.Printing_progress import Ui_printing_progress
from UI.Form_modification import Ui_Form_Modification
from UI.Sign_in import Ui_Login
from UI.Update_machine_info import Ui_Update_machine_info
from UI.Sync_missing_data import Ui_Sync_Missing_Data
from UI.Downtime_input_window import Ui_DowntimeInputWindow
from UI.Group_choose import Ui_Group_choose
from UI.Error_code_management import Ui_Error_Code_Management
from Calculation.OEE_cal_result import OEE_result
from Database.MariaDB import Database_process
from Stock_control.stock_delegate import StockItemDelegate,ImageCache
from Stock_control.image_loader import ImageLoaderRunnable
from Maintenance.printer import Printer_process
from Maintenance.scan_qrcode import Scan_record_process
from Maintenance.attached_equipment import DynamicSuggestion
from Downtimes.Excel_processing import Downtime_Excel_Processor
# import plotly.graph_objects as go
# from plotly.subplots import make_subplots
# from PyQt5.QtWebEngineWidgets import QWebEngineView
# from PyQt5.QtWebChannel import QWebChannel
# from PyQt5 import QtWebEngineWidgets
import pyqtgraph as pg
import numpy as np
import sys
import os
import pandas as pd
import json
import requests
import bcrypt
from PyQt5 import QtWidgets, QtCore, QtGui, sip
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
import fitz, shutil
import datetime as dt
import re
from dateutil.relativedelta import relativedelta
from pyqtspinner.spinner import WaitingSpinner
from sqlalchemy import text
STRICT_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class OEEAppWindow(QtWidgets.QMainWindow):
    def __init__(self,login_info = None):
        super().__init__()
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setup_signals()
        self.login_info = login_info
        self.Setting_windown = None
        self.View_result_windown = None
        self.df_windown = None
        self.df = None
        self.Open_Setting_windown_Flag = False
        self.Flag_data_process = False
        self.list_df_molding_result = None
        self.is_expanded = False
        self.animation_group = QtCore.QParallelAnimationGroup()
        self.week_num = self.ui.company_week_number(self.ui.today)
        self.month_num = self.ui.today.month
        self.year_num = self.ui.today.year
        self.qty_week = self.ui.company_week_number(dt.date(self.year_num,12,31))
        self.spinner = WaitingSpinner(self, center_on_parent=True, disable_parent_when_spinning=True,speed=1.1)
        self.spinner.roundness = 70.0
        self.spinner.line_length = 30
        self.spinner.line_width = 10
        self.spinner.inner_radius = 40
        self.spinner.number_of_lines = 100
        self.spinner.color = QtGui.QColor(68, 60, 113)
        # self.ui.OEE_btn.setEnabled(False)
        # self.ui.Stock_btn.setEnabled(False)
        # self.ui.Order_btn.setEnabled(False)
        
        # self.ui.Downtime_btn.setEnabled(False)

    def setup_signals(self):
        self.ui.Show_file_bt_FG.clicked.connect(
            lambda: self.handle_df_show(self.ui.Text_FG))
        self.ui.Show_file_bt_NG.clicked.connect(
            lambda: self.handle_df_show(self.ui.Text_NG))
        self.ui.Setting_bt.clicked.connect(self.Open_Setting_windown)
        self.ui.Data_process_bt.clicked.connect(self.Data_process)
        self.ui.Export_Excel_bt.clicked.connect(
            lambda: self.Choose_export_machine(file_export=None, where="Excel"))
        self.ui.View_data_bt.clicked.connect(self.Open_View_result_windown)
        self.ui.Home_btn.clicked.connect(self.Home_page)
        self.ui.OEE_btn.clicked.connect(self.OEE_page)
        self.ui.Maintenance_btn.clicked.connect(self.Maintenance_page)
        self.ui.Order_btn.clicked.connect(self.Part_order_page)
        self.ui.Stock_btn.clicked.connect(self.Stock_control_page)
        self.ui.Downtime_btn.clicked.connect(self.Downtime_page)
        self.ui.Main_Home_btn.clicked.connect(self.Mainten_Home_page)
        self.ui.Main_Input_record_btn.clicked.connect(self.Mainten_Input_page)
        self.ui.Main_Print_record_btn.clicked.connect(self.Mainten_Print_page)
        self.ui.Main_detail_plan_btn.clicked.connect(self.Mainten_Detail_plan_page)
        self.ui.filter_mainten_btn.clicked.connect(self.show_filter)
        self.ui.reset_filter_mainten_btn.clicked.connect(self.reset_filter)
        self.ui.weekly_btn.clicked.connect(self.monitor_week_page)
        self.ui.monthly_btn.clicked.connect(self.monitor_month_page)
        self.ui.inyear_btn.clicked.connect(self.monitor_inyear_page)
        self.ui.monitor_next_btn.clicked.connect(self.next_monitor_page)
        self.ui.monitor_back_btn.clicked.connect(self.back_monitor_page)
        self.ui.filter_stock_btn.clicked.connect(self.show_filter_stock)
        self.ui.reset_filter_stock_btn.clicked.connect(self.reset_filter_stock)
        self.ui.Group_cbb_PF.currentIndexChanged.connect(self.add_item_line_PF)
        self.ui.profile_btn.clicked.connect(lambda _: self.ui.frame_60.show()  if self.ui.frame_60.isHidden() else self.ui.frame_60.hide())
        self.ui.return_home_btn.clicked.connect(lambda _: self.return_home())
        self.ui.user_info_btn.clicked.connect(lambda _: self.user_info())
        self.ui.change_password_btn.clicked.connect(lambda _: self.change_password(form_user_info= False))
        self.ui.change_password_inside_btn.clicked.connect(lambda _: self.change_password(form_user_info=True))
    
    def _init_database(self):
        try:
            self.database_process = Database_process()
            self.group = self.database_process.query( sql=''' SELECT department_name FROM `Departments` ''' )
        except ConnectionError as e:
            QtWidgets.QMessageBox.critical(self, "Error", str(e))
            self.close()
    
    def safe_connect(self,signal, slot):
        try:
            signal.disconnect()
        except TypeError:
            pass
        signal.connect(slot)

#==========================Function of Maintenance page ==================================================================================BEGIN
#==========================Function of Maintenance page ==================================================================================BEGIN
#==========================Function of Maintenance page ==================================================================================BEGIN

    def expand_windown_animation(self,is_expand = False):
        size_animation = QtCore.QPropertyAnimation(self, b"size")
        size_animation.setDuration(250)
        size_animation2 = QtCore.QPropertyAnimation(self.ui.func_frame, b"size")
        size_animation2.setDuration(250)
        size_animation3 = QtCore.QPropertyAnimation(self.ui.main_stacked, b"size")
        size_animation3.setDuration(250)
        size_animation4 = QtCore.QPropertyAnimation(self.ui.Mainten_widget, b"size")
        size_animation4.setDuration(250)
        size_animation5 = QtCore.QPropertyAnimation(self.ui.Mainten_frame, b"size")
        size_animation5.setDuration(250)
        size_animation6 = QtCore.QPropertyAnimation(self.ui.Maintenance_stacked, b"size")
        size_animation6.setDuration(250)
        pos_animation = QtCore.QPropertyAnimation(self, b"pos", self)
        pos_animation.setDuration(250)
        pos_animation.setEasingCurve(QtCore.QEasingCurve.OutCubic)
        if is_expand:
            size_animation.setStartValue(QtCore.QSize(932, 545))
            size_animation.setEndValue(QtCore.QSize(1500, 800))
            size_animation2.setStartValue(QtCore.QSize(121, 551))
            size_animation2.setEndValue(QtCore.QSize(121, 800))
            size_animation3.setStartValue(QtCore.QSize(811, 551))
            size_animation3.setEndValue(QtCore.QSize(1379, 800))
            size_animation4.setStartValue(QtCore.QSize(811, 551))
            size_animation4.setEndValue(QtCore.QSize(1379, 800))
            size_animation5.setStartValue(QtCore.QSize(811, 471))
            size_animation5.setEndValue(QtCore.QSize(1379, 720))
            size_animation6.setStartValue(QtCore.QSize(811, 431))
            size_animation6.setEndValue(QtCore.QSize(1379, 680))
            pos_animation.setStartValue(self.pos())
            pos_animation.setEndValue(QtCore.QPoint(100, 100))
            self.animation_group.stop()
            self.animation_group.clear()
            self.animation_group.addAnimation(size_animation)
            self.animation_group.addAnimation(size_animation2)
            self.animation_group.addAnimation(size_animation3)
            self.animation_group.addAnimation(size_animation4)
            self.animation_group.addAnimation(size_animation5)
            self.animation_group.addAnimation(size_animation6)
            self.animation_group.addAnimation(pos_animation)
            self.animation_group.start()
        else:
            size_animation.setStartValue(QtCore.QSize(1500, 800))
            size_animation.setEndValue(QtCore.QSize(932, 545))
            size_animation2.setStartValue(QtCore.QSize(121, 800))
            size_animation2.setEndValue(QtCore.QSize(121, 551))
            size_animation3.setStartValue(QtCore.QSize(1379, 800))
            size_animation3.setEndValue(QtCore.QSize(811, 551))
            size_animation4.setStartValue(QtCore.QSize(1379, 800))
            size_animation4.setEndValue(QtCore.QSize(811, 551))
            size_animation5.setStartValue(QtCore.QSize(1379, 720))
            size_animation5.setEndValue(QtCore.QSize(811, 471))
            size_animation6.setStartValue(QtCore.QSize(1379, 680))
            size_animation6.setEndValue(QtCore.QSize(811, 431))
            pos_animation.setStartValue(self.pos())
            pos_animation.setEndValue(QtCore.QPoint(493, 212))
            self.animation_group.stop()
            self.animation_group.clear()
            self.animation_group.addAnimation(size_animation)
            self.animation_group.addAnimation(size_animation2)
            self.animation_group.addAnimation(size_animation3)
            self.animation_group.addAnimation(size_animation4)
            self.animation_group.addAnimation(size_animation5)
            self.animation_group.addAnimation(size_animation6)
            self.animation_group.addAnimation(pos_animation)
            self.animation_group.start()

    def set_stylesheet_change_page(self,button:tuple):
        button[0].setStyleSheet('''
                                    QPushButton {
                                        background-color: rgba(0, 0, 255, 0.07);
                                        border: none;                     
                                        border-top: 1px solid rgba(0, 0, 255, 1);
                                        border-bottom: 1px solid rgba(0, 0, 255, 1);
                                    }
        ''')
        for i in range(1,len(button)):
            button[i].setStyleSheet('''
                                    QPushButton {
                                                background-color: transparent;
                                                border: none;
                                                    }
                                    QPushButton:hover {
                                                background-color: rgba(0, 0, 255, 0.07);
                                                        }
                                    ''')
    
    @QtCore.pyqtSlot()
    def Home_page(self):
        self.notification_list = []
        self.ui.main_stacked.setCurrentWidget(self.ui.Home_page)
        self.set_stylesheet_change_page((self.ui.Home_btn,self.ui.OEE_btn,self.ui.Maintenance_btn,self.ui.Order_btn, self.ui.Stock_btn,self.ui.Downtime_btn))
        if self.is_expanded:
            self.is_expanded = False
            self.expand_windown_animation(self.is_expanded)
        if self.login_info is not None:
            self.ui.welcome_label.setText(f"Hello {self.login_info['first_name']} {self.login_info['last_name']}")
            self.ui.profile_btn.setText(f"{self.login_info['last_name']}") 
            self.ui.user_id_lbl.setText(str(self.login_info['user_id']))
            self.ui.user_name_lbl.setText(self.login_info['user_name'])
            self.ui.password_lnedit.setText("*********")
            self.ui.first_name_lnedit.setText(self.login_info['first_name'])
            self.ui.last_name_lnedit.setText(self.login_info['last_name'])
            self.ui.group_lbl.setText(self.login_info['department'])
            self.ui.position_lbl.setText(self.login_info['role_level'])
        try:
            notifications = self.database_process.query(sql='''
                                                            SELECT * FROM `Notifications`
                                                            WHERE ( receiver_id = :id or receiver_id IS NULL ) AND STATUS NOT IN ('CLOSE','REJECTED','ACCEPTED')
                                                            ORDER BY created_at DESC
                                                        ''',params={"id": self.login_info['user_id']})
            self.ui.notification_listwidget.clear()
            for note in notifications:
                item_widget = NotificationItem( notification_content = note, parent = self,isYours= False)
                list_item = QtWidgets.QListWidgetItem()
                list_item.setSizeHint(item_widget.sizeHint())
                self.ui.notification_listwidget.addItem(list_item)
                self.ui.notification_listwidget.setItemWidget(list_item, item_widget)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load notifications: {e}")
        
        try:
            your_requests = self.database_process.query(sql='''
                                                            SELECT * FROM `Notifications`
                                                            WHERE sender_id = :id AND lifecycle_status NOT IN ('CLOSED')
                                                            ORDER BY created_at DESC
                                                        ''',params={"id": self.login_info['user_id']})
            self.ui.your_request_listwidget.clear()
            for note in your_requests:
                item_widget = NotificationItem(notification_content = note, parent = self,isYours= True)
                list_item = QtWidgets.QListWidgetItem()
                list_item.setSizeHint(item_widget.sizeHint())
                self.ui.your_request_listwidget.addItem(list_item)
                self.ui.your_request_listwidget.setItemWidget(list_item, item_widget)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load notifications: {e}")
        self.safe_connect(self.ui.update_first_name_btn.clicked,lambda _: self.update_user_info(update_column="first_name",update_content=self.ui.first_name_lnedit.text()))
        self.safe_connect(self.ui.update_last_name_btn.clicked,lambda _: self.update_user_info(update_column="last_name",update_content=self.ui.last_name_lnedit.text()))
        self.safe_connect(self.ui.update_password_btn.clicked,lambda _: self.update_user_info(update_column="password_hash",update_content=self.ui.confirm_password_lnedit.text()))
        self.safe_connect(self.ui.logout_btn.clicked,self.logout)
    
    @QtCore.pyqtSlot()
    def return_home(self):
        self.ui.frame_60.hide()
        self.ui.frame_58.setMaximumWidth(16777215)
        self.ui.frame_59.setMaximumWidth(16777215)
        self.ui.frame_61.setMaximumWidth(0)
        self.ui.change_password_frame.setMaximumWidth(0)
        self.ui.horizontalLayout_41.setContentsMargins(0,0,0,0)
        self.ui.horizontalLayout_41.setSpacing(0)
    
    @QtCore.pyqtSlot()
    def user_info(self):
        self.ui.frame_60.hide()
        self.ui.frame_58.setMaximumWidth(0)
        self.ui.frame_59.setMaximumWidth(0)
        self.ui.frame_61.setMaximumWidth(350)
        self.ui.change_password_frame.setMaximumWidth(0)
        self.ui.horizontalLayout_41.setContentsMargins(0,0,400,0)
        self.ui.horizontalLayout_41.setSpacing(0)

    @QtCore.pyqtSlot()
    def change_password(self,form_user_info = False):
        if not form_user_info:
            self.ui.frame_60.hide()
            self.ui.frame_58.setMaximumWidth(0)
            self.ui.frame_59.setMaximumWidth(0)
            self.ui.frame_61.setMaximumWidth(0)
            self.ui.change_password_frame.setMaximumWidth(16777215)
            self.ui.horizontalLayout_41.setContentsMargins(250,0,250,0)
            self.ui.horizontalLayout_41.setSpacing(0) 
        else:
            self.ui.frame_58.setMaximumWidth(0)
            self.ui.frame_59.setMaximumWidth(0)
            self.ui.frame_61.setMaximumWidth(350)
            self.ui.change_password_frame.setMaximumWidth(16777215)
            self.ui.horizontalLayout_41.setContentsMargins(0,0,150,0)
            self.ui.horizontalLayout_41.setSpacing(10)
    
    @QtCore.pyqtSlot()
    def update_user_info(self,update_column,update_content):
        try:
            if update_column != "password_hash":
                self.database_process.query(f''' UPDATE `Users` 
                                                SET {update_column} = :update_content 
                                                WHERE user_id = :id''',params={'update_content':update_content,
                                                                               'id' :self.login_info['user_id']})
            else:
                result = self.database_process.query(sql = '''  SELECT password_hash FROM `Users`
                                                                WHERE user_id = :id''',params = {'id':self.login_info['user_id']})
                if bcrypt.checkpw(self.ui.current_password_lnedit.text().strip().encode('utf-8'), result[0][0].encode('utf-8')):
                    if self.ui.new_password_lnedit.text() == self.ui.confirm_password_lnedit.text():
                        self.database_process.query(f''' UPDATE `Users` 
                                                SET {update_column} = :update_content 
                                                WHERE user_id = :id''',params={'update_content':bcrypt.hashpw(self.ui.confirm_password_lnedit.text().strip().encode('utf-8'), bcrypt.gensalt()).decode('utf-8'),
                                                                               'id' :self.login_info['user_id']})
                    else:
                        QtWidgets.QMessageBox.warning(self,"Wrong Password","Incorrect confirmation for the New password")
                        return
                else:
                    QtWidgets.QMessageBox.warning(self,"Wrong Password","Incorrect Current password")
                    return
            QtWidgets.QMessageBox.information(self,"Update success","Information updated successfully")
            self.logout(needconfirm=False)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to update data: {e}")
    
    @QtCore.pyqtSlot() 
    def logout(self,needconfirm = True):
        if needconfirm:
            reply = QtWidgets.QMessageBox.question(
                self,
                "Confirm Logout",
                "Are you sure you want to log out?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
            )
            if reply == QtWidgets.QMessageBox.Yes:
                self.cleanup_before_logout()
                self.logout_triggered = True
                self.close()
        else:
            self.cleanup_before_logout()
            self.logout_triggered = True
            self.close()

    def cleanup_before_logout(self):
        try:
            self.login_info = None
            self.ui.main_stacked.setCurrentIndex(0)
            self.ui.notification_listwidget.clear()
            self.ui.your_request_listwidget.clear()
            if hasattr(self, "worker") and self.worker.isRunning():
                self.worker.terminate()
            if hasattr(self, "database_process"):
                self.database_process.close()
        except Exception as e:
            pass
    
    @QtCore.pyqtSlot()
    def OEE_page(self):
        self.ui.main_stacked.setCurrentWidget(self.ui.OEE_page)
        self.set_stylesheet_change_page((self.ui.OEE_btn,self.ui.Home_btn,self.ui.Maintenance_btn,self.ui.Order_btn, self.ui.Stock_btn,self.ui.Downtime_btn))
        if self.is_expanded:
            self.is_expanded = False
            self.expand_windown_animation(self.is_expanded)
        self.list_df_coil_result = None
        try:
            self.machine_info = self.database_process.query(sql='''SELECT `Production_Lines`.line_name,`Machines`.machine_name,`Machine_CycleTime`.machine_id,`Machine_CycleTime`.cycletime 
                                        FROM `Machine_CycleTime`
                                        JOIN `Machines` ON `Machines`.machine_id = `Machine_CycleTime`.machine_id
                                        JOIN `Production_Lines` ON `Production_Lines`.line_id = `Machines`.line_id;''', params=None)
            self.machine_info = pd.DataFrame(self.machine_info)
        except Exception as e:
            if hasattr(self, 'database_process'):
                self.database_process.close()
            QtWidgets.QMessageBox.critical(self, "Error", str(e))
        try:
            self.model = OEE_result(self.machine_info)
            self.model.default_setting()
            self.params = {
                "NG_Coil_Sheetname": self.model.default_NG_Coil_Sheetname,
                "date_col_NG_Coil": self.model.default_date_col_NG_Coil,
                "begin_NG_coil": self.model.default_begin_NG_coil,
                "end_NG_coil": self.model.default_end_NG_coil,
                "NG_Molding_Sheetname": self.model.default_NG_Molding_Sheetname,
                "date_col_NG_Molding": self.model.default_date_col_NG_Molding,
                "begin_NG_Molding": self.model.default_begin_NG_Molding,
                "end_NG_Molding": self.model.default_end_NG_Molding,
                "month": self.model.default_month,
                "year": self.model.default_year,
                "FG_sheet_name": self.model.default_FG_sheet_name,
                "FG_date_col": self.model.default_FG_date_col,
                "FG_line_col": self.model.default_FG_line_col,
                "Molding_lt_sheet_name": self.model.default_Molding_lt_sheet_name,
                "Coil_lt_sheet_name": self.model.default_Coil_lt_sheet_name,
                "lt_date_col": self.model.default_lt_date_col
            }
        except Exception as e:
            pass
    
    @QtCore.pyqtSlot()
    def Maintenance_page(self):
        self.spinner.start()
        self.ui.main_stacked.setCurrentWidget(self.ui.Maintenance_page)
        self.scan_QRcode = Scan_record_process()
        self.set_stylesheet_change_page((self.ui.Maintenance_btn,self.ui.OEE_btn,self.ui.Home_btn,self.ui.Order_btn, self.ui.Stock_btn,self.ui.Downtime_btn))
        if not self.is_expanded:
            self.is_expanded = True
            self.expand_windown_animation(self.is_expanded)
        self.draw_circle(self.ui.upcoming_label, 85, 15, 90, (0, 255, 0))
        self.draw_circle(self.ui.neardue_label, 85, 15, 90, (196, 41, 0))
        self.Mainten_Home_page()
        QtCore.QTimer.singleShot(1000, self.spinner.stop)
    
    @QtCore.pyqtSlot()
    def Mainten_Home_page(self):
        self.style_button_with_shadow((self.ui.Main_Home_btn,self.ui.Main_detail_plan_btn,self.ui.Main_Input_record_btn,self.ui.Main_Print_record_btn))
        self.ui.Maintenance_stacked.setCurrentWidget(self.ui.Home_page_M)
        def job():
            result = self.database_process.query( sql='''
                SELECT machine_code,machine_name,department_name,line_name, working_week,status 
                FROM maintenance_with_status
                ORDER BY working_week ASC
            ''')
            return result

        self.worker = WorkerThread(job)
        self.worker.finished.connect(lambda result: self.on_home_page_data_ready(result))
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.start()

    @QtCore.pyqtSlot()
    def on_home_page_data_ready(self, result):
        try:
            headers = ["Code","Name","Group","Line", "Working\nWeek", "Status","Action"]
            self.data_model = QtGui.QStandardItemModel()
            self.data_model.setHorizontalHeaderLabels(headers)
            self.add_data_to_model(result, self.ui.Maintenance_table, self.data_model)

            delegate = StatusColorDelegate(self.ui.Maintenance_table)
            self.ui.Maintenance_table.setItemDelegate(delegate)
            delegate_btn = ButtonDelegate(buttons=("Detail","Update"))
            self.ui.Maintenance_table.setItemDelegateForColumn(6, delegate_btn)
            self.safe_connect(delegate_btn.ButtonClicked, lambda name, idx : self.on_delegate_btn_clicked(name, idx))
            self.ui.Maintenance_table.setMouseTracking(True)
            self.ui.Maintenance_table.viewport().setMouseTracking(True)
            self.ui.Maintenance_table.setSortingEnabled(True)
            self.ui.Maintenance_table.setColumnWidth(1,230)
            for i in range(2,7):
                self.ui.Maintenance_table.setColumnWidth(i,80)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
        self.monitor_week_page()

    def add_data_to_model(self,data,target,model):
        model.removeRows(0, model.rowCount())
        for row in data:
            items = []
            for col in row:
                item = QtGui.QStandardItem(str(col) if col is not None else "")
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                items.append(item)
            model.appendRow(items)
        target.setModel(model)
        self.count_equipment()

    @QtCore.pyqtSlot()
    def on_delegate_btn_clicked(self,name,index):
        model = index.model()
        row = index.row()
        code = model.data(model.index(row, 0))
        dep = model.data(model.index(row,2))
        if name == "Detail":
            self.detail_machine_information = Machine_information(database=self.database_process,code = code)
            self.detail_machine_information.show()
        else:
            if self.login_info["role_level"] in ["Manager","Admin"]:
                pass
            elif ( self.login_info["department"] == dep ) and ( self.login_info["role_level"] in ["Supervisor"]):
                pass
            else:
                QtWidgets.QMessageBox.information(self,"Permission denied","Your don't have permission to update this machine info")
                return
            self.update_info_dialog = Update_machine_info(parent= self, code = code)
            self.update_info_dialog.show()
    
    @QtCore.pyqtSlot()
    def show_filter(self):
        if self.ui.line_cbb.count() == 0:
            query = '''
                    SELECT DISTINCT line_name
                    FROM maintenance_with_status;
                    '''
            result = self.database_process.query(sql=query)
            self.ui.line_cbb.addItems([""] + [line[0] for line in result])
            self.ui.group_cbb.addItems([""] + [item[0] for item in self.group])
            self.ui.status_cbb.addItems(["","Upcoming","Near due","Overdue","No schedule"])
            self.safe_connect(self.ui.apply_btn.clicked, self.filter_process)
            self.safe_connect(self.ui.cancel_btn.clicked,self.hide_filter)
            self.safe_connect(self.ui.code_lnedit.textChanged,lambda : self.filter_suggestion(self.ui.code_lnedit,"machine_code","maintenance_with_status"))
            self.safe_connect(self.ui.name_lnedit.textChanged, lambda : self.filter_suggestion(self.ui.name_lnedit,"machine_name","maintenance_with_status"))
            self.safe_connect(self.ui.group_cbb.currentTextChanged, self.group_cbb_Home_Maintenance_change)
        self.ui.filter_mainten_frame.show()
    
    @QtCore.pyqtSlot()
    def hide_filter(self):
        self.ui.filter_mainten_frame.hide()
    
    @QtCore.pyqtSlot()
    def filter_suggestion(self,target,text,table,where = None):
        if len(target.text())<2:
            return
        suggestions = []
        machine_code = []
        script = f'''SELECT {text} FROM {table}'''
        if where != None:
            script = script + where + "LIMIT 10;"
        else:
            script = script + f" WHERE {text} LIKE '%{target.text()}%' " + "LIMIT 10;"
        try:
            machine_code = self.database_process.query(sql=script)
            suggestions = [str(name[0]) if len(name) == 1 else f"{name[0]} : {name[1]}" for name in machine_code ]
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to fetch machine names: {e}")
            suggestions = []
        if suggestions:
            self.completer = QtWidgets.QCompleter(suggestions, self)
            self.completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
            self.completer.popup().setAlternatingRowColors(True)
            target.setCompleter(self.completer)
    
    @QtCore.pyqtSlot()
    def filter_process(self):
        try:
            query = []
            if self.ui.code_lnedit.text() != "":
                query.append(f'machine_code LIKE "%{self.ui.code_lnedit.text()}%"')
            if self.ui.name_lnedit.text() != "":
                query.append(f'machine_name LIKE "%{self.ui.name_lnedit.text()}%"')
            if self.ui.group_cbb.currentText() != "":
                query.append(f'department_name = "{self.ui.group_cbb.currentText()}"')
            if self.ui.line_cbb.currentText() != "":
                query.append(f'line_name = "{self.ui.line_cbb.currentText()}"')
            if self.ui.status_cbb.currentText() != "":
                query.append(f'status COLLATE utf8mb4_unicode_ci = "{self.ui.status_cbb.currentText()}"')
            query = " AND ".join(query)
            if query == "":
                result = self.database_process.query(sql='''SELECT machine_code,machine_name,department_name,line_name,working_week,status 
                                                            FROM maintenance_with_status
                                                            ORDER BY next_due_date ASC''')
                self.add_data_to_model(result,self.ui.Maintenance_table,self.data_model)
                self.hide_filter() 
                return
            final_query = f'''SELECT machine_code,machine_name,department_name,line_name,working_week,status 
                                                                FROM maintenance_with_status
                                                                WHERE {query}  ORDER BY next_due_date ASC'''
            result = self.database_process.query(sql=final_query)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to filter data: {e}")
            return
        self.add_data_to_model(result,self.ui.Maintenance_table,self.data_model)
        self.hide_filter()
    
    @QtCore.pyqtSlot()
    def reset_filter(self):
        try:
            self.ui.code_lnedit.clear()
            self.ui.name_lnedit.clear()
            self.ui.group_cbb.setCurrentIndex(0)
            self.ui.line_cbb.clear()
            query = '''
                        SELECT DISTINCT line_name
                        FROM maintenance_with_status;
                        '''
            result = self.database_process.query(sql=query)
            self.ui.line_cbb.addItems([""] + [line[0] for line in result])
            self.ui.status_cbb.setCurrentIndex(0)
            self.filter_process()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot()
    def group_cbb_Home_Maintenance_change(self):
        dep = self.ui.group_cbb.currentText()
        try:
            self.ui.line_cbb.clear()
            line_list = self.database_process.query(sql=''' SELECT DISTINCT line_name
                                                            FROM maintenance_with_status
                                                            WHERE department_name = :dep''',params= {'dep':dep})
            self.ui.line_cbb.addItems([""] + [line[0] for line in line_list])
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    def count_equipment(self):
        try:
            upcoming = self.database_process.query(sql='''SELECT COUNT(status)
                                                        FROM maintenance_with_status
                                                        WHERE status COLLATE utf8mb4_unicode_ci = "Upcoming";''')
            overdue = self.database_process.query(sql='''SELECT COUNT(status)
                                                        FROM maintenance_with_status
                                                        WHERE status COLLATE utf8mb4_unicode_ci = "Overdue";''')
            near_due = self.database_process.query(sql='''SELECT COUNT(status)
                                                        FROM maintenance_with_status
                                                        WHERE status COLLATE utf8mb4_unicode_ci = "Near due";''')
            def set_fontsize(target,num):
                if num > 999:
                    target.setStyleSheet('font-size: 14pt')
                    return
                elif num> 99:
                    target.setStyleSheet('font-size: 20pt')
                    return
                else:
                    return
            set_fontsize(self.ui.upcoming_num,upcoming[0][0])
            set_fontsize(self.ui.neardue_num,near_due[0][0])
            self.ui.upcoming_num.setText(str(upcoming[0][0]))
            self.ui.neardue_num.setText(str(near_due[0][0]))
        except Exception as e:
             QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to count data: {e}")
    
    @QtCore.pyqtSlot()
    def monitor_week_page(self):
        self.ui.monitor_stacked.setCurrentWidget(self.ui.monitor_week_page)
        self.style_button_with_shadow((self.ui.weekly_btn,self.ui.monthly_btn,self.ui.inyear_btn))
        def job():
            Total = self.database_process.query(
                ''' SELECT COUNT(DISTINCT mp.line_id) AS total
                    FROM `Maintenance_plan` mp
                    JOIN `Production_Lines` p ON p.line_id = mp.line_id
                    JOIN `Departments` d ON p.department_id = d.department_id
                    JOIN `Months_Years` as my ON mp.month_year_id = my.month_year_id
                    WHERE week = :week AND d.department_id < 7 AND my.year = :year AND (mp.status IN ('Ontime','Overdue') OR mp.status IS NULL);''',
                params={'week': self.week_num,'year':self.year_num}
            )
            sql = '''SELECT p.line_name, COUNT(p.line_name) AS plan_count
                    FROM Maintenance_plan mp
                    JOIN Production_Lines p ON p.line_id = mp.line_id
                    JOIN Departments d ON p.department_id = d.department_id
                    JOIN `Months_Years` as my ON mp.month_year_id = my.month_year_id
                    WHERE d.department_name = :dept AND mp.week = :week AND my.year = :year AND (mp.status IN ('Ontime','Overdue') OR mp.status IS NULL)
                    GROUP BY p.line_name;'''
            PE1 = self.database_process.query(sql=sql, params={'week': self.week_num, 'dept': "PE1",'year':self.year_num})
            PE2 = self.database_process.query(sql=sql, params={'week': self.week_num, 'dept': "PE2",'year':self.year_num})
            PE3 = self.database_process.query(sql=sql, params={'week': self.week_num, 'dept': "PE3",'year':self.year_num})
            PE4 = self.database_process.query(sql=sql, params={'week': self.week_num, 'dept': "PE4",'year':self.year_num})
            PE5 = self.database_process.query(sql=sql, params={'week': self.week_num, 'dept': "PE5",'year':self.year_num})
            sql2 = '''SELECT 
                            COALESCE(
                                COUNT(DISTINCT CASE WHEN mp.status IN ('Ontime','Overdue') OR mp.status IS NULL THEN mp.line_id END) 
                                - COUNT(DISTINCT CASE WHEN mp.status IS NULL OR mp.status = 'Overdue' THEN mp.line_id END), 0
                            ) AS complete,
                            COALESCE(
                                COUNT(CASE WHEN mp.status IN ('Ontime','Overdue') OR mp.status IS NULL THEN mp.machine_id END) 
                                - COUNT(CASE WHEN mp.status IS NULL OR mp.status = 'Overdue' THEN mp.machine_id END), 0
                            ) AS complete_mc
                        FROM (
                            SELECT 'PE1' AS department_name
                            UNION ALL SELECT 'PE2'
                            UNION ALL SELECT 'PE3'
                            UNION ALL SELECT 'PE4'
                            UNION ALL SELECT 'PE5'
                        ) AS d
                        LEFT JOIN `Departments` as dep ON dep.department_name = d.department_name
                        LEFT JOIN `Production_Lines` as p ON p.department_id = dep.department_id
                        LEFT JOIN `Maintenance_plan` as mp ON mp.line_id = p.line_id AND mp.week = :week
                        LEFT JOIN `Months_Years` as my ON mp.month_year_id = my.month_year_id AND my.year = :year
                        GROUP BY d.department_name
                        ORDER BY d.department_name;'''
            result = self.database_process.query(sql=sql2, params={'week': self.week_num,'year':self.year_num})

            return {"Total": Total, "PE1": PE1, "PE2": PE2, "PE3": PE3,
                    "PE4": PE4, "PE5": PE5, "Result": result}
        self.worker = WorkerThread(job)
        self.worker.finished.connect(lambda data: self.on_monitor_data_ready(data=data))
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.start()
    
    @QtCore.pyqtSlot()
    def on_monitor_data_ready(self, data):
        try:
            Total, PE1, PE2, PE3, PE4, PE5, result = (
                data["Total"], data["PE1"], data["PE2"],
                data["PE3"], data["PE4"], data["PE5"], data["Result"]
            )

            headers = ["Item","Content"]
            monitor_model = QtGui.QStandardItemModel()
            monitor_model.setHorizontalHeaderLabels(headers)

            PE1_line = []
            PE2_line =[]
            PE3_line =[]
            PE4_line =[]
            PE5_line =[]
            machine_qty = [0,0,0,0,0]
            def insert_into_item(data:list,target:list,group:str):
                if data and data[0][0] is not None:
                    target.extend([(row[0],) for row in data])
                    self.insert_item(target, group, monitor_model)
                else:
                    self.insert_item("", group, monitor_model)
            def total_machine(data:list,target:list,index:int):
                machine_qty[index] = sum([data[i][1] for i in range(len(data))])
            self.insert_item([(str(Total[0][0]),)],"Total",monitor_model)
            insert_into_item(PE1,PE1_line,"PE1")
            insert_into_item(PE2,PE2_line,"PE2")
            insert_into_item(PE3,PE3_line,"PE3")
            insert_into_item(PE4,PE4_line,"PE4")
            insert_into_item(PE5,PE4_line,"PE5")
            total_machine(PE1,machine_qty,0)
            total_machine(PE2,machine_qty,1)
            total_machine(PE3,machine_qty,2)
            total_machine(PE4,machine_qty,3)
            total_machine(PE5,machine_qty,4)

            self.ui.week_plan_table_line.setModel(monitor_model)
            self.ui.week_plan_table_line.setColumnWidth(0, 10)
            self.ui.week_plan_table_line.resizeRowsToContents()

            self.draw_monitor_chart(target= self.ui.week_plan_chart_line ,lines = ["PE1", "PE2", "PE3", "PE4", "PE5"],
                                plan = [len(PE1), len(PE2), len(PE3), len(PE4), len(PE5)], 
                                result = [result[0][0], result[1][0], result[2][0], result[3][0], result[4][0]])
            self.draw_monitor_chart(target= self.ui.week_plan_chart_mc ,lines = ["PE1", "PE2", "PE3", "PE4", "PE5"],
                                plan = machine_qty, 
                                result = [result[0][1], result[1][1], result[2][1], result[3][1], result[4][1]],set_title= True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")

    def insert_item(self,data,name,model):
        data = set(data)     
        str_line =[]
        for item in data:
            str_line.append(item[0])
        str_line = ", ".join(str_line)
        row = [ QtGui.QStandardItem(name),QtGui.QStandardItem(str_line) ]
        model.appendRow(row)

    def draw_monitor_chart(self,target,lines, plan, result,set_title=False):
        layout = target.layout()
        if layout is None:
            layout = QtWidgets.QVBoxLayout(target)
            target.setLayout(layout)
        if not hasattr(target, "canvas"):
            fig, ax = plt.subplots(figsize=(5, 3))
            target.canvas = FigureCanvas(fig)
            target.canvas.setFixedSize(target.width()-5,target.height()-10)
            target.ax = ax
            target.fig = fig
            layout.addWidget(target.canvas)
        else:
            target.ax.clear()
        ax = target.ax
        fig = target.fig
        x = range(len(lines))
        fig.patch.set_alpha(0.0)
        ax.set_facecolor("none")
        ax.tick_params(axis="y", length=0)
        ax.tick_params(axis="x", length=0)
        plan_col = ax.bar([i - 0.2 for i in x], plan, width=0.4, label="Plan", color=(18/255, 184/255, 234/255, 1))
        result_col = ax.bar([i + 0.2 for i in x], result, width=0.4, label="Result", color =(63/255, 218/255, 155/255, 1))
        max_val = max(max(plan), max(result))
        margin = max_val * 0.8 
        ax.set_ylim(0, max_val + margin)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)
        for bar in plan_col:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, height + 0.1,
                    f"{height}", ha='center', va='bottom', fontsize=9)
        for bar in result_col:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, height + 0.1,
                    f"{height}", ha='center', va='bottom', fontsize=9)
        if set_title == True:
            ax.text(-0.1, 1.05, 
                    f"Plan: {sum(plan)}",
                    transform=ax.transAxes,
                    fontsize=9,
                    fontweight="bold",
                    color=(18/255, 184/255, 234/255, 1),
                    ha="left", va="bottom")     
            ax.text(
                -0.1, 0.95,  
                f"Result: {sum(result)}",
                transform=ax.transAxes,
                fontsize=9,
                fontweight="bold",
                color=(63/255, 218/255, 155/255, 1),
                ha="left", va="bottom"
            )  
        ax.set_xticks(x)
        ax.set_xticklabels(lines)
        ax.legend(loc="upper right")
        layout.addWidget(target.canvas)
        target.canvas.draw()
        plt.close(fig)   
    
    @QtCore.pyqtSlot()
    def monitor_month_page(self):
        self.ui.monitor_stacked.setCurrentWidget(self.ui.monitor_month_page)
        self.style_button_with_shadow((self.ui.monthly_btn,self.ui.weekly_btn,self.ui.inyear_btn))
        try: 
            Total = self.database_process.query(sql='''SELECT COUNT(DISTINCT mp.line_id) AS total
                                                        FROM `Maintenance_plan` as mp
                                                        JOIN `Production_Lines` as p
                                                        ON p.line_id = mp.line_id
                                                        JOIN `Departments` as d
                                                        ON p.department_id = d.department_id
                                                        JOIN `Months_Years` as my
                                                        ON mp.month_year_id = my.month_year_id
                                                        WHERE my.month = :month AND my.year = :year AND d.department_id < 7 AND ( mp.status IN ('Ontime','Overdue') OR mp.status IS NULL);''',
                                                        params={'month': self.month_num , 'year':self.year_num })
            sql = '''SELECT p.line_name, COUNT( p.line_name) AS plan_count
                                                    FROM Maintenance_plan mp 
                                                    JOIN Production_Lines p
                                                    ON p.line_id = mp.line_id
                                                    JOIN Departments d 
                                                    ON p.department_id = d.department_id
                                                    JOIN `Months_Years` as my
                                                    ON mp.month_year_id = my.month_year_id
                                                    WHERE d.department_name = :dept
                                                    AND my.month = :month AND my.year = :year AND (mp.status IN ('Ontime','Overdue') OR mp.status IS NULL)
                                                    GROUP BY p.line_name; '''                
            PE1 = self.database_process.query(sql=sql,params={'dept':"PE1",'month': self.month_num , 'year':self.year_num})
            PE2 = self.database_process.query(sql=sql,params={'dept':"PE2",'month': self.month_num , 'year':self.year_num})
            PE3 = self.database_process.query(sql=sql,params={'dept':"PE3",'month': self.month_num , 'year':self.year_num})
            PE4 = self.database_process.query(sql=sql,params={'dept':"PE4",'month': self.month_num , 'year':self.year_num})
            PE5 = self.database_process.query(sql=sql,params={'dept':"PE5",'month': self.month_num , 'year':self.year_num})
            sql = '''SELECT 
                        COALESCE(
                            COUNT(DISTINCT CASE WHEN mp.status IN ('Ontime','Overdue') OR mp.status IS NULL THEN mp.line_id END) 
                            - COUNT(DISTINCT CASE WHEN mp.status IS NULL OR mp.status IN ('Overdue') THEN mp.line_id END), 0
                        ) AS complete,
                        COALESCE(
                            COUNT(CASE WHEN mp.status IN ('Ontime','Overdue') OR mp.status IS NULL THEN mp.machine_id END) 
                            - COUNT(CASE WHEN mp.status IS NULL OR mp.status IN ('Overdue') THEN mp.machine_id END), 0
                        ) AS complete_mc
                    FROM (
                        SELECT 'PE1' AS department_name
                        UNION ALL SELECT 'PE2'
                        UNION ALL SELECT 'PE3'
                        UNION ALL SELECT 'PE4'
                        UNION ALL SELECT 'PE5'
                    ) AS d
                    LEFT JOIN Departments dep ON dep.department_name = d.department_name
                    LEFT JOIN Production_Lines p ON p.department_id = dep.department_id
                    LEFT JOIN Maintenance_plan mp
                    ON mp.line_id = p.line_id
                    AND mp.month_year_id = (
                        SELECT month_year_id
                        FROM Months_Years
                        WHERE year = :year AND month = :month
                        LIMIT 1
                    )
                    GROUP BY d.department_name
                    ORDER BY d.department_name;'''
            result = self.database_process.query(sql=sql, params={'month':self.month_num,'year':self.year_num})
            headers = ["Item","Content"]
            monitor_model = QtGui.QStandardItemModel()
            monitor_model.setHorizontalHeaderLabels(headers)
            monitor_model.removeRows(0, monitor_model.rowCount())
            PE1_line = []
            PE2_line =[]
            PE3_line =[]
            PE4_line =[]
            PE5_line =[]
            machine_qty = [0,0,0,0,0]
            def insert_into_item(data:list,target:list,group:str):
                if data and data[0][0] is not None:
                    target.extend([(row[0],) for row in data])
                    self.insert_item(target, group, monitor_model)
                else:
                    self.insert_item("", group, monitor_model)
            def total_machine(data:list,target:list,index:int):
                machine_qty[index] = sum([data[i][1] for i in range(len(data))])

            self.insert_item([(str(Total[0][0]),)],"Total",monitor_model)
            insert_into_item(PE1,PE1_line,"PE1")
            insert_into_item(PE2,PE2_line,"PE2")
            insert_into_item(PE3,PE3_line,"PE3")
            insert_into_item(PE4,PE4_line,"PE4")
            insert_into_item(PE5,PE5_line,"PE5")
            total_machine(PE1,machine_qty,0)
            total_machine(PE2,machine_qty,1)
            total_machine(PE3,machine_qty,2)
            total_machine(PE4,machine_qty,3)
            total_machine(PE5,machine_qty,4)
            self.ui.month_plan_table_line.setModel(monitor_model)
            self.ui.month_plan_table_line.setColumnWidth(0,10)
            self.ui.month_plan_table_line.resizeRowsToContents()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")
        self.draw_monitor_chart(target= self.ui.month_plan_chart_line ,lines = ["PE1", "PE2", "PE3", "PE4", "PE5"],
                             plan = [len(PE1),len(PE2),len(PE3),len(PE4),len(PE5)], 
                             result =  [result[0][0], result[1][0], result[2][0], result[3][0], result[4][0]]) 
        self.draw_monitor_chart(target= self.ui.month_plan_chart_mc ,lines = ["PE1", "PE2", "PE3", "PE4", "PE5"],
                             plan = machine_qty, 
                             result =  [result[0][1], result[1][1], result[2][1], result[3][1], result[4][1]],set_title= True)
    
    @QtCore.pyqtSlot()
    def monitor_inyear_page(self):
        def safe_divide(a, b):
            if b == 0:
                return 0  
            return a / b
        self.ui.monitor_stacked.setCurrentWidget(self.ui.monitor_year_page)
        self.style_button_with_shadow((self.ui.inyear_btn,self.ui.weekly_btn,self.ui.monthly_btn))
        headers = ["","Total","Overdue"]
        model = QtGui.QStandardItemModel()
        model.setHorizontalHeaderLabels(headers)
        self.ui.KPI_table.setModel(model)
        self.ui.KPI_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        total_sql ='''SELECT
                        (SELECT "Total" ) as department_name,
                        COUNT(CASE WHEN mp.status IN ('Ontime','Overdue') THEN 1 END) AS total,
                        COUNT(CASE WHEN mp.status = "Overdue" THEN 1 END) AS overdue
                        FROM Maintenance_plan mp
                        JOIN `Months_Years` as my ON my.month_year_id = mp.month_year_id
                        WHERE my.year = :year;'''
        dep_sql ='''SELECT
                d.department_name,
                COUNT(CASE WHEN mp.status IN ('Ontime','Overdue') THEN 1 END) AS total,
                COUNT(CASE WHEN mp.status ='Overdue' THEN 1 END) AS overdue
            FROM (
                SELECT 'PE1' AS department_name
                UNION ALL SELECT 'PE2'
                UNION ALL SELECT 'PE3'
                UNION ALL SELECT 'PE4'
                UNION ALL SELECT 'PE5'
            ) AS d
            LEFT JOIN `Departments` as dep ON dep.department_name = d.department_name
            LEFT JOIN `Production_Lines` as p ON p.department_id = dep.department_id
            LEFT JOIN `Maintenance_plan` as mp ON mp.line_id = p.line_id
                AND mp.month_year_id IN (   SELECT month_year_id
                                            FROM Months_Years
                                            WHERE year = :year)
            GROUP BY d.department_name;'''
        try: 
            result = self.database_process.query(sql=total_sql,params= {'year':self.year_num})
            result += self.database_process.query(sql = dep_sql,params= {'year':self.year_num})
            for row in result:
                items = []
                for col in row:
                    item = QtGui.QStandardItem(str(col) if col is not None else "")
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    items.append(item)
                model.appendRow(items)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")
        self.ui.KPI_table.setAlternatingRowColors(True)
        self.create_pie_chart(self.ui.total_kpi,result[0][1]-result[0][2],result[0][2])
        self.ui.PE1_KPI.setValue(int((1-safe_divide(result[1][2],result[1][1]))*100))
        self.ui.PE2_KPI.setValue(int((1-safe_divide(result[2][2],result[2][1]))*100))
        self.ui.PE3_KPI.setValue(int((1-safe_divide(result[3][2],result[3][1]))*100))
        self.ui.PE4_KPI.setValue(int((1-safe_divide(result[4][2],result[4][1]))*100))
        self.ui.PE5_KPI.setValue(int((1-safe_divide(result[5][2],result[5][1]))*100))
    
    def create_pie_chart(self,target,plan,result,fontdict={'fontsize': 14,
                                                            'fontweight': 'bold',       
                                                            'color': '#008b8b',         
                                                            'fontname': "Comic Sans MS" 
                                                        }):
        def to_number(x):
            try:
                return float(x) if x is not None else 0.0
            except Exception:
                return 0.0

        plan_val = max(0.0, to_number(plan))
        result_val = max(0.0, to_number(result))
        total = plan_val + result_val

        layout = target.layout()
        if layout is None:
            layout = QtWidgets.QVBoxLayout(target)
            layout.setContentsMargins(0, 0, 0, 0)
            target.setLayout(layout)
        if not hasattr(target, "canvas"):
            fig, ax = plt.subplots(figsize=(5, 3))
            fig.patch.set_alpha(0.0)
            ax.set_facecolor("none")
            target.canvas = FigureCanvas(fig)
            target.canvas.setFixedSize(target.width(), target.height())
            target.ax = ax
            target.fig = fig
            layout.addWidget(target.canvas)
        else:
            target.ax.clear()

        ax = target.ax
        fig = target.fig

        # Nếu không có dữ liệu, hiển thị thông báo thay vì vẽ pie để tránh NaN
        if total <= 0:
            ax.set_title("KPI chart", fontdict=fontdict)
            ax.axis("off")
            ax.text(0.5, 0.5, "No data", ha="center", va="center",
                    fontsize=12, color="#888", transform=ax.transAxes)
            target.canvas.draw()
            return

        sizes = [plan_val, result_val]
        labels = ["Ontime", "Overdue"]
        colors = ["#008b8b", "#CC9B9BFC"]

        def autopct_format(pct):
            return ('%1.1f%%' % pct) if pct > 5 else ''

        wedges, texts, autotexts = ax.pie(
            sizes,
            colors=colors,
            autopct=autopct_format,
            startangle=90
        )
        for autotext in autotexts:
            autotext.set_color("white")
            autotext.set_fontfamily("Comic Sans MS")
            autotext.set_fontsize(12)
            autotext.set_weight("bold")

        ax.set_title("KPI chart", fontdict=fontdict)
        ax.legend(loc="upper right", bbox_to_anchor=(0.25, 1),
                  fontsize=8, labels=labels)
        target.canvas.draw()
    
    @QtCore.pyqtSlot()
    def next_monitor_page(self):
        self.ui.monitor_next_btn.setEnabled(False)
        QtCore.QTimer.singleShot(300,lambda:self.ui.monitor_next_btn.setEnabled(True))
        if (self.ui.monitor_stacked.currentWidget() is self.ui.monitor_week_page):
            self.week_num = self.week_num + 1
            if self.week_num > self.qty_week:
                self.week_num = 1
            self.ui.weekly_btn.setText(f"Week: {self.week_num}")
            self.monitor_week_page()
        elif (self.ui.monitor_stacked.currentWidget() is self.ui.monitor_month_page):
            self.month_num = self.month_num + 1
            if self.month_num > 12:
                self.month_num = 1
            self.ui.monthly_btn.setText(f"Month: {self.month_num}")
            self.monitor_month_page()
    
    @QtCore.pyqtSlot()
    def back_monitor_page(self):
        self.ui.monitor_back_btn.setEnabled(False)
        QtCore.QTimer.singleShot(300,lambda:self.ui.monitor_back_btn.setEnabled(True))
        if (self.ui.monitor_stacked.currentWidget() is self.ui.monitor_week_page):
            self.week_num = self.week_num - 1
            if (self.week_num < 1):
                self.week_num = self.ui.company_week_number(dt.date(self.year_num,12,31))
            self.ui.weekly_btn.setText(f"Week: {self.week_num}")
            self.monitor_week_page()
        elif (self.ui.monitor_stacked.currentWidget() is self.ui.monitor_month_page):
            self.month_num = self.month_num - 1
            if self.month_num < 1:
                self.month_num = 12
            self.ui.monthly_btn.setText(f"Month: {self.month_num}")
            self.monitor_month_page()
    
    @QtCore.pyqtSlot()
    def Mainten_Print_page(self):
        self.style_button_with_shadow((self.ui.Main_Print_record_btn,self.ui.Main_detail_plan_btn,self.ui.Main_Home_btn,self.ui.Main_Input_record_btn))
        self.ui.Maintenance_stacked.setCurrentWidget(self.ui.Print_page_M)
        try:
            if self.ui.Group_cbb_PF.count() == 0:
                self.result_print_record = []
                self.ui.Group_cbb_PF.addItems([group[0] for group in self.group])
                self.ui.Group_cbb_PF.setCurrentText(self.login_info['department'])
                self.department_print_record = self.ui.Group_cbb_PF.currentText()
                if self.login_info['role_level'] == 'Admin':
                    self.ui.Group_cbb_PF.setEnabled(True)
                else:
                    self.ui.Group_cbb_PF.setEnabled(False)
                current_week = self.ui.company_week_number(self.ui.today)
                week_lst = [str(week_num) for week_num in range(1,self.qty_week+1)]
                self.ui.FromWeek_cbb_PF.addItems(week_lst)
                self.ui.FromWeek_cbb_PF.setCurrentIndex(current_week - 1)
                self.ui.ToWeek_cbb_PF.addItems(week_lst)
                self.ui.ToWeek_cbb_PF.setCurrentIndex(current_week - 1)
                header = ["Machine code", "Machine Name","Attached\nequipment" ,"Group","Line","Last maintenance\n date","Week plan","Issued Maintenance\n date","Technical","Form type"]
                self.ui.print_record_table.setColumnCount(len(header))
                self.ui.print_record_table.setHorizontalHeaderLabels(header)
                self.ui.print_record_table.setColumnWidth(0,100)
                self.ui.print_record_table.setColumnWidth(1,300)
                self.ui.print_record_table.setColumnWidth(2,100)
                self.ui.print_record_table.setColumnWidth(3,100)
                self.ui.print_record_table.setColumnWidth(4,100)
                self.ui.print_record_table.setColumnWidth(5,100)
                self.ui.print_record_table.setColumnWidth(6,100)
                self.ui.print_record_table.setColumnWidth(7,120)
                self.ui.print_record_table.setColumnWidth(8,100)
                self.ui.print_record_table.setColumnWidth(9,300)
                self.ui.print_record_table.setSortingEnabled(True)
                self.add_item_line_PF()
                self.safe_connect(self.ui.FromWeek_cbb_PF.currentIndexChanged, self.check_cbb)
                self.safe_connect(self.ui.ToWeek_cbb_PF.currentIndexChanged, self.check_cbb)
                self.safe_connect(self.ui.Load_btn.clicked, lambda _: self.load_record_form()) 
                self.safe_connect(self.ui.insert_row_btn_PF.clicked, lambda _: self.insert_row(target = self.ui.print_record_table))
                self.safe_connect(self.ui.del_row_btn_PF.clicked, lambda _: self.delete_row(target = self.ui.print_record_table ,save_list= self.result_print_record))
                self.safe_connect(self.ui.print_btn_PF.clicked, lambda _: self.print_record())
                self.safe_connect(self.ui.Update_form_btn.clicked, lambda _: self.update_register_form())
                self.safe_connect(self.ui.Register_form_btn.clicked, lambda _: self.register_new_form())
                self.safe_connect(self.ui.Clear_btn.clicked, lambda _: self.clear_print_data())
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Fail to load data: {e}")      
    
    @QtCore.pyqtSlot()
    def add_item_line_PF(self):
        try: 
            lines = self.database_process.query(sql='''SELECT p.line_name
                                                        FROM `Production_Lines` as p 
                                                        JOIN `Departments` as d
                                                        ON p.department_id = d.department_id
                                                        WHERE d.department_name = :dep
                                                        ORDER BY p.line_name ASC''',params={'dep': self.ui.Group_cbb_PF.currentText()})
            items = ["All"] + [line[0] for line in lines]
            self.ui.Line_cbb_PF.clear()
            self.ui.Line_cbb_PF.addItems(items)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
    
    @QtCore.pyqtSlot()
    def check_cbb(self):
        if self.ui.ToWeek_cbb_PF.currentIndex() >= self.ui.FromWeek_cbb_PF.currentIndex():
            return   
        QtWidgets.QMessageBox.warning(self, "Wrong select", f"""You have selected the wrong format."From week" must be smaller than "To Week", please select again.""")
        self.ui.ToWeek_cbb_PF.setCurrentIndex(self.ui.FromWeek_cbb_PF.currentIndex()) 

    def add_item_into_print_record(self,data:list,c = None,r = None,is_readOnly = None,widget = None):
        if widget == None:
            editor = QtWidgets.QLineEdit()
            if len(data) == 1:
                r_data = 0
            else:
                r_data = r
            if isinstance(data[r_data][c],str) == True: 
                editor.setText(data[r_data][c])
            else:
                editor.setText(str(data[r_data][c]))
            editor.setFrame(False)
            editor.setAlignment(QtCore.Qt.AlignCenter) 
            if not is_readOnly:
                editor.setStyleSheet("background-color: rgba(0,0,0,0.05);")
            else:
                self.safe_connect(editor.textChanged, self.handle_text_changed)
                self.safe_connect(editor.editingFinished, self.handle_editing_finished)
            editor.setEnabled(is_readOnly)
            self.ui.print_record_table.setCellWidget(r, c, editor)
        else:
            widget.setText(str(data[0][c]))
    
    @QtCore.pyqtSlot(str)
    def handle_text_changed(self, text):
        editor = self.sender()
        if editor is None:
            return
        table = self.ui.print_record_table
        for r in range(table.rowCount()):
            for c in range(table.columnCount()):
                if table.cellWidget(r, c) is editor:
                    return self.on_text_changed(
                        text=text, c=c, r=r, target=table,
                        select_col="machine_code", table="`Machines` as m ",
                        where=f"WHERE m.machine_code LIKE '%{text}%'"
                )
    
    @QtCore.pyqtSlot()
    def handle_editing_finished(self):
        editor = self.sender()
        if editor is None:
            return
        table = self.ui.print_record_table
        for r in range(table.rowCount()):
            for c in range(table.columnCount()):
                if table.cellWidget(r, c) is editor:
                    return self.reload_data(r, c)
    
    @QtCore.pyqtSlot()
    def load_record_form(self):
        self.ui.print_record_table.clearContents()
        filter_script = [f" department_name = '{self.ui.Group_cbb_PF.currentText()}'"]
        if self.ui.Line_cbb_PF.currentText() != "All":
            filter_script.append(f"line_name = '{self.ui.Line_cbb_PF.currentText()}'")
        filter_script = filter_script + [f"week >= {self.ui.FromWeek_cbb_PF.currentText()}",f"week <= {self.ui.ToWeek_cbb_PF.currentText()}"]
        filter_script = " AND ".join(filter_script)
        try: 
            self.result_print_record = self.database_process.query(sql=f'''SELECT machine_code,machine_name,NULL,department_name,line_name,last_maintenance_date, week,(SELECT CURRENT_DATE),technician,form_name,form_link FROM  maintenance_form_info
                                                            WHERE {filter_script}''')
            if len(self.result_print_record) == 0:
                raise ValueError("Don't see machine in maintenance plan")
            self.ui.print_record_table.setRowCount(len(self.result_print_record))
            for row in range(len(self.result_print_record)):
                self.add_item_into_print_record(self.result_print_record,0,row,True)
                self.add_item_into_print_record(self.result_print_record,1,row,False)
                self.add_item_into_print_record(self.result_print_record,3,row,False)
                self.add_item_into_print_record(self.result_print_record,4,row,True)
                self.add_item_into_print_record(self.result_print_record,5,row,False)
                self.add_item_into_print_record(self.result_print_record,6,row,False)
                self.add_item_into_print_record(self.result_print_record,7,row,True)
                self.add_item_into_print_record(self.result_print_record,8,row,True)
                self.add_item_into_print_record(self.result_print_record,9,row,False)
            delegate = DynamicSuggestion(database=self.database_process,dep=self.ui.Group_cbb_PF.currentText(),year = self.year_num)
            self.ui.print_record_table.setItemDelegateForColumn(2, delegate)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
    
    @QtCore.pyqtSlot()
    def on_text_changed(self,text,c,r,target,select_col,table,where):
        if c == 0:
            self.filter_suggestion(target.cellWidget(r,c),select_col,table,where)      
        elif c == 4:
            editor = target.cellWidget(r, 3)
            value = editor.text()
            self.filter_suggestion(target.cellWidget(r,c),"p.line_name","`Production_Lines` as p ",f'''JOIN `Departments` as d
                                                                                                        ON p.department_id = d.department_id
                                                                                                        WHERE d.department_name = "{value}"
                                                                                                        AND p.line_name LIKE "%{text}%"''')
    
    @QtCore.pyqtSlot()
    def reload_data(self,r,c):
        editor = self.ui.print_record_table.cellWidget(r, c)
        if editor is None:
            return 
        if c != 0:
            value = editor.text()
            if c == 4:
                dep_editor = self.ui.print_record_table.cellWidget(r, 3)
                dep = dep_editor.text()
                iscorrectDep = self.database_process.query(sql = ''' SELECT 1 FROM `Production_Lines` as p
                                                                    JOIN Departments as d
                                                                    ON p.department_id = d.department_id
                                                                    WHERE p.line_name = :line AND d.department_name = :dep '''
                                                                    , params= {'line':value,'dep':dep})
                if not iscorrectDep:
                    editor.clear()
                    return
            edit_row = list(self.result_print_record[r])
            edit_row[c] = value.upper()
            self.result_print_record[r] = tuple(edit_row)
            return
        editor = self.ui.print_record_table.cellWidget(r, 0)
        value = editor.text()
        try: 
            result = self.database_process.query(sql=f'''SELECT machine_code,machine_name,NULL,department_name,line_name,last_maintenance_date, week,(SELECT CURRENT_DATE),technician,form_name,form_link 
                                                            FROM  maintenance_form_info
                                                            WHERE machine_code = :code 
                                                            ORDER BY week ASC LIMIT 1;''', params={'code':value})
            if len(result) == 0:
                raise ValueError("Don't see machine in maintenance plan")
            current_week = self.ui.company_week_number(self.ui.today)
            if (result[0][6] > ( current_week + 4 )) or (result[0][6] < (current_week - 4)):
                raise ValueError("This machine is not in the maintenance plan of time")
            self.add_item_into_print_record(data = result, c = 0,is_readOnly = True,widget= self.ui.print_record_table.cellWidget(r, 0))
            self.add_item_into_print_record(data = result, c = 1,is_readOnly = False,widget= self.ui.print_record_table.cellWidget(r, 1))
            self.add_item_into_print_record(data = result, c = 3,is_readOnly = False , widget= self.ui.print_record_table.cellWidget(r, 3))
            self.add_item_into_print_record(data = result, c = 4,is_readOnly = True , widget= self.ui.print_record_table.cellWidget(r, 4))
            self.add_item_into_print_record(data = result, c = 5,is_readOnly = False, widget= self.ui.print_record_table.cellWidget(r, 5))
            self.add_item_into_print_record(data = result, c = 6,is_readOnly = False, widget= self.ui.print_record_table.cellWidget(r, 6))
            self.add_item_into_print_record(data = result, c = 7,is_readOnly = True,  widget= self.ui.print_record_table.cellWidget(r, 7))
            self.add_item_into_print_record(data = result, c = 8,is_readOnly = True,  widget= self.ui.print_record_table.cellWidget(r, 8))
            self.add_item_into_print_record(data = result, c = 9,is_readOnly = False, widget= self.ui.print_record_table.cellWidget(r, 9))
            self.result_print_record[r] = result[0]
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Fail to load data: {e}")      
    
    @QtCore.pyqtSlot()
    def insert_row(self,form = None, target = None):
        current_row = target.rowCount()
        target.insertRow(current_row)
        if form == None:
            data = [tuple(["" for _ in range(10)])]
            self.add_item_into_print_record(data = data, c = 0, r = current_row, is_readOnly = True )
            self.add_item_into_print_record(data = data, c = 1, r = current_row, is_readOnly = False )
            self.add_item_into_print_record(data = data, c = 3, r = current_row, is_readOnly = False )
            self.add_item_into_print_record(data = data, c = 4, r = current_row, is_readOnly = True )
            self.add_item_into_print_record(data = data, c = 5, r = current_row, is_readOnly = False )
            self.add_item_into_print_record(data = data, c = 6, r = current_row, is_readOnly = False )
            self.add_item_into_print_record(data = data, c = 7, r = current_row, is_readOnly = True )
            self.add_item_into_print_record(data = data, c = 8, r = current_row, is_readOnly = True )
            self.add_item_into_print_record(data = data, c = 9, r = current_row, is_readOnly = False )
            self.result_print_record += data
        return current_row
    
    @QtCore.pyqtSlot()
    def delete_row(self,form = None, target = None, save_list = None,oneFile = None):
        current_row = target.currentRow()
        if form == None:
            for col in range(target.columnCount()):
                editor = target.cellWidget(current_row, col)
                if editor is not None:
                    try:
                        editor.editingFinished.disconnect()
                    except (TypeError, RuntimeError):
                        pass
                    try:
                        editor.textChanged.disconnect()
                    except (TypeError, RuntimeError):
                        pass
            try:
                save_list.pop(current_row)
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Error: {e}")
        else:
            try:
                code = target.item(current_row, 0).text()
            except:
                code = "text"
            try:
                if oneFile == False:
                    index = next(i for i, row in enumerate(save_list) if row[1]["machine_code"] == code)
                    save_list.pop(index)
                else:
                    index = next(i for i, row in enumerate(save_list) if row[1]["machine_code"] == code)
                    save_list[index][1] = "text" 
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Error: {e}")
        target.removeRow(current_row)
    
    @QtCore.pyqtSlot()
    def print_record(self):
        role, dep = self.login_info['role_level'], self.login_info['department']
        if role not in ["Admin", "Manager"] and dep != self.department_print_record:
            QtWidgets.QMessageBox.information(self, "Permission denied", "You don't have permission to print maintenance record")
            return

        rows = self.ui.print_record_table.rowCount()
        cols = self.ui.print_record_table.columnCount()
        self.duplicate = []
        attached_machine = {}
        try:
            res_pending = self.database_process.query('''
                SELECT m.machine_code 
                FROM Record_pending AS rp
                JOIN Machines AS m ON rp.machine_id = m.machine_id
            ''')
            res_attach_exist = self.database_process.query('''
                SELECT m.machine_code AS attach_machine, m2.machine_code AS attach_of
                FROM Record_pending AS rp
                JOIN Machines AS m ON rp.machine_id = m.machine_id
                JOIN Machines AS m2 ON rp.attached_equipment = m2.machine_id
            ''')
            self.record_pending = [r[0] for r in res_pending]
            exist_map = {r[0]: r[1] for r in res_attach_exist} 
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Database load failed: {e}")
            return
        for r in range(rows):
            for c in range(cols):
                if c == 5:
                    continue
                elif c == 2:
                    editor = self.ui.print_record_table.cellWidget(r, 0)
                    main_code = editor.text().strip()
                    text = self.ui.print_record_table.model().index(r, c).data(QtCore.Qt.EditRole)
                    if text is None or text == "":
                        continue
                    attached_machine[main_code] = [p.strip() for p in text.split(';') if p and p.strip()]
                elif c == 7:
                    try:
                        editor = self.ui.print_record_table.cellWidget(r, c)
                        text = editor.text().strip()
                        if not STRICT_DATE.match(text):
                            raise ValueError("Format must be YYYY-MM-DD")
                        dt.datetime.strptime(text, "%Y-%m-%d")
                    except ValueError:
                        QtWidgets.QMessageBox.critical(self, "Error", f"Please enter the correct date format (YYYY-MM-DD) at row {r+1}, column {c+1}")
                        return
                else:
                    try:
                        editor = self.ui.print_record_table.cellWidget(r, c)
                        text = editor.text().strip()
                        if editor is None or text == "" or text == "None":
                            QtWidgets.QMessageBox.critical(self, "Error", f"Please fill in all blanks or cells with 'None' content")
                            return
                    except Exception as e:
                        QtWidgets.QMessageBox.critical(self, "Error", f"Error at row {r+1}, column {c+1}: {e}")
                        return
            if (main_code in self.record_pending) and (main_code not in  exist_map.keys()) and (main_code not in  exist_map.values()):
                reply = QtWidgets.QMessageBox.question(
                    self, "Record Exists",
                    f"Record for {main_code} already printed.\nDo you want to print again?",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                    QtWidgets.QMessageBox.No
                )
                if reply == QtWidgets.QMessageBox.No:
                    self.result_print_record = [rec for rec in self.result_print_record if rec[0] != main_code]
                    continue
                try:
                    self.database_process.query(
                        sql='''
                            UPDATE record_pending
                            SET line_id = (SELECT line_id FROM Production_Lines WHERE line_name = :line),
                                technical = :tech,
                                maintenance_date = :date
                            WHERE machine_id = (SELECT machine_id FROM Machines WHERE machine_code = :code);
                        ''',
                        params={
                            'line': self.ui.print_record_table.cellWidget(r, 4).text(),
                            'tech': self.ui.print_record_table.cellWidget(r, 8).text(),
                            'date': self.ui.print_record_table.cellWidget(r, 7).text(),
                            'code': main_code }
                    )
                    self.duplicate.append(r)
                except Exception as e:
                    QtWidgets.QMessageBox.critical(self, "Error", f"Update failed: {e}")
                    return
            elif (main_code in self.record_pending) and ((main_code in  exist_map.keys()) or (main_code in  exist_map.values())):
                partner = exist_map.get(main_code)
                if partner is None:
                    partner = next((k for k, v in exist_map.items() if v == main_code), None)

                code_list = [main_code]
                if partner:
                    code_list.append(partner)

                self.database_process.query(
                    sql='''
                        DELETE FROM record_pending
                        WHERE machine_id IN (
                            SELECT machine_id FROM Machines WHERE machine_code IN :code_list
                        );
                    ''',
                    params={'code_list': code_list}
                )
        try:
            attach_codes = [c for lst in attached_machine.values() for c in lst]
            dup = [c for c in set(attach_codes) if attach_codes.count(c) > 1]
            if dup:
                QtWidgets.QMessageBox.critical(self, "Error", f"Duplicated attached equipment: {', '.join(dup)}")
                return

            if attach_codes:
                code_list = ', '.join(f"'{c}'" for c in attach_codes)
                query = f'''
                    SELECT DISTINCT m.machine_code
                    FROM Machines AS m
                    JOIN Maintenance_plan AS mp ON m.machine_id = mp.machine_id
                    JOIN Production_Lines AS p ON mp.line_id = p.line_id
                    JOIN Departments AS d ON p.department_id = d.department_id
                    JOIN Months_Years AS my ON mp.month_year_id = my.month_year_id
                    WHERE my.year = :year
                    AND d.department_name = :dep
                    AND m.machine_code IN ({code_list})
                '''
                res_valid = self.database_process.query(query, params={'year': self.year_num, 'dep': dep})
                valid_codes = {r[0] for r in res_valid}

                for main, attaches in attached_machine.items():
                    for ac in attaches:
                        if ac not in valid_codes:
                            QtWidgets.QMessageBox.critical(
                                self, "Error",
                                f"Attached equipment {ac} not in your department or has no maintenance plan."
                            )
                            return
                        if ac in exist_map and exist_map[ac] != main:
                            QtWidgets.QMessageBox.critical(
                                self, "Error",
                                f"Attached equipment {ac} already goes with another machine, you need to print again record of machine {main} without {ac}"
                            )
                            return
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Validation failed: {e}")
            return
        self.print_selector = Print_selector(self,quantity=len(self.result_print_record),data= self.result_print_record,attached_machine = attached_machine ,database= self.database_process, duplicate = self.duplicate)
        self.print_selector.show()
    
    @QtCore.pyqtSlot()
    def register_new_form(self):
        if (not hasattr(self, "form_modification")) or sip.isdeleted(self.form_modification):
            self.form_modification = Form_Modification(parent = self)
        self.form_modification.register_form_page()
        self.form_modification.show()
        self.form_modification.ui.stackedWidget.setCurrentWidget(self.form_modification.ui.register_form_page)
    
    @QtCore.pyqtSlot()
    def update_register_form(self):
        if (not hasattr(self, "form_modification")) or sip.isdeleted(self.form_modification):
            self.form_modification = Form_Modification(parent = self)
        self.form_modification.update_form_page()
        self.form_modification.show()
        self.form_modification.ui.stackedWidget.setCurrentWidget(self.form_modification.ui.update_form_page)
    
    @QtCore.pyqtSlot()
    def clear_print_data(self):
        self.ui.print_record_table.clearContents()
        self.ui.print_record_table.setRowCount(0)
        self.result_print_record = []

    @QtCore.pyqtSlot()
    def Mainten_Input_page(self):
        self.style_button_with_shadow((self.ui.Main_Input_record_btn,self.ui.Main_detail_plan_btn,self.ui.Main_Home_btn,self.ui.Main_Print_record_btn))
        self.wrong_scan = []
        try:
            model = self.ui.Main_pending_table.model()
            model.removeRows(0, model.rowCount())
        except AttributeError:
            pass
        self.scan_result_list = None
        self.scan_QRcode_link = None
        self.ui.Maintenance_stacked.setCurrentWidget(self.ui.Input_page_M)
        self.ui.Main_save_link.setPlaceholderText(r"\\172.30.73.156\nd_ie2\Noise Device - IE Data\DANH MUC THIET BI BAO TRI\2025")
        self.ui.Main_scan_link.setPlaceholderText("Just accept folder or pdf file")
        if self.ui.Main_update_group_cbb.count() <= 0:
            self.ui.Main_update_group_cbb.addItems([""]+[group[0] for group in self.group])
            self.ui.Main_update_group_cbb.setCurrentIndex(0)
        self.data_model = QtGui.QStandardItemModel()
        headers = ["Machine code", "Machine Name", "Group","Line","Technical","Maintenance\ndate", "Attached of"]
        self.data_model.setHorizontalHeaderLabels(headers)
        headers = ["Machine code", "Machine Name", "Group","Line","Technical","Maintenance\ndate","Next due\ndate","Attached\nof","Page\nNumber"]
        self.ui.Main_scan_result_table.setColumnCount(len(headers))
        self.ui.Main_scan_result_table.setHorizontalHeaderLabels(headers)
        self.ui.Main_pending_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.Main_pending_table.setAlternatingRowColors(True)
        self.ui.Main_scan_result_table.setAlternatingRowColors(True)
        def job():
            try:
                PE1_pending = self.database_process.query(sql = ''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name = "PE1"''')
                PE2_pending = self.database_process.query(sql = ''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name = "PE2"''')
                PE3_pending = self.database_process.query(sql = ''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name = "PE3"''')
                PE5_pending = self.database_process.query(sql = ''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name = "PE5"''')
                PE4_pending = self.database_process.query(sql = ''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name = "PE4"''')
                ELSE_pending = self.database_process.query(sql = ''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name NOT IN ("PE1","PE2","PE3","PE5","PE4") ''')
                result_pending = PE1_pending + PE2_pending + PE3_pending + PE4_pending + PE5_pending + ELSE_pending
                pending_record_dep = {"PE1": {(code[0],code[3],str(code[5])) for code in PE1_pending}, "PE2":{(code[0],code[3],str(code[5])) for code in PE2_pending},"PE3":{(code[0],code[3],str(code[5])) for code in PE3_pending},"PE4":{(code[0],code[3],str(code[5])) for code in PE4_pending},"PE5":{(code[0],code[3],str(code[5])) for code in PE5_pending},"ELSE":{(code[0],code[3],str(code[5])) for code in ELSE_pending}}
                temp = [code[0] for code in result_pending]
                placeholders = ",".join(f"'{k}'" for k in temp)
                sql = f'''
                    SELECT machine_code,maintenance_frequency
                    FROM machines
                    WHERE machine_code IN ({placeholders})
                '''
                result  = self.database_process.query(sql = sql )
                maintenance_frequency_dict = dict(result)
            except Exception as e:
                return {"result_pending":False,"error":e}
            return {"result_pending":result_pending , "pending_record_dep":pending_record_dep ,"maintenance_frequency_dict":maintenance_frequency_dict }
        self.worker = WorkerThread(job)
        self.worker.finished.connect(lambda data: self.on_update_data_ready(data=data))
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.start()
    
    @QtCore.pyqtSlot()
    def on_update_data_ready(self,data):
        if not data["result_pending"]:
            error_message = str(data['error'])  
            if "')' at line 3" in error_message:
                QtWidgets.QMessageBox.information(self, "Information", "All records are up to date. No pending records found.")
                return       
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {error_message}")
            return
        try:
            self.ui.Main_scan_result_table.clearContents()
            self.ui.Main_scan_result_table.setRowCount(0)
            result_pending, self.pending_record_dep, self.maintenance_frequency_dict = (data["result_pending"],data["pending_record_dep"],data["maintenance_frequency_dict"])
            self.add_data_to_model(data = result_pending,target = self.ui.Main_pending_table,model = self.data_model)
            self.ui.Main_pending_table.setColumnWidth(0,100)
            self.ui.Main_pending_table.setColumnWidth(1,200)
            self.ui.Main_pending_table.setColumnWidth(2,50)
            self.ui.Main_pending_table.setColumnWidth(3,50)
            self.ui.Main_pending_table.setColumnWidth(4,60)
            self.ui.Main_pending_table.setColumnWidth(5,90)
            self.ui.Main_pending_table.setColumnWidth(6,100)
            self.ui.Main_scan_result_table.setColumnWidth(0,100)
            self.ui.Main_scan_result_table.setColumnWidth(1,180)
            self.ui.Main_scan_result_table.setColumnWidth(2,50)
            self.ui.Main_scan_result_table.setColumnWidth(3,50)
            self.ui.Main_scan_result_table.setColumnWidth(4,60)
            self.ui.Main_scan_result_table.setColumnWidth(5,90)
            self.ui.Main_scan_result_table.setColumnWidth(6,90)
            self.ui.Main_scan_result_table.setColumnWidth(7,90)
            self.ui.Main_scan_result_table.setColumnWidth(8,50) 
            self.safe_connect(self.ui.Main_scan_btn.clicked,lambda _: self.scan_record())
            self.safe_connect(self.ui.Main_update_group_cbb.currentIndexChanged,lambda _: self.change_text_pending(dep_changed=True))
            self.safe_connect(self.ui.Main_update_line_cbb.currentIndexChanged,lambda _: self.change_text_pending())
            self.safe_connect(self.ui.Main_Scan_result_insert_btn.clicked, lambda _: self.insert_scan_result_row())
            self.list_of_keys = ["machine_code","machine_name","group","line","technical","maintenance_date","attached_machine","page_num"]
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
    
    @QtCore.pyqtSlot()
    def scan_record(self):
        self.scan_result_final = []
        try:
            self.ui.Main_scan_result_table.clearContents()
            self.ui.Main_scan_result_table.setRowCount(0)
            link = self.ui.Main_scan_link.text()
            link = link.strip().replace('"','').replace("'",'')
            if not link:
                QtWidgets.QMessageBox.critical(self, "Error", f"Please enter the path to the record scan folder.")
                return
            self.scan_QRcode_link = self.scan_QRcode.paths(link)
            self.safe_connect(self.ui.Main_delete_row_btn.clicked, lambda _: self.delete_row(form="update_table", target=self.ui.Main_scan_result_table, save_list = self.scan_QRcode_link,oneFile= self.scan_QRcode.oneFile))
            self.safe_connect(self.ui.Main_update_btn.clicked, lambda _: self.update_record())
            self.safe_connect(self.ui.Main_sync_missing_data_btn.clicked, lambda _: self.Sync_missing_data())
            self.scan_worker = Worker_Pool(self.scan_job, self.scan_QRcode_link)
            self.progress_window = Printer_progress(max=len(self.scan_QRcode_link), text="scanned")
            self.progress_window.ui.label.setText("Scanning...")
            self.scan_worker.signals.progress_changed.connect(lambda value: self.progress_window.update_progress(value=value))
            self.scan_worker.signals.finished.connect(self.progress_window.on_finished)
            self.scan_worker.signals.result_ready.connect(lambda row,scan_result :self.update_table_row(row,scan_result))
            self.scan_worker.signals.error.connect(lambda msg: 
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {msg}"))
            self.progress_window.show()
            QtCore.QThreadPool.globalInstance().start(self.scan_worker)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to scan data: {e}")
            self.progress_window.close()

    def scan_job(self, paths):
        if os.path.isdir(self.scan_QRcode.link):
            self.scan_result_list = []
            for i, path in enumerate(paths):
                try:
                    scan_result = self.scan_QRcode.scanning_dir(path)
                    try:
                        scan_result = json.loads(scan_result) 
                        self.scan_QRcode_link[i] = [self.scan_QRcode_link[i],] + [scan_result]
                        self.scan_result_list.append(scan_result)
                    except json.JSONDecodeError:
                        self.scan_QRcode_link[i] = [self.scan_QRcode_link[i],] + ["text"]
                        self.scan_result_list.append("text")
                except Exception as e:
                    self.scan_worker.signals.error.emit(str(e))  
                self.scan_worker.signals.progress_changed.emit(i+1)
            for row, item in enumerate(self.scan_result_list):
                self.scan_worker.signals.result_ready.emit(row, item)
        else:
            try:
                temp = []
                self.take_code_and_page = []
                self.scan_result_list = self.scan_QRcode.scanning_oneFile(paths[0])
                for row, item in enumerate(self.scan_result_list):
                    try:
                        scan_result = json.loads(item)
                        temp.append([self.scan_QRcode.link,] + [scan_result])
                        self.scan_worker.signals.result_ready.emit(row, scan_result)
                        self.take_code_and_page.append(scan_result["machine_code"])
                    except json.JSONDecodeError:
                        continue
                self.scan_QRcode_link = temp
            except Exception as e:
                self.scan_worker.signals.error.emit(str(e))  
    
    @QtCore.pyqtSlot()           
    def update_table_row(self, row, scan_result):
        table_row = self.insert_row(form="update_table", target=self.ui.Main_scan_result_table)
        for col in range(self.ui.Main_scan_result_table.columnCount()):
            try: 
                self.add_item_to_scan_result(table_row,col,scan_result)
            except Exception as e:
                if col ==  0:
                    current_row = self.ui.Main_scan_result_table.rowCount() - 1
                    editor = QtWidgets.QLineEdit()
                    editor.setStyleSheet(''' border: none;''')
                    self.safe_connect( editor.textChanged, 
                        lambda text, r=current_row, c=0: self.on_text_changed(text = text, c = c, r = r, target = self.ui.Main_scan_result_table , 
                                                                              select_col = "machine_code" , table = "View_Record_Pending ", where = f"WHERE machine_code LIKE '%{text}%'")
                    )
                    self.safe_connect( editor.editingFinished, lambda r=current_row: self.load_pending_record(row = current_row))
                    self.ui.Main_scan_result_table.setCellWidget(current_row,0,editor)
    
    @QtCore.pyqtSlot()
    def change_text_pending(self,dep_changed = False):
        self.data_model.removeRows(0, self.data_model.rowCount())
        sql = '''SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                                        FROM `View_Record_Pending`'''
        dep = self.ui.Main_update_group_cbb.currentText()
        line = self.ui.Main_update_line_cbb.currentText()
        if dep_changed:
            line_list = self.database_process.query(sql = "SELECT DISTINCT line_name FROM `View_Record_Pending` WHERE department_name = :dep",params = {'dep':dep})
            self.ui.Main_update_line_cbb.clear()
            self.ui.Main_update_line_cbb.addItem("")
            self.ui.Main_update_line_cbb.addItems([line[0] for line in line_list])
        if dep == "":
            self.ui.Main_update_line_cbb.clear()
            sql = sql + ';'
        else:
            sql = sql + ' WHERE department_name = :dep'         
            if line != "":
                sql = sql + ' AND line_name = :line;'
            else:
                sql = sql + ';'
        try:
            result = self.database_process.query(sql = sql,params = {'dep':dep,'line':line} )
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
        self.add_data_to_model(result,self.ui.Main_pending_table,self.data_model)
    
    @QtCore.pyqtSlot()  
    def update_record(self):
        if self.login_info['role_level'] not in ["Secretary","Manager","Admin"]:
            QtWidgets.QMessageBox.information(self,"Permission denied","Your don't have permission update maintenance record")
            return
        update_list = []
        machine_update_status = []
        save_link = self.ui.Main_save_link.text()
        jobs = []
        record_link_dict = {}
        if not os.path.exists(save_link):
            QtWidgets.QMessageBox.critical(self, "Error", f"Save link not exists")
            return
        def save_pdf(path,save_link):
            for index in range(len(path)):
                if ( path[index][1] != "text" ) and ( path[index][1]["machine_code"] not in self.wrong_scan ):
                    machine_info_list = [path[index][1]["machine_code"],path[index][1]["machine_name"],path[index][1]["group"],path[index][1]["line"],path[index][1]["technical"],path[index][1]["maintenance_date"]]
                    file_list = [[path[index][1]["machine_code"],"_".join(x.strip().replace(" ", "_") for x in machine_info_list) + ".pdf"]]
                    if path[index][1]["attached_machine"] != "" and path[index][1]["attached_machine"] != []:
                        attached_name = self.database_process.query(f'''SELECT machine_name FROM `Machines` WHERE machine_code IN ({",".join([f"'{c}'" for c in path[index][1]["attached_machine"]])});''')
                        for index_1,item in enumerate(path[index][1]["attached_machine"]):
                            attached_file = [item,attached_name[index_1][0]] + machine_info_list[2:]
                            file_list.append([item,"_".join(x.strip().replace(" ", "_") for x in attached_file) + ".pdf"])
                    for file in file_list:
                        record_link_dict[file[0]] = f"{save_link}/{file[1].replace('/', '-')}"
                        if os.path.isdir(self.scan_QRcode.link):
                            pass
                        else:
                            try:
                                result = self.database_process.query('''SELECT machine_code, page_num FROM `maintenance_form_info`;''' )
                                self.page_num_dict = dict(result)
                                pages = int(self.page_num_dict.get(path[index][1]["machine_code"], 1)) 
                                start_page = int(path[index][1]["page_num"])                 
                                to_page = start_page + pages - 1    
                                jobs.append(
                                lambda f=file[1], s=start_page, e=to_page: self.scan_QRcode.split_pdf(
                                    input_file=self.scan_QRcode.link,
                                    start=s,
                                    end=e,
                                    output_file=rf"{save_link}\{f.replace('/', '-')}"
                                )
                            )
                            except Exception as e:
                                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data scan: {e}")
                else: continue
        save_pdf(path = self.scan_QRcode_link,save_link = save_link)
        try:
            for row in range(self.ui.Main_scan_result_table.rowCount()):
                if self.ui.Main_scan_result_table.item(row,0).text() in self.wrong_scan:
                    continue
                else:
                    machine_code = self.ui.Main_scan_result_table.item(row,0).text()
                    line_name =  self.ui.Main_scan_result_table.item(row,3).text()
                    technical =  self.ui.Main_scan_result_table.item(row,4).text()
                    maintenance_date =  self.ui.Main_scan_result_table.item(row,5).text()
                    next_date =  self.ui.Main_scan_result_table.item(row,6).text()
                    record_link = record_link_dict[machine_code]
                    update_list.append({'code':machine_code,
                                        'maintenance_date':maintenance_date,
                                        'technical':technical,
                                        'next_date':next_date,
                                        'line':line_name,
                                        'link':record_link})
                    machine_update_status.append({'code':machine_code,
                                                  'line':line_name,
                                                  'status':'GOOD'})
            success = self.database_process.executemany(sql = f'''INSERT INTO `maintenance_records` ( machine_id, maintenance_date, technician, `next_due_date`, line_id, record_link )
                                                SELECT m.machine_id,
                                                    :maintenance_date,
                                                    :technical,
                                                    :next_date,
                                                    p.line_id,
                                                    :link
                                                FROM `machines` m
                                                JOIN `production_Lines` p ON p.line_name = :line
                                                WHERE m.machine_code = :code;''', params_list = update_list )
            if success > 0:
                QtWidgets.QMessageBox.information(self, "Success", f"Đã cập nhật {success} bản ghi.")
                self.database_process.executemany(sql = '''UPDATE `machines` AS m
                                                            JOIN `production_lines` AS p ON p.line_name = :line
                                                            SET m.machine_status = :status, m.line_id = p.line_id
                                                            WHERE m.machine_code = :code;''', params_list = machine_update_status )                              
                for job in jobs: job()
            else:
                QtWidgets.QMessageBox.warning(self, "No Change", "Không có bản ghi nào được cập nhật.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
        self.ui.Main_scan_result_table.clearContents()
        self.ui.Main_scan_result_table.setRowCount(0)
        self.scan_QRcode.link = ""
        self.change_text_pending()
    
    def add_item_to_scan_result(self,row , col,scan_result):
        if col < 6:
            item = QtWidgets.QTableWidgetItem(str(scan_result[self.list_of_keys[col]]))
            if ( col == 0 ) and ( (scan_result[self.list_of_keys[0]],scan_result[self.list_of_keys[3]],scan_result[self.list_of_keys[5]]) not in self.pending_record_dep.get(f"{scan_result[self.list_of_keys[2]]}",self.pending_record_dep["ELSE"])):
                item.setBackground(QtGui.QColor(255, 80, 80))
                self.wrong_scan.append(str(scan_result[self.list_of_keys[col]]))
                if scan_result[self.list_of_keys[-2]] != "" and scan_result[self.list_of_keys[-2]] != []:
                    self.wrong_scan + scan_result[self.list_of_keys[-2]]
            item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.ui.Main_scan_result_table.setItem(row, col, item)
        elif col == 6:
            date_object = dt.datetime.strptime(self.ui.Main_scan_result_table.item(row,col-1).text(), '%Y-%m-%d').date()
            next_date = date_object + relativedelta(months=int(self.maintenance_frequency_dict[self.ui.Main_scan_result_table.item(row,0).text()]))
            item = QtWidgets.QTableWidgetItem(next_date.strftime("%Y-%m-%d"))
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.ui.Main_scan_result_table.setItem(row, col, item)
        elif col == 7:
            if scan_result[self.list_of_keys[-2]] == "" or scan_result[self.list_of_keys[-2]] == [] :
                return
            else:
                for item in scan_result[self.list_of_keys[-2]]:
                    self.copy_row(row,item)
        else:
            item = QtWidgets.QTableWidgetItem(str(scan_result[self.list_of_keys[-1]]))
            item.setFlags(item.flags() | QtCore.Qt.ItemIsEditable)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.ui.Main_scan_result_table.setItem(row, col, item)

    def copy_row(self, row_index, attached_machine):
        row_count = self.ui.Main_scan_result_table.rowCount()
        self.ui.Main_scan_result_table.insertRow(row_count)
        for column_index in range(self.ui.Main_scan_result_table.columnCount()):
            src_item = self.ui.Main_scan_result_table.item(row_index, column_index)
            if src_item is not None:
                if column_index == 0:
                    new_item = QtWidgets.QTableWidgetItem(str(attached_machine))
                else:
                    new_item = QtWidgets.QTableWidgetItem(src_item.text())
                new_item.setBackground(src_item.background())
                new_item.setForeground(src_item.foreground())
                new_item.setFont(src_item.font())
            else:
                if column_index == 7:
                    text0 = self.ui.Main_scan_result_table.item(row_index, 0)
                    new_item = QtWidgets.QTableWidgetItem(text0.text() if text0 is not None else "")
                    if text0 is not None:
                        new_item.setBackground(text0.background())
                        new_item.setForeground(text0.foreground())
                        new_item.setFont(text0.font())
                else:
                    new_item = QtWidgets.QTableWidgetItem("")
            new_item.setFlags(new_item.flags() & ~QtCore.Qt.ItemIsEditable)
            new_item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.ui.Main_scan_result_table.setItem(row_count, column_index, new_item)

    @QtCore.pyqtSlot()  
    def insert_scan_result_row(self):
        if self.scan_QRcode.link == "" or self.scan_QRcode.link is None:
            QtWidgets.QMessageBox.critical(self, "Error", f"Please scan QR code first.")
            return
        row_count = self.ui.Main_scan_result_table.rowCount()
        self.ui.Main_scan_result_table.insertRow(row_count)
        for column_index in range(self.ui.Main_scan_result_table.columnCount()):
            if column_index != 0:
                new_item = QtWidgets.QTableWidgetItem("")
                new_item.setFlags(new_item.flags() & ~QtCore.Qt.ItemIsEditable)
                new_item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.ui.Main_scan_result_table.setItem(row_count, column_index, new_item)
            else:
                editor = QtWidgets.QLineEdit()
                editor.setStyleSheet(''' border: none;''')
                self.safe_connect( editor.textChanged, 
                    lambda text, r=row_count, c=0: self.on_text_changed(text = text, c = c, r = r, target = self.ui.Main_scan_result_table , 
                                                                          select_col = "machine_code" , table = "View_Record_Pending ", where = f"WHERE machine_code LIKE '%{text}%'")
                )
                self.safe_connect( editor.editingFinished, lambda r=row_count: self.load_pending_record(row = r))
                self.ui.Main_scan_result_table.setCellWidget(row_count,0,editor)

    @QtCore.pyqtSlot() 
    def load_pending_record(self,row):
        code = self.ui.Main_scan_result_table.cellWidget(row,0).text()
        for item in self.scan_QRcode_link:
            if item[1]["machine_code"] == code:
                return
        try:
            result = self.database_process.query(sql = ''' SELECT * FROM  `View_Record_Pending` WHERE machine_code = :code''',params= {'code':code})
            if not result:
                QtWidgets.QMessageBox.warning(self, "Not found", f"Machine code '{code}' not found in pending records.")
                return
            page_num, ok = QtWidgets.QInputDialog.getInt(
                self,
                "Input page number",
                f"Enter page number for machine {result[0][0]}:",
                value=1,
                min=0,
                max=9999,
                step=1
            )
            if not ok:
                return
            scan_result_dict = {
                "machine_code": result[0][0],
                "machine_name": result[0][1],
                "group": result[0][2],
                "line": result[0][3],
                "technical": result[0][4],
                "maintenance_date": str(result[0][5]),
                "attached_machine": result[0][6].split(",") if result[0][6] else "",
                "page_num": page_num}
            self.scan_QRcode_link.append([self.scan_QRcode.path[0],scan_result_dict])
            for col in range(self.ui.Main_scan_result_table.columnCount()):
                self.add_item_to_scan_result(row,col,scan_result_dict)                                                                                                                                                            
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
    
    @QtCore.pyqtSlot() 
    def Sync_missing_data(self):
        if not self.scan_QRcode_link:
            QtWidgets.QMessageBox.information(self, "Info", "No scanned data to sync.")
            return
        existing_data = []
        line_name = self.scan_QRcode_link[0][1]["line"]
        for item in self.scan_QRcode_link:
            if item[1] == "text":
                continue
            if item[1]["line"] != line_name:
                QtWidgets.QMessageBox.critical(self, "Error", "Scanned data contains multiple line names. Please scan data from the same line to sync missing records.")
                return
            existing_data.append(item[1]["machine_code"])
            if item[1]["attached_machine"] != "" and item[1]["attached_machine"] != []:
                existing_data += item[1]["attached_machine"]
        existing_data = list(set(existing_data))
        try:
            temp = self.database_process.query(sql = f''' SELECT * FROM `View_Record_Pending`
                                                            WHERE line_name = :line 
                                                                AND machine_code NOT IN ({",".join([f"'{code}'" for code in existing_data])});''', 
                                                            params = {'line':line_name})
            self.sync_missing_list = {}
            for record in temp:
                self.sync_missing_list[record[0]] =  {
                "machine_code": record[0],
                "machine_name": record[1],
                "group": record[2],
                "line": record[3],
                "technical": record[4],
                "maintenance_date": str(record[5]),
                "attached_machine": record[6].split(",") if record[6] else "",
                "page_num": None
            }
            self.sync_window = Sync_Missing_Data(parent=self, line_name = line_name ,data_list=[[data["machine_code"],data["page_num"]] for key, data in self.sync_missing_list.items()])
            self.sync_window.synced.connect(self.on_missing_data_synced)
            self.sync_window.show()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to sync data: {e}")
    
    @QtCore.pyqtSlot()
    def on_missing_data_synced(self):
        if not self.sync_missing_list:
            return
        for code,data in self.sync_missing_list.items():
            if data["page_num"] is not None:
                row_count = self.ui.Main_scan_result_table.rowCount()
                self.ui.Main_scan_result_table.insertRow(row_count)
                for col in range(self.ui.Main_scan_result_table.columnCount()):
                    self.add_item_to_scan_result(row_count,col,data)
                self.scan_QRcode_link.append([self.scan_QRcode.path[0],data])

    @QtCore.pyqtSlot()
    def Mainten_Detail_plan_page(self):
        self.style_button_with_shadow((self.ui.Main_detail_plan_btn,self.ui.Main_Input_record_btn,self.ui.Main_Home_btn,self.ui.Main_Print_record_btn))
        self.ui.Maintenance_stacked.setCurrentWidget(self.ui.Detail_plan_page_M)
        self.itemChanged = {"update": set(), "insert":set()}
        self.department_maintenance_plan = None
        if self.ui.Group_cbb_DP.count() <= 0:
            self.ui.Year_plan_cbb_DP.clear()
            self.ui.Year_plan_cbb_DP.addItems([str(y) for y in range(2025 , 2035 )])
            self.ui.Year_plan_cbb_DP.setCurrentText(str(self.year_num))
            headers = ["Machine\ncode", "Machine Name", "Group"]
            self.ui.Detail_plan_frze_table.setColumnCount(len(headers))
            self.ui.Detail_plan_frze_table.setHorizontalHeaderLabels(headers)
            self.ui.Detail_plan_frze_table.setColumnWidth(0,80)
            self.ui.Detail_plan_frze_table.setColumnWidth(1,240)
            self.ui.Detail_plan_frze_table.setColumnWidth(2,50)
            self.ui.Detail_plan_frze_table.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
            hearders = ["Week\n(M1)", "Status\n(M1)", "Line\n(M1)","Record\n(M1)","Week\n(M2)", "Status\n(M2)", "Line\n(M2)","Record\n(M2)","Week\n(M3)", "Status\n(M3)", "Line\n(M3)","Record\n(M3)","Week\n(M4)", "Status\n(M4)", "Line\n(M4)","Record\n(M4)",
                        "Week\n(M5)", "Status\n(M5)", "Line\n(M5)","Record\n(M5)","Week\n(M6)", "Status\n(M6)", "Line\n(M6)","Record\n(M6)","Week\n(M7)", "Status\n(M7)", "Line\n(M7)","Record\n(M7)","Week\n(M8)", "Status\n(M8)", "Line\n(M8)","Record\n(M8)",
                        "Week\n(M9)", "Status\n(M9)", "Line\n(M9)","Record\n(M9)","Week\n(M10)", "Status\n(M10)", "Line\n(M10)","Record\n(M10)","Week\n(M11)", "Status\n(M11)", "Line\n(M11)","Record\n(M11)","Week\n(M12)", "Status\n(M12)", "Line\n(M12)","Record\n(M12)"]
            self.ui.Detail_plan_table.setColumnCount(len(hearders))
            self.ui.Detail_plan_table.setHorizontalHeaderLabels(hearders)
            self.ui.Detail_plan_table.verticalHeader().setVisible(False)
            for i in range(self.ui.Detail_plan_table.columnCount()):
                self.ui.Detail_plan_table.setColumnWidth(i,59)
            self.ui.Detail_plan_table.setAlternatingRowColors(True)
            self.ui.Detail_plan_frze_table.setAlternatingRowColors(True)
            self.ui.Detail_plan_table.verticalScrollBar().valueChanged.connect(
                self.ui.Detail_plan_frze_table.verticalScrollBar().setValue
            )
            self.ui.Detail_plan_frze_table.verticalScrollBar().valueChanged.connect(
                self.ui.Detail_plan_table.verticalScrollBar().setValue
            )
            self.ui.Group_cbb_DP.addItems([d[0] for d in self.group])
            self.ui.Group_cbb_DP.setCurrentText(self.login_info['department'])
            self.safe_connect( self.ui.Group_cbb_DP.currentIndexChanged, lambda _: self.group_cbb_DP_change())
            self.safe_connect( self.ui.Code_lnedit_DP.textChanged, lambda text: self.filter_suggestion(target = self.ui.Code_lnedit_DP,
                                                                                        text = "DISTINCT ( m.machine_code )",table = "`Maintenance_plan` as mp",
                                                                                        where = f""" JOIN `Machines` as m
                                                                                                        ON mp.machine_id = m.machine_id
                                                                                                        JOIN `Production_Lines` as p
                                                                                                        ON mp.line_id = p.line_id
                                                                                                        JOIN `Months_Years` as my ON my.month_year_id = mp.month_year_id
                                                                                                        WHERE p.line_name = '{self.ui.Line_cbb_DP.currentText()}' AND m.machine_code LIKE '%{text}%' AND my.year = {self.year_num} """))
            self.safe_connect( self.ui.Load_btn_DP.clicked, lambda _: self.Load_Maintenance_plan())
            self.safe_connect( self.ui.Update_btn_DP.clicked, lambda _: self.Update_maintenance_plan())
            self.safe_connect( self.ui.Delete_btn_DP.clicked, lambda _: self.Delete_plan())
            self.safe_connect( self.ui.Insert_btn_DP.clicked, lambda _: self.Insert_plan())

    @QtCore.pyqtSlot() 
    def group_cbb_DP_change(self):
        dep = self.ui.Group_cbb_DP.currentText()
        try:
            lines = self.database_process.query(sql = '''SELECT DISTINCT( p.line_name )
                                                        FROM  `Maintenance_plan` as mp
                                                        JOIN `Production_Lines` as p
                                                        ON mp.line_id = p.line_id
                                                        JOIN `Departments` as d
                                                        ON p.department_id = d.department_id
                                                        JOIN `Months_Years` as my 
                                                        ON my.month_year_id = mp.month_year_id
                                                        WHERE d.department_name = :dep AND my.year = :year;''', params = {'dep':dep,'year':self.year_num})
            items = [" "] + [line[0] for line in lines]
            self.ui.Line_cbb_DP.clear()
            self.ui.Line_cbb_DP.addItems(items)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot() 
    def Load_Maintenance_plan(self):
        self.pdf_windows = []
        self.itemChanged = {"update": set(), "insert":set()}
        self.ui.Detail_plan_table.clearContents()
        self.ui.Detail_plan_table.setRowCount(0)
        self.ui.Detail_plan_frze_table.clearContents()
        self.ui.Detail_plan_frze_table.setRowCount(0)
        self.department_maintenance_plan = self.ui.Group_cbb_DP.currentText()                         
        column_colors = {
                            range(3, 7): (0, 51, 102,50),
                            range(7, 11): (255, 182, 193,50),
                            range(11, 15): (102, 205, 170,50),
                            range(15, 19): (255, 239, 153,50),
                            range(19, 23): (64, 224, 208,50),
                            range(23, 27): (255, 140, 0,50),
                            range(27, 31): (220, 20, 60,50),
                            range(31, 35): (255, 215, 0,50),
                            range(35, 39): (184, 115, 51,50),
                            range(39, 43): (204, 85, 0,50),
                            range(43, 47): (54, 69, 79,50),
                            range(47, 51): (128, 0, 32,50),
                        }
        def get_color_for_column(col):
            for r, color in column_colors.items():
                if col in r:
                    return color
            return None
        try:
            self.ui.Detail_plan_table.itemChanged.disconnect()
        except:
            pass
        script = "my.year = :year AND d.department_name = :dep "
        params = {'dep':self.ui.Group_cbb_DP.currentText(),'year':self.ui.Year_plan_cbb_DP.currentText()}
        if self.ui.Line_cbb_DP.currentText() != " " and self.ui.Line_cbb_DP.currentText() != "" :
            script += " AND p.line_name = :line"
            params['line'] = self.ui.Line_cbb_DP.currentText()
        if self.ui.Code_lnedit_DP.text() != "":
            script += f" AND m.machine_code LIKE '%{self.ui.Code_lnedit_DP.text()}%'"
        try:
            result = self.database_process.query(sql = f''' 
                                        SELECT
                                        m.machine_code,
                                        m.machine_name,
                                        d.department_name,
                                        MAX(CASE WHEN my.month = 1 AND mp.quarter = 1 THEN mp.week END) AS week_q1,
                                        MAX(CASE WHEN my.month = 1 AND mp.quarter = 1 THEN status END) AS status_q1,
                                        MAX(CASE WHEN my.month = 1 AND mp.quarter = 1 THEN p.line_name END) AS line_q1,
                                        MAX(CASE WHEN my.month = 1 AND mp.quarter = 1 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q1,
                                        MAX(CASE WHEN my.month = 2 AND mp.quarter = 1 THEN week END) AS week_q2,
                                        MAX(CASE WHEN my.month = 2 AND mp.quarter = 1 THEN status END) AS status_q2,
                                        MAX(CASE WHEN my.month = 2 AND mp.quarter = 1 THEN p.line_name END) AS line_q2,
                                        MAX(CASE WHEN my.month = 2 AND mp.quarter = 1 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q2,
                                        MAX(CASE WHEN my.month = 3 AND mp.quarter = 1 THEN week END) AS week_q3,
                                        MAX(CASE WHEN my.month = 3 AND mp.quarter = 1 THEN status END) AS status_q3,
                                        MAX(CASE WHEN my.month = 3 AND mp.quarter = 1 THEN p.line_name END) AS line_q3,
                                        MAX(CASE WHEN my.month = 3 AND mp.quarter = 1 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q3,
                                        MAX(CASE WHEN my.month = 4 AND mp.quarter = 2 THEN week END) AS week_q4,
                                        MAX(CASE WHEN my.month = 4 AND mp.quarter = 2 THEN status END) AS status_q4,
                                        MAX(CASE WHEN my.month = 4 AND mp.quarter = 2 THEN p.line_name END) AS line_q4,
                                        MAX(CASE WHEN my.month = 4 AND mp.quarter = 2 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q4,
                                        MAX(CASE WHEN my.month = 5 AND mp.quarter = 2 THEN mp.week END) AS week_q5,
                                        MAX(CASE WHEN my.month = 5 AND mp.quarter = 2 THEN status END) AS status_q5,
                                        MAX(CASE WHEN my.month = 5 AND mp.quarter = 2 THEN p.line_name END) AS line_q5,
                                        MAX(CASE WHEN my.month = 5 AND mp.quarter = 2 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q5,
                                        MAX(CASE WHEN my.month = 6 AND mp.quarter = 2 THEN week END) AS week_q6,
                                        MAX(CASE WHEN my.month = 6 AND mp.quarter = 2 THEN status END) AS status_q6,
                                        MAX(CASE WHEN my.month = 6 AND mp.quarter = 2 THEN p.line_name END) AS line_q6,
                                        MAX(CASE WHEN my.month = 6 AND mp.quarter = 2 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q6,
                                        MAX(CASE WHEN my.month = 7 AND mp.quarter = 3 THEN week END) AS week_q7,
                                        MAX(CASE WHEN my.month = 7 AND mp.quarter = 3 THEN status END) AS status_q7,
                                        MAX(CASE WHEN my.month = 7 AND mp.quarter = 3 THEN p.line_name END) AS line_q7,
                                        MAX(CASE WHEN my.month = 7 AND mp.quarter = 3 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q7,
                                        MAX(CASE WHEN my.month = 8 AND mp.quarter = 3 THEN week END) AS week_q8,
                                        MAX(CASE WHEN my.month = 8 AND mp.quarter = 3 THEN status END) AS status_q8,
                                        MAX(CASE WHEN my.month = 8 AND mp.quarter = 3 THEN p.line_name END) AS line_q8,
                                        MAX(CASE WHEN my.month = 8 AND mp.quarter = 3 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q8,
                                        MAX(CASE WHEN my.month = 9 AND mp.quarter = 3 THEN mp.week END) AS week_q9,
                                        MAX(CASE WHEN my.month = 9 AND mp.quarter = 3 THEN status END) AS status_q9,
                                        MAX(CASE WHEN my.month = 9 AND mp.quarter = 3 THEN p.line_name END) AS line_q9,
                                        MAX(CASE WHEN my.month = 9 AND mp.quarter = 3 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q9,
                                        MAX(CASE WHEN my.month = 10 AND mp.quarter = 4 THEN week END) AS week_q10,
                                        MAX(CASE WHEN my.month = 10 AND mp.quarter = 4 THEN status END) AS status_q10,
                                        MAX(CASE WHEN my.month = 10 AND mp.quarter = 4 THEN p.line_name END) AS line_q10,
                                        MAX(CASE WHEN my.month = 10 AND mp.quarter = 4 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q10,
                                        MAX(CASE WHEN my.month = 11 AND mp.quarter = 4 THEN week END) AS week_q11,
                                        MAX(CASE WHEN my.month = 11 AND mp.quarter = 4 THEN status END) AS status_q11,
                                        MAX(CASE WHEN my.month = 11 AND mp.quarter = 4 THEN p.line_name END) AS line_q11,
                                        MAX(CASE WHEN my.month = 11 AND mp.quarter = 4 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q11,
                                        MAX(CASE WHEN my.month = 12 AND mp.quarter = 4 THEN week END) AS week_q12,
                                        MAX(CASE WHEN my.month = 12 AND mp.quarter = 4 THEN status END) AS status_q12,
                                        MAX(CASE WHEN my.month = 12 AND mp.quarter = 4 THEN p.line_name END) AS line_q12,
                                        MAX(CASE WHEN my.month = 12 AND mp.quarter = 4 AND mr.maintenance_date = mp.maintenance_date THEN mr.record_link END) AS record_link_Q12                                       
                                        FROM `Maintenance_plan` AS mp
                                        JOIN `Machines` AS m
                                        ON mp.machine_id = m.machine_id
                                        JOIN `Production_Lines` AS p
                                        ON mp.line_id = p.line_id
                                        JOIN `Departments` as d
                                        ON d.department_id = p.department_id
                                        JOIN `Months_Years` as my
                                        ON mp.month_year_id = my.month_year_id 
                                        LEFT JOIN `Maintenance_records` as mr
                                        ON mr.machine_id = mp.machine_id
                                        WHERE {script}
                                        GROUP BY    m.machine_code,
                                                    m.machine_name,
                                                    d.department_name
                                        ORDER BY week_q1 ASC;''', params = params)
            self.ui.Detail_plan_table.setRowCount(len(result))
            self.ui.Detail_plan_frze_table.setRowCount(len(result))
            for row in range(len(result)):
                for col in range(len(result[row])):
                    if result[row][col] is None:
                        item = QtWidgets.QTableWidgetItem("")
                    else:
                        item = QtWidgets.QTableWidgetItem(str(result[row][col]))
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    if col < 3:
                        item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                        self.ui.Detail_plan_frze_table.setItem(row, col, item)
                    else:
                        color = get_color_for_column(col)
                        item.setBackground(QtGui.QColor(color[0],color[1],color[2],color[3]))
                        if result[row][col] == "Overdue":
                            item.setForeground(QtGui.QBrush(QtGui.QColor(255, 0, 0)))
                        elif result[row][col] == "Ontime":
                            item.setForeground(QtGui.QBrush(QtGui.QColor(0, 128, 0))) 
                        elif str(result[row][col]).lower().endswith(".pdf"):
                            btn = QtWidgets.QPushButton("")
                            icon = QtGui.QIcon()
                            icon.addFile(resource_path(u"Icons/hyperlink.ico"), QtCore.QSize(), QtGui.QIcon.Normal, QtGui.QIcon.Off)
                            btn.setIcon(icon)
                            base_style = """
                                            QPushButton:hover {
                                                background-color: rgba(255, 183, 153, 80); 
                                                border-width: 1px;
                                                border-top-color: rgb(255,150,60);
                                                border-right-color: qlineargradient(spread:pad, x1:0, y1:1, x2:1, y2:0, stop:0 rgba(200, 70, 20, 255), stop:1 rgba(255,150,60, 255));
                                                border-left-color: qlineargradient(spread:pad, x1:1, y1:0, x2:0, y2:0, stop:0 rgba(200, 70, 20, 255), stop:1 rgba(255,150,60, 255));
                                                border-bottom-color: rgb(200,70,20);
                                                border-style: solid;
                                                padding: 2px;
                                            }
                                        """
                            dynamic_style = f"QPushButton {{ border: none; background-color: rgba({color[0]},{color[1]},{color[2]},{color[3]}); }}"
                            btn.setStyleSheet(dynamic_style + base_style)
                            self.safe_connect( btn.clicked, lambda _, link = result[row][col].lower(): self.open_pdf( link = link))
                            self.ui.Detail_plan_table.setCellWidget(row, col - 3, btn)
                            for index in range(1,4):
                                item = self.ui.Detail_plan_table.item(row,col - 3 -index)
                                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                            continue
                        self.ui.Detail_plan_table.setItem(row, col - 3, item)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
        self.ui.Detail_plan_table.itemChanged.connect(lambda item: self.on_item_in_Detail_plan_table_change(item=item))
    
    @QtCore.pyqtSlot() 
    def on_item_in_Detail_plan_table_change(self, item):
        row = item.row()
        col = item.column()
        if row not in self.itemChanged["insert"]:
            self.itemChanged["update"].add(row)
        item.setBackground(QtGui.QColor(255, 255, 150))
    
    @QtCore.pyqtSlot() 
    def Update_maintenance_plan(self):
        if self.login_info["role_level"] in ["Manager","Admin"]:
                pass
        elif ( self.login_info["department"] == self.department_maintenance_plan ) and ( self.login_info["role_level"] in ["Supervisor"]):
            pass
        else:
            QtWidgets.QMessageBox.information(self,"Permission denied","Your don't have permission to update this machine info")
            return
        update_list = []
        insert_list = []
        finish = 0
        year = int(self.ui.Year_plan_cbb_DP.currentText())
        def update_job(type :str, list:list):
            for r in self.itemChanged[type]:
                try:
                    code = self.ui.Detail_plan_frze_table.item(r,0).text()
                except:
                    code = self.ui.Detail_plan_frze_table.cellWidget(r,0).text()
                for col_offset in range(0,12):
                    if not self.ui.Detail_plan_table.item(r,0+4*col_offset):
                        continue    
                    else:
                        line = self.ui.Detail_plan_table.item(r,2+4*col_offset).text()
                        week = self.ui.Detail_plan_table.item(r,0+4*col_offset).text()
                        try:
                            status = self.ui.Detail_plan_table.item(r,1+4*col_offset).text()
                        except:
                            status = ""
                        if week == "":
                            new_month = ""
                            quarter = ""
                        else:
                            new_month = self.ui.company_week_month(year,int(week))
                            quarter = (new_month - 1) // 3 + 1
                        old_month = col_offset + 1
                        list.append({'code':code,'line':line,'quarter':quarter,'new_month':new_month,'year':year,'week':week,'old_month':old_month,'status':status})
        if self.itemChanged["update"]:
            update_job(type = "update", list = update_list)
            delete_month = []
            update_list_month = []
            for item in update_list:
                if item['week'] == "" and item['line'] == "":
                    delete_month.append({'code':item['code'],'del_month':item['old_month'], 'year':item['year']})
                else:
                    update_list_month.append(item)
            try:
                finish += self.database_process.executemany(sql=''' DELETE mp
                                                                FROM `Maintenance_plan` AS mp
                                                                JOIN `Machines` AS m
                                                                ON mp.machine_id = m.machine_id
                                                                JOIN `Months_Years` AS my
                                                                ON mp.month_year_id = my.month_year_id
                                                                WHERE m.machine_code = :code AND my.month = :del_month AND my.year = :year; ''', params_list=delete_month)
                check_sql = '''
                            SELECT 1
                            FROM `Maintenance_plan` AS mp
                            JOIN `Machines` AS m ON mp.machine_id = m.machine_id
                            JOIN `Months_Years` AS my ON mp.month_year_id = my.month_year_id
                            WHERE m.machine_code = :code AND my.month = :old_month AND my.year = :year;
                            '''
                for data in update_list_month:
                    exists = self.database_process.query(sql=check_sql, params={'code': data['code'], 'old_month': data['old_month'], 'year': data['year']})
                    if len(exists)>0:
                        finish += self.database_process.query(sql=''' UPDATE Maintenance_plan AS mp 
                                                                            JOIN Machines AS m ON mp.machine_id = m.machine_id 
                                                                            JOIN Months_Years AS my ON mp.month_year_id = my.month_year_id 
                                                                            SET 
                                                                            mp.line_id = ( SELECT p.line_id FROM Production_Lines as p WHERE p.line_name = :line ),
                                                                            mp.month_year_id = ( SELECT my2.month_year_id FROM Months_Years as my2 WHERE my2.month = :new_month AND my2.year = :year ), 
                                                                            mp.week = :week,
                                                                            mp.status = :status
                                                                    WHERE m.machine_code = :code AND my.month = :old_month AND my.year = :year ;''', params=data)
                        self.database_process.query(sql = '''UPDATE machines
                                                            SET machine_status = 'GOOD'
                                                            WHERE machine_code = :code;''', params = {'code':data['code']})
                    else:
                        finish += self.database_process.query(sql=''' INSERT INTO `Maintenance_plan` (machine_id,line_id,month_year_id,quarter,week,original_week)
                                                                            SELECT m.machine_id,p.line_id,my.month_year_id,:quarter,:week,:week
                                                                            FROM `Machines` as m
                                                                            JOIN `Production_Lines` as p 
                                                                            ON p.line_name = :line
                                                                            JOIN `Months_Years` as my
                                                                            ON my.month = :new_month AND my.year = :year
                                                                            WHERE m.machine_code = :code; ''', params=data)
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to update data: {e}")
        if self.itemChanged["insert"]:
            update_job(type = "insert", list = insert_list)
            try:
                finish += self.database_process.executemany(sql=''' INSERT INTO `Maintenance_plan` (machine_id,line_id,month_year_id,quarter,week,original_week)
                                                                    SELECT m.machine_id,p.line_id,my.month_year_id,:quarter,:week,:week
                                                                    FROM `Machines` as m
                                                                    JOIN `Production_Lines` as p 
                                                                    ON p.line_name = :line
                                                                    JOIN `Months_Years` as my
                                                                    ON my.month = :new_month AND my.year = :year
                                                                    WHERE m.machine_code = :code; ''', params_list=insert_list)
                for data in insert_list:
                    self.database_process.query(sql = '''UPDATE machines
                                                        SET machine_status = 'GOOD'
                                                        WHERE machine_code = :code;''', params = {'code':data['code']})

            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to update data: {e}")
        try:
            self.database_process.query(sql= "CALL update_maintenance_plan_status;")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to update data: {e}")
            return
        if finish > 0:
            QtWidgets.QMessageBox.information(self, "Success", f"Đã cập nhật {finish} bản ghi.")
            return
    
    @QtCore.pyqtSlot()         
    def open_pdf(self,link):
        pdf = pdf_view(link)
        self.pdf_windows.append(pdf)
        pdf.show()
    
    @QtCore.pyqtSlot() 
    def Delete_plan(self):
        if self.login_info["role_level"] in ["Manager","Admin"]:
                pass
        elif ( self.login_info["department"] == self.department_maintenance_plan ) and ( self.login_info["role_level"] in ["Supervisor"]):
            pass
        else:
            QtWidgets.QMessageBox.information(self,"Permission denied","Your don't have permission to update this machine info")
            return
        current_row = self.ui.Detail_plan_frze_table.currentRow()
        code_item = self.ui.Detail_plan_frze_table.item(current_row, 0)
        if code_item is None:
            self.ui.Detail_plan_table.removeRow(current_row)
            self.ui.Detail_plan_frze_table.removeRow(current_row)
            return
        code = code_item.text()
        question = QtWidgets.QMessageBox.question(self,"Delete",f"Are you sure to delete the maintenance plan for the machine '{code}'?",QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,QtWidgets.QMessageBox.No)
        if question == QtWidgets.QMessageBox.Yes:
            try:
                self.database_process.query(sql = '''   DELETE mp
                                                        FROM `Maintenance_plan` as mp 
                                                        JOIN `Machines` AS m 
                                                        ON mp.machine_id = m.machine_id
                                                        JOIN `Months_Years` AS my
                                                        ON mp.month_year_id = my.month_year_id
                                                        WHERE m.machine_code = :code AND my.year = :year ; ''',params = {'code':code,'year':self.year_num})
                self.ui.Detail_plan_table.removeRow(current_row)
                self.ui.Detail_plan_frze_table.removeRow(current_row)
                self.itemChanged["update"] = {r - 1 if r > current_row else r for r in self.itemChanged["update"] if r != current_row}
                self.itemChanged["insert"] = {r - 1 if r > current_row else r for r in self.itemChanged["insert"] if r != current_row}
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
    
    @QtCore.pyqtSlot() 
    def Insert_plan(self):
        if self.login_info["role_level"] in ["Manager","Admin"]:
                pass
        elif ( self.login_info["department"] == self.department_maintenance_plan ) and ( self.login_info["role_level"] in ["Supervisor"]):
            pass
        else:
            QtWidgets.QMessageBox.information(self,"Permission denied","Your don't have permission to update this machine info")
            return
        row = self.ui.Detail_plan_frze_table.rowCount()
        self.ui.Detail_plan_frze_table.insertRow(row)
        self.ui.Detail_plan_table.insertRow(row)
        editor = QtWidgets.QLineEdit()
        editor.setAlignment(QtCore.Qt.AlignCenter)
        editor.setStyleSheet(''' border: none;''')
        self.safe_connect(editor.textChanged,self.handle_text_changed_DP)
        self.safe_connect(editor.editingFinished, self.handle_editing_finished_DP)
        self.ui.Detail_plan_frze_table.setCellWidget(row,0,editor)
        self.itemChanged["insert"].add(row)
    
    @QtCore.pyqtSlot(str)
    def handle_text_changed_DP(self, text):
        editor = self.sender()
        if editor is None:
            return
        table = self.ui.Detail_plan_frze_table
        for r in range(table.rowCount()):
            if table.cellWidget(r, 0) is editor:
                return self.on_text_changed(text = text, c = 0, r = r, target = self.ui.Detail_plan_frze_table , 
                                                                              select_col = "machine_code" , table = "Machines ", where = f"WHERE machine_code LIKE '%{text}%'")
    
    @QtCore.pyqtSlot()
    def handle_editing_finished_DP(self):
        editor = self.sender()
        if editor is None:
            return
        table = self.ui.Detail_plan_frze_table
        for r in range(table.rowCount()):
            if table.cellWidget(r, 0) is editor:
                    return self.load_machine_detail_plan(r)
    
    @QtCore.pyqtSlot() 
    def load_machine_detail_plan(self,r):
        item = self.ui.Detail_plan_frze_table.cellWidget(r,0)
        code = item.text()
        dep = self.ui.Group_cbb_DP.currentText()
        try:
            isCorrectGroup = self.database_process.query(sql = '''SELECT m.machine_name,d.department_name 
                                                                FROM `Machines` as m
                                                                JOIN `Production_Lines` as p
                                                                ON m.line_id = p.line_id
                                                                JOIN `Departments` as d
                                                                ON p.department_id = d.department_id
                                                                WHERE d.department_name = :dep AND m.machine_code = :code
                                                                GROUP BY m.machine_id ;''', params = {'code':code,'dep':dep})
            if not isCorrectGroup:
                raise Exception("The machine is not in your group")
            isCodeInPlan = self.database_process.query(sql = '''SELECT 1  
                                                                FROM `Machines` as m
                                                                JOIN `Maintenance_plan` as mp
                                                                ON m.machine_id = mp.machine_id
                                                                JOIN `Production_Lines` as p
                                                                ON m.line_id = p.line_id
                                                                JOIN `Departments` as d
                                                                ON p.department_id = d.department_id
                                                                JOIN `Months_Years` AS my
                                                                ON mp.month_year_id = my.month_year_id
                                                                WHERE d.department_name = :dep AND m.machine_code = :code AND my.year = :year
                                                                GROUP BY m.machine_id ;''', params = {'code':code,'dep':dep,'year':self.year_num})
            if isCodeInPlan:
                raise Exception("The machine already have plan")
            self.ui.Detail_plan_frze_table.setItem(r,1,QtWidgets.QTableWidgetItem(isCorrectGroup[0][0]))
            self.ui.Detail_plan_frze_table.setItem(r,2,QtWidgets.QTableWidgetItem(isCorrectGroup[0][1]))
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
            item.clear()

    def send_notification(self, data: dict):
        try:
            clean_data = {k: v for k, v in data.items() if v is not None}

            if 'payload' in clean_data and isinstance(clean_data['payload'], (dict, list)):
                clean_data['payload'] = json.dumps(clean_data['payload'], ensure_ascii=False)

            columns = ', '.join(clean_data.keys())
            placeholders = ','.join([f":{k}" for k in clean_data.keys()])
            sql = f"INSERT INTO `Notifications` ({columns}) VALUES ({placeholders})"

            self.database_process.query(sql=sql, params=clean_data)

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to send notification: {e}")

#==========================Function of Maintenance page ==================================================================================END
#==========================Function of Maintenance page ==================================================================================END
#==========================Function of Maintenance page ==================================================================================END

#==========================Function of Part Order page ==================================================================================BEGIN
#==========================Function of Part Order page ==================================================================================BEGIN
#==========================Function of Part Order page ==================================================================================BEGIN
    
    @QtCore.pyqtSlot()
    def Part_order_page(self):
        self.ui.main_stacked.setCurrentWidget(self.ui.Part_order_page)
        self.set_stylesheet_change_page((self.ui.Order_btn,self.ui.OEE_btn,self.ui.Home_btn,self.ui.Maintenance_btn, self.ui.Stock_btn,self.ui.Downtime_btn))
        if not self.is_expanded:
            self.is_expanded = True
            self.expand_windown_animation(self.is_expanded)

#==========================Function of Part Order page ==================================================================================END
#==========================Function of Part Order page ==================================================================================END
#==========================Function of Part Order page ==================================================================================END

#==========================Function of Stock control page ==================================================================================BEGIN
#==========================Function of Stock control page ==================================================================================BEGIN
#==========================Function of Stock control page ==================================================================================BEGIN

    @QtCore.pyqtSlot()
    def Stock_control_page(self):
        if not self.is_expanded:
            self.is_expanded = True
            self.expand_windown_animation(self.is_expanded)
        self.safe_connect(self.ui.update_inventory_btn.clicked, lambda _:self.run_inventory_update())
        self.ui.main_stacked.setCurrentWidget(self.ui.Stock_control_page)
        self.set_stylesheet_change_page((self.ui.Stock_btn,self.ui.OEE_btn,self.ui.Home_btn,self.ui.Maintenance_btn, self.ui.Order_btn,self.ui.Downtime_btn))
        try:
            if hasattr(self, 'stock_model'):
                return
            if self.ui.group_stock_cbb.count() == 0:
                self.ui.group_stock_cbb.addItem("")
                for group in self.group:
                    self.ui.group_stock_cbb.addItem(group[0])
            else:
                return
            url = "https://open.er-api.com/v6/latest/USD"
            response = requests.get(url)
            self.exchange_rates = response.json()["rates"]
            header = ["Spare part","Safety\nstock", "Current\nstock", "Stock up\nreminder", "Unit\nprice","Total\ncost","Lead\ntime", "Life\ntime","Last\nrequest", "Group","Add PO"]
            self.stock_model = QtGui.QStandardItemModel(0, len(header))
            self.stock_model.setHorizontalHeaderLabels(header)
            self.ui.stock_table.setModel(self.stock_model)
            result = self.database_process.query('''SELECT * FROM `Spare_part_View`
                                                    ORDER BY stockup DESC;''')
            self.inventory_update_date = self.database_process.query('''SELECT MAX(update_at) FROM `inventory`;''')[0][0]
            self.ui.inventory_update_date.setText(str(self.inventory_update_date))
            self.image_files = {
                    item[0] : item[-1]
                    for item in result
                    if item[-1] is not None
                }
            ImageCache.init(self.ui.stock_table)
            delegate = StockItemDelegate(buttons=("+",))
            self.safe_connect( delegate.clicked, self.on_button_clicked_stock)
            self.ui.stock_table.setItemDelegate(delegate)
            self.add_data_to_stock_model(result)
            self.ui.stock_table.setUpdatesEnabled(True)
            self.ui.stock_table.setMouseTracking(True)
            self.ui.stock_table.viewport().update()
            self.ui.stock_table.viewport().setMouseTracking(True)
            header = self.ui.stock_table.horizontalHeader()
            self.ui.stock_table.setColumnWidth(0,600)
            for col in range(1, 11):
                header.setSectionResizeMode(col, QtWidgets.QHeaderView.Stretch)
            self.ui.stock_table.horizontalHeader().setStyleSheet("QHeaderView::section { qproperty-alignment: AlignCenter; }")
            self.ui.stock_table.setSortingEnabled(True)
            self.ui.stock_table.setAlternatingRowColors(True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")
    
    @QtCore.pyqtSlot() 
    def on_button_clicked_stock(self):
        # model = index.model()
        # row = index.row()
        QtWidgets.QMessageBox.information(self, "Info", "Function to add PO is still under development.")
    
    @QtCore.pyqtSlot()
    def show_filter_stock(self):
        self.ui.filter_stock_frame.show()
        self.safe_connect( self.ui.apply_stock_btn.clicked, self.filter_process_stock)
        self.safe_connect( self.ui.cancel_stock_btn.clicked, self.hide_filter_stock)
        self.safe_connect( self.ui.code_stock_lnedit.textChanged, lambda text: self.filter_suggestion_stock(text, "code"))
        self.safe_connect( self.ui.name_stock_lnedit.textChanged, lambda text: self.filter_suggestion_stock(text, "name"))
    
    @QtCore.pyqtSlot()
    def hide_filter_stock(self):
        self.ui.filter_stock_frame.hide()
    
    @QtCore.pyqtSlot() 
    def filter_suggestion_stock(self,text,fill=""):
        if fill == "code":
            if len(text)<3:
                return
            SCRIPT = '''SELECT part_code FROM spare_part_view WHERE part_code LIKE :text LIMIT 5;'''
            target = self.ui.code_stock_lnedit
        else:
            if len(text)<3 :
                return
            SCRIPT = '''SELECT part_name FROM spare_part_view WHERE part_name LIKE :text LIMIT 5;'''
            target = self.ui.name_stock_lnedit
        suggestions = []
        part_code = []
        try:
            part_code = self.database_process.query(SCRIPT, params={"text": f"%{text}%"})
            suggestions = [str(name[0]) for name in part_code]
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to fetch machine names: {e}")
        completer = QtWidgets.QCompleter(suggestions, self)
        completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        completer.setFilterMode(QtCore.Qt.MatchContains)
        target.setCompleter(completer)
        completer.complete()
    
    @QtCore.pyqtSlot() 
    def filter_process_stock(self):
        try:
            query = []
            if self.ui.code_stock_lnedit.text() != "":
                query.append(f'part_code = "{self.ui.code_stock_lnedit.text()}"')
            if self.ui.name_stock_lnedit.text() != "":
                query.append(f'part_name = "{self.ui.name_stock_lnedit.text()}"')
            if self.ui.group_stock_cbb.currentText() != "":
                query.append(f'department_name = "{self.ui.group_stock_cbb.currentText()}"')
            query = " AND ".join(query)
            if query == "":
                result = self.database_process.query(sql='''SELECT * FROM `spare_part_view`
                                                            ORDER BY stockup DESC;''')

                self.add_data_to_stock_model(result)
                self.hide_filter_stock() 
                return
            final_query = f'''  SELECT *
                                FROM `spare_part_view`
                                WHERE {query}  ORDER BY stockup DESC;'''
            result = self.database_process.query(sql=final_query)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to filter data: {e}")
            return
        self.add_data_to_stock_model(result)
        self.hide_filter_stock()
    
    @QtCore.pyqtSlot()
    def reset_filter_stock(self):
        self.ui.code_stock_lnedit.clear()
        self.ui.name_stock_lnedit.clear()
        self.ui.group_stock_cbb.setCurrentIndex(0)
        self.filter_process_stock()
    
    def make_item(self,text, align=QtCore.Qt.AlignCenter):
        value = round(text, 0) if isinstance(text, float) and text == 0 else text
        item = QtGui.QStandardItem(str(value))
        item.setTextAlignment(align)
        return item

    def add_data_to_stock_model(self, result):
        self.stock_model.removeRows(0, self.stock_model.rowCount())
        cost = 0
        value = 0
        code = 0
        for row in range(len(result)):
            image_path = self.image_files.get(result[row][0], None)
            data = {"image":image_path,
                    "name": f"{result[row][1]}", "code": f"{result[row][0]}"}
            spare_part = QtGui.QStandardItem()
            spare_part.setData(data, QtCore.Qt.UserRole)
            if (float(result[row][4]) > 0):
                code += 1
            value += float(result[row][4])
            cost += float(result[row][6]) / float(self.exchange_rates[result[row][11]])
            row_items = [
                spare_part,
                self.make_item(float(result[row][2])), # safety stock
                self.make_item(float(result[row][3])),  # current stock         
                self.make_item(float(result[row][4])),  # stock up   
                self.make_item(round(float( result[row][5]) / float(self.exchange_rates[result[row][11]]),2)),  # unit cost   
                self.make_item(round( float(result[row][6]) / float(self.exchange_rates[result[row][11]]),2)),  # total cost                
                self.make_item(result[row][7]), # lead time             
                self.make_item(result[row][8]), # life time  
                self.make_item(result[row][9]), # last request    
                self.make_item(result[row][10]) # department name 
            ]
            self.stock_model.appendRow(row_items)
        self.ui.total_part_num.setText(f"{len(result)}")
        self.ui.code_need_order_num.setText(f"{int(code)}")
        self.ui.quantity_order_num.setText(f"{int(value)}")
        self.ui.total_cost_num.setText(f"${round(cost,2)}")
        self.ui.stock_table.resizeRowsToContents()

    def call_inventory_update(self):
        url = os.getenv("API_UPDATE_INVENTORY")
        try:
            response = requests.get(url, timeout= 180)
            return response.json()
        except Exception as e:
            return {"status": "error", "detail": str(e)}
    
    @QtCore.pyqtSlot()
    def run_inventory_update(self):
        QtWidgets.QMessageBox.information(self, "Info", "Inventory update on the client side is under development, and can be auto-updated from the server side for now.")
        return
        try: 
            isLocked = self.database_process.query('''SELECT GET_LOCK('inventory_update', 3);''')[0][0]
            if isLocked != 1:
                return QtWidgets.QMessageBox.information(  self,
                                                            "In Progress",
                                                            "Another update session is running. Try again after 2 minutes."
                                                        )
            last_update = self.database_process.query('''SELECT MAX(update_at) FROM `inventory`;''')
            last_update = last_update[0][0]
            now = dt.datetime.now()
            diff = now - last_update
            minutes = diff.total_seconds() / 60
            if last_update != self.inventory_update_date and minutes > 3 :
                return self._start_inventory_update()

            if last_update == self.inventory_update_date and minutes > 3 :
                return self._start_inventory_update()
            
            if last_update != self.inventory_update_date:
                return self.Stock_control_page()

            return QtWidgets.QMessageBox.information(
                self,
                "Up to date",
                "Has been updated to the latest value."
            )
        except Exception as e:
            self.setEnabled(True)
            self.spinner.stop()
            self.database_process.query('''SELECT RELEASE_LOCK('inventory_update');''')
            QtWidgets.QMessageBox.warning(
                self, "Error", f"Inventory update failed: {str(e)}"
            )
    
    def _start_inventory_update(self):
        self.setEnabled(False)
        self.spinner.start()

        worker = Worker_Pool(self.call_inventory_update)
        worker.signals.finished.connect(self.on_inventory_update_done)
        worker.signals.error.connect(self.on_inventory_update_error)
        QtCore.QThreadPool.globalInstance().start(worker)

    @QtCore.pyqtSlot(object)
    def on_inventory_update_done(self, data):
        self.spinner.stop()
        self.setEnabled(True)
        if data.get("status") == "finish":
            QtWidgets.QMessageBox.information(
                self, "Finished", "Inventory update completed.")
            self.Stock_control_page()
        else:
            QtWidgets.QMessageBox.warning(
                self, "Error",f"Error in update Inventory process: {str(data)}")
        self.database_process.query('''SELECT RELEASE_LOCK('inventory_update');''')
    
    @QtCore.pyqtSlot()
    def on_inventory_update_error(self, err):
        self.spinner.stop()
        self.setEnabled(True)
        QtWidgets.QMessageBox.warning(
                self,"Error",f"Inventory update failed: {str(err)}")
        self.database_process.query('''SELECT RELEASE_LOCK('inventory_update');''')
#==========================Function of Stock control page ==================================================================================END
#==========================Function of Stock control page ==================================================================================END
#==========================Function of Stock control page ==================================================================================END


#==========================Function of OEE page ====================================================================================
#==========================Function of OEE page ====================================================================================
#==========================Function of OEE page ====================================================================================
    def Show_windown(self, windown_attr, ui_Window, data=None) : 
        if getattr(self, windown_attr) is not None:
            getattr(self, windown_attr).close()
            getattr(self, windown_attr).deleteLater()
        setattr(self, windown_attr, QtWidgets.QMainWindow())
        UI_windown = ui_Window()
        if data is None:
            UI_windown.setupUi(getattr(self, windown_attr))
        else:
            UI_windown.setupUi(getattr(self, windown_attr), data)
        getattr(self, windown_attr).setWindowFlags(
            getattr(self, windown_attr).windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
        getattr(self, windown_attr).show()
        getattr(self, windown_attr).raise_()
        getattr(self, windown_attr).activateWindow()
        return

    def textbox_read(self, textbox):
        file_path = textbox.toPlainText().strip()
        file_path = file_path.replace("file:///", "")
        file_path = file_path.replace("file:", "")

        if not file_path:
            QtWidgets.QMessageBox.warning(
                self, "Thiếu thông tin", "Vui lòng nhập đường dẫn file Excel.")
            return

        if not os.path.exists(file_path):
            QtWidgets.QMessageBox.critical(
                self, "Lỗi đường dẫn", "Đường dẫn file không tồn tại.")
            return

        try:
            self.df = pd.read_excel(file_path)
        except Exception as e:
            error_message = f"Không thể đọc file. Vui lòng kiểm tra định dạng file.\nChi tiết: {e}"
            QtWidgets.QMessageBox.critical(self, "Lỗi đọc file", error_message)
            return
        return file_path
    
    @QtCore.pyqtSlot()
    def handle_df_show(self, textbox):
        self.textbox_read(textbox)
        if self.df is None:
            return
        self.Show_windown('df_windown', df_show, self.df)
        return
    
    @QtCore.pyqtSlot()
    def Open_Setting_windown(self):
        if self.Setting_windown is not None:
            self.Setting_windown.close()
            self.Setting_windown.deleteLater()
        self.Setting_windown = Setting_windown(parent=self,  model=self.model)
        self.Setting_windown.show()
        return
    
    @QtCore.pyqtSlot()
    def Open_View_result_windown(self):
        if self.View_result_windown is not None:
            self.View_result_windown.close()
            self.View_result_windown.deleteLater()
        self.View_result_windown = View_result_windown(parent=self)
        self.View_result_windown.show()
        return
    
    @QtCore.pyqtSlot()
    def Data_process(self):
        self.link_NG = self.textbox_read(self.ui.Text_NG)
        self.link_FG = self.textbox_read(self.ui.Text_FG)
        # self.link_FG = r"Excel_data\Molding daily Mar-25.xlsx"
        # self.link_NG = r"Excel_data\Tong hop so luong NG OK 3-25.xls"
        # self.link_FG = r"C:\Users\2173452100291\Documents\Excel_data\Molding daily Mar-25.xlsx"
        # self.link_NG = r"C:\Users\2173452100291\Documents\Excel_data\Tong hop so luong NG OK 3-25.xls"
        self.list_df_molding_result, self.list_df_molding_monthly_result, self.list_df_coil_result, self.list_df_coil_monthly_result = None, None, None, None
        if self.link_NG is None or self.link_FG is None:
            return
        try:
            self.Flag_data_process = True
            self.list_df_molding_result, self.list_df_molding_monthly_result, self.list_df_coil_result, self.list_df_coil_monthly_result = self.model.OEE_cal_result(
                FG_file=self.link_FG, NG_file=self.link_NG)
            QtWidgets.QMessageBox.information(
                self, "Hoàn tất", "Xử lý dữ liệu thành công")
        except Exception as e:
            error_message = f"Không thể xử lý dữ liệu. Vui lòng kiểm tra lại.\nChi tiết: {e}"
            QtWidgets.QMessageBox.critical(
                self, "Lỗi xử lý dữ liệu", error_message)
            return

    def Export_Excel(self, file_export=None, machine=None):
        if self.Flag_data_process is False:
            QtWidgets.QMessageBox.warning(
                self, "Chưa xử lý dữ liệu", "Vui lòng xử lý dữ liệu trước khi xuất file Excel.")
            return
        if machine == "Molding":
            self.export_list = self.list_df_molding_result
        elif machine == "Coil":
            self.export_list = self.list_df_coil_result
        save_path = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Chọn thư mục lưu file", "")
        if not save_path:
            QtWidgets.QMessageBox.warning(
                self, "Thiếu thông tin", "Vui lòng chọn thư mục để lưu file.")
            return

        try:
            if file_export == None:
                for line, df in self.export_list:
                    df["Day"] = df.index
                    df.to_excel(os.path.join(
                        save_path, f"OEE_{machine}_result_{line}.xlsx"), index=False)
            else:
                self.file_export = file_export
                for line, df in self.export_list:
                    if line in self.file_export:
                        df["Day"] = df.index
                        df.to_excel(os.path.join(
                            save_path, f"OEE_{machine}_result_{line}.xlsx"), index=False)
            QtWidgets.QMessageBox.information(
                self, "Thành công", "File Excel đã được lưu thành công.")
            self.Flag_data_process = True
        except Exception as e:
            error_message = f"Không thể lưu file. Vui lòng kiểm tra lại.\nChi tiết: {e}"
            QtWidgets.QMessageBox.critical(self, "Lỗi lưu file", error_message)
    
    @QtCore.pyqtSlot()
    def Choose_export_machine(self, file_export, where=None):
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Choose Export Option")
        dialog.setMinimumSize(300, 200)
        layout = QtWidgets.QVBoxLayout(dialog)
        label = QtWidgets.QLabel("Select the type of machine to export:")
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(16)
        label.setFont(font)
        layout.addWidget(label)

        button_molding = QtWidgets.QPushButton("Molding")
        button_molding.setMinimumSize(200, 50)
        button_molding.setFont(font)

        button_coil = QtWidgets.QPushButton("Coil")
        button_coil.setMinimumSize(200, 50)
        button_coil.setFont(font)

        layout.addWidget(button_molding)
        layout.addWidget(button_coil)
        self.flag_return = False
        if where == "Excel":
            button_molding.clicked.connect(
                lambda: self.return_choose(type="Excel", machine="Molding", file_export=file_export))
            button_coil.clicked.connect(
                lambda: self.return_choose(type="Excel", machine="Coil", file_export=file_export))
            dialog.exec_()
        elif where == "DB":
            button_molding.clicked.connect(
                lambda: self.return_choose(type="DB", machine="Molding", file_export=file_export))
            button_coil.clicked.connect(
                lambda: self.return_choose(type="DB", machine="Coil", file_export=file_export))
            dialog.show()
            while not self.flag_return:
                QtWidgets.QApplication.processEvents()
            dialog.close()
    
    @QtCore.pyqtSlot() 
    def return_choose(self, type=None, machine=None, file_export=None):
        if type == "Excel":
            return self.Export_Excel(file_export=file_export, machine=machine)
        elif type == "DB":
            self.machine_type_for_db = "Molding" if machine == "Molding" else "Coil"
            self.flag_return = True
            return

    def closeEvent(self, event):
        if hasattr(self, "database_process") and self.database_process:
            try:
                self.database_process.close()
            except Exception as e:
                print("Error closing database_process:", e)

        event.accept() 
        QtWidgets.QApplication.quit()

    def draw_circle(self,widget, x, y, r, color=()):
        pixmap = QtGui.QPixmap(widget.size())       # tạo pixmap có cùng size widget
        pixmap.fill(QtCore.Qt.transparent)          # nền trong suốt

        painter = QtGui.QPainter(pixmap)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        # Gradient hình nón (conical), tâm ở giữa hình tròn
        gradient = QtGui.QConicalGradient(x + r/2, y + r/2,127)  
        
        gradient.setColorAt(0.0, QtGui.QColor(color[0], color[1], color[2], 20))    
        gradient.setColorAt(0.25, QtGui.QColor(color[0], color[1], color[2], 127))  
        gradient.setColorAt(0.5, QtGui.QColor(color[0], color[1], color[2], 255))    
        gradient.setColorAt(0.75, QtGui.QColor(color[0], color[1], color[2], 127)) 
        gradient.setColorAt(1.0, QtGui.QColor(color[0], color[1], color[2], 20))   

        # Gán gradient làm bút vẽ
        pen = QtGui.QPen(QtGui.QBrush(gradient), 20)  # brush chứa gradient, độ dày 20
        painter.setPen(pen)

        painter.drawEllipse(int(x), int(y), int(r), int(r))
        painter.end()

        widget.setPixmap(pixmap)  

    def style_button_with_shadow(self,button:tuple):
        button[0].setStyleSheet('''
                                    QPushButton {
                                                background-color: rgba(0, 0, 0, 0.08);
                                                border: none;
                                                border-bottom: 2px solid rgba(0, 0, 255, 1);
                                                padding: 5px 15px;
                                                font-weight: bold;
                                    }
        ''')
        for i in range(1,len(button)):
            button[i].setStyleSheet('''
                                    QPushButton {
                                                    background-color: transparent;
                                                    border: none;
                                                    border-radius: 0px;
                                                    padding: 5px 15px;
                                                    }
                                    QPushButton:hover {
                                                        background-color: rgba(0, 0, 0, 0.15);
                                                        border-bottom: 1px solid rgba(0, 0, 255, 1);
                                                        padding: 5px 15px;
                                                        }
                                    ''')

#==========================================================================================================================


#==========================================================================================================================


#==========================================================================================================================
    @QtCore.pyqtSlot()
    def Downtime_page(self):
        self.ui.main_stacked.setCurrentWidget(self.ui.Downtime_page)
        self.set_stylesheet_change_page((self.ui.Downtime_btn,self.ui.OEE_btn,self.ui.Home_btn,self.ui.Maintenance_btn, self.ui.Stock_btn, self.ui.Order_btn))
        if not self.is_expanded:
            self.is_expanded = True
            self.expand_windown_animation(self.is_expanded)
        self.Dashboard_Downtime_page()
        self.safe_connect(self.ui.DT_dashboard_btn.clicked, self.Dashboard_Downtime_page)
        self.safe_connect(self.ui.DT_data_btn.clicked, self.Data_Downtime_page)
        self.safe_connect(self.ui.DT_import_data_btn.clicked, self.Import_data_Downtime_page)
        self.safe_connect(self.ui.DT_problem_report_btn.clicked, self.Problem_report_Downtime_page)
    
    @QtCore.pyqtSlot()
    def Dashboard_Downtime_page(self):
        self.style_button_with_shadow((self.ui.DT_dashboard_btn,self.ui.DT_data_btn,self.ui.DT_import_data_btn,self.ui.DT_problem_report_btn))
        self.ui.DT_stacked_widget.setCurrentWidget(self.ui.DT_Dashboard_widget)
        try:
            areas = [area[0] for area in self.database_process.query(sql = '''SELECT downtime_area_name
                                                                                FROM `downtime_areas`;''')]
            self.ui.DT_area_cbb.clear()
            # self.ui.DT_area_cbb.addItem("")
            self.ui.DT_area_cbb.addItems(areas)
            self.DT_silde_bar_animation = QtCore.QPropertyAnimation(self.ui.frame_106, b"maximumWidth")
            self.DT_silde_bar_animation.setDuration(310)
            self.safe_connect(self.ui.DT_show_sorting_btn.clicked, self.show_sorting_filter_Downtime)
            area_name = self.ui.DT_area_cbb.currentText()
            nearest_date = self.database_process.query(sql = '''SELECT MAX(downtime_date) FROM `downtime_records`;''')[0][0]
            self.ui.DT_date_edit_2.setDate(QtCore.QDate.fromString(str(nearest_date), "yyyy-MM-dd"))
            self.Dashboard_Downtime_page_refresh(area_name, nearest_date)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
            return

    def Dashboard_Downtime_page_refresh(self, area_name, date):
        try:
            data = self.database_process.query(sql = '''SELECT d.downtime_start_time, d.downtime_start_repair_time, d.downtime_end_time, 
                                                            TIMESTAMPDIFF
                                                            (
                                                                MINUTE,
                                                                CONCAT(d.downtime_date,' ', d.downtime_start_time),
                                                                CASE 
                                                                    WHEN d.downtime_end_time < d.downtime_start_time 
                                                                    THEN CONCAT(DATE_ADD(d.downtime_date, INTERVAL 1 DAY),' ', d.downtime_end_time)
                                                                    ELSE CONCAT(d.downtime_date,' ', d.downtime_end_time)
                                                            END  ) AS total_loss,
                                                            TIMESTAMPDIFF
                                                                (
                                                                    MINUTE,
                                                                    CONCAT(d.downtime_date,' ', d.downtime_start_repair_time),
                                                                    CASE 
                                                                        WHEN d.downtime_end_time < d.downtime_start_repair_time 
                                                                        THEN CONCAT(DATE_ADD(d.downtime_date, INTERVAL 1 DAY),' ', d.downtime_end_time)
                                                                        ELSE CONCAT(d.downtime_date,' ', d.downtime_end_time)
                                                                END  ) AS wait_technical,
                                                                d.staff_name, d.error_code, m.machine_code, p.line_name
                                                        FROM `downtime_records` d
                                                        JOIN `machines` m ON d.machine_id = m.machine_id
                                                        JOIN `production_lines` p ON m.line_id = p.line_id
                                                        JOIN `downtime_areas_production_lines` dapl ON dapl.line_id = p.line_id
                                                        JOIN `downtime_areas` da ON dapl.downtime_area_id = da.downtime_area_id
                                                        WHERE da.downtime_area_name = :area_name AND d.downtime_date = :downtime_date
                                                        ORDER BY d.downtime_start_time,d.downtime_date;''', params = {"area_name": area_name,"downtime_date":date})
            working_time = self.database_process.query(sql = '''SELECT SUM(operation_hours) FROM `line_operation_times` as lot
                                                                JOIN downtime_areas_production_lines as dapl ON lot.line_id = dapl.line_id
                                                                JOIN downtime_areas as da ON dapl.downtime_area_id = da.downtime_area_id
                                                                WHERE da.downtime_area_name = :area_name AND lot.operation_date = :operation_date''', params={"area_name": area_name, "operation_date": date})[0][0]
            if not data:
                QtWidgets.QMessageBox.information(self, "No data", "No downtime records found for the selected area and date.")
                return
            self.data = pd.DataFrame(data, columns=["Downtime Start Time", "Downtime Start Repair Time", "Downtime End Time", "Total Loss Time", "Wait Technical Time", "Staff Name", "Error Code", "Machine Code", "Line Name"])
            total_loss = self.data["Total Loss Time"].sum()
            downtime_count = len(self.data)
            mttr = self.data["Total Loss Time"].mean() if downtime_count > 0 else 0
            mttr = str(dt.timedelta(minutes=int(mttr)))
            mtbf = (working_time * 60 - total_loss) / downtime_count if downtime_count > 0 else working_time * 60
            mtbf = str(dt.timedelta(minutes=int(mtbf)))
            total_loss = str(dt.timedelta(minutes=int(total_loss)))
            self.ui.DTime_value.setText(str(total_loss))
            self.ui.DEvent_value.setText(str(downtime_count))
            self.ui.MTTR_value.setText(str(mttr))
            self.ui.MTBF_value.setText(str(mtbf))
            DE_perhours = pd.DataFrame(columns=["Date_time"])
            DE_perhours["Date_time"] = pd.to_datetime(date) + self.data["Downtime Start Time"]
            DE_perhours["Date_time"] = DE_perhours["Date_time"].dt.floor("h")
            full_hours = pd.date_range(start=pd.to_datetime(date),periods=24,freq="h")
            Event_hourly_df = (
                DE_perhours.groupby("Date_time")
                .size()
                .reset_index(name="count_event")
            )
            Event_hourly_df = (
                Event_hourly_df
                .set_index("Date_time")
                .reindex(full_hours, fill_value=0)
                .rename_axis("Date_time")
                .reset_index()
            )
            MTTR_perhours = pd.DataFrame(columns=["Date_time", "MTTR"])
            MTTR_perhours["Date_time"] = pd.to_datetime(date) + self.data["Downtime Start Time"]
            MTTR_perhours["MTTR"] = self.data["Total Loss Time"]
            MTTR_perhours["Date_time"] = MTTR_perhours["Date_time"].dt.floor("h")
            MTTR_perhours = (
                MTTR_perhours
                .groupby("Date_time")["MTTR"]
                .mean()
                .reindex(full_hours, fill_value=0)
                .rename_axis("Date_time")
                .reset_index()
            )
            MTBF_perhours = pd.DataFrame(columns=["Date_time", "MTBF"])
            MTBF_perhours["Date_time"] = pd.to_datetime(date) + self.data["Downtime Start Time"]
            MTBF_perhours["MTBF"] =  MTBF_perhours["Date_time"].diff().dt.total_seconds().div(60).fillna(0)
            MTBF_perhours["Date_time"] = MTBF_perhours["Date_time"].dt.floor("h")
            MTBF_perhours = (
                MTBF_perhours
                .dropna(subset=["MTBF"])
                .groupby("Date_time")["MTBF"]
                .mean()
                .reindex(full_hours, fill_value=0)
                .rename_axis("Date_time")
                .reset_index()
            )
            self.Sparkline_chart(self.ui.DTime_chart, self.data["Total Loss Time"].tolist(), (165, 201, 229), "Downtime vs Time")
            self.Sparkline_chart(self.ui.DEvent_chart, Event_hourly_df["count_event"].tolist(), (165, 201, 229), "Downtime Events vs Time")
            self.Sparkline_chart(self.ui.MTTR_chart, MTTR_perhours["MTTR"].tolist(), (165, 201, 229), "MTTR vs Time")
            self.Sparkline_chart(self.ui.MTBF_chart, MTBF_perhours["MTBF"].tolist(), (165, 201, 229), "MTBF vs Time")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
            return
        
    @QtCore.pyqtSlot()
    def show_sorting_filter_Downtime(self):
        if self.ui.DT_show_sorting_btn.isChecked():
            self.ui.frame_106.setEnabled(True)
            self.DT_silde_bar_animation.setStartValue(0)
            self.DT_silde_bar_animation.setEndValue(535)
        else:
            self.ui.frame_106.setEnabled(False)
            self.DT_silde_bar_animation.setStartValue(535)
            self.DT_silde_bar_animation.setEndValue(0)
        self.DT_silde_bar_animation.start()

    def Sparkline_chart(self, widget,data, color,title=""):
        old_layout = widget.layout()
        if old_layout is not None:
            while old_layout.count():
                item = old_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
        else:
            new_layout = QtWidgets.QVBoxLayout()
            new_layout.setContentsMargins(0, 0, 0, 30)
            new_layout.setSpacing(0)
            widget.setLayout(new_layout)
        plot = pg.PlotWidget()
        plot.setFixedSize(140, 100)
        plot.setBackground(None)
        plot.hideAxis('left')
        plot.hideAxis('bottom')
        plot.setTitle(f'<span style="color: grey; font-size: 8pt">{title}</span>')
        plot.setMouseEnabled(x=False, y=False)
        plot.setMenuEnabled(False)
        x = np.arange(len(data))
        y = np.array(data, dtype=float)
        curve = pg.PlotCurveItem(x, y, pen=pg.mkPen(color, width=1.5))
        fill = pg.FillBetweenItem(
            curve,
            pg.PlotCurveItem(x, np.zeros_like(y)),
            brush=pg.mkBrush(116, 185, 232, 80)
        )
        plot.addItem(curve)
        plot.addItem(fill)
        widget.setMaximumSize(150, 150)
        widget.layout().addWidget(plot)
        

    @QtCore.pyqtSlot()
    def Data_Downtime_page(self):
        self.style_button_with_shadow((self.ui.DT_data_btn,self.ui.DT_import_data_btn,self.ui.DT_problem_report_btn,self.ui.DT_dashboard_btn))
        self.ui.DT_stacked_widget.setCurrentWidget(self.ui.DT_detail_data_page)
    
    @QtCore.pyqtSlot()
    def Import_data_Downtime_page(self):
        self.style_button_with_shadow((self.ui.DT_import_data_btn,self.ui.DT_problem_report_btn,self.ui.DT_dashboard_btn,self.ui.DT_data_btn))
        self.ui.DT_stacked_widget.setCurrentWidget(self.ui.DT_import_data_page)
        self.style_button_with_shadow((self.ui.DT_error_chart_btn,self.ui.DT_line_chart_btn,self.ui.DT_machine_chart_btn))
        headers = ["Date", "Line", "Start\nTime","Technical\nStart","Finish\nTime","Total Loss\nTime","Wait\nTechnical","Technical\nName", "Failure\nCode", "Machine Code"]
        self.DT_model = QtGui.QStandardItemModel(0, len(headers))
        self.DT_model.setHorizontalHeaderLabels(headers)
        self.ui.DT_data_table.setModel(self.DT_model)
        self.ui.DT_data_table.setColumnWidth(0,100)
        vetical_header = ["Area", "Total Failure", "Total Loss", "MTTR", "MTBF","Machine with Most Failure", "Failure Code Most Frequent"]
        self.DT_summary_model = QtGui.QStandardItemModel(len(vetical_header), 2)
        self.ui.DT_summary_table.setUpdatesEnabled(False)
        self.ui.DT_summary_table.setSortingEnabled(False)
        for i in range(len(vetical_header)):
            item = QtGui.QStandardItem(vetical_header[i])
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.DT_summary_model.setItem(i, 0, item)
        self.ui.DT_summary_table.setModel(self.DT_summary_model)
        self.ui.DT_summary_table.horizontalScrollBar().setVisible(False)
        self.ui.DT_summary_table.setColumnWidth(0,180)
        self.ui.DT_summary_table.setColumnWidth(1,self.ui.DT_summary_table.width()-179)
        self.ui.DT_summary_table.setUpdatesEnabled(True)
        self.ui.DT_summary_table.setSortingEnabled(True)
        self.safe_connect(self.ui.DT_upload_data_btn.clicked, self.DT_excel_upload)
        self.safe_connect(self.ui.DT_error_code_btn.clicked, self.DT_error_code_show)

    @QtCore.pyqtSlot()
    def DT_excel_upload(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        def get_area_name():
            try:
                group_choose = Group_Area_Choose(parent=self,database=self.database_process, file_path=file_path)
                if group_choose.exec() == QtWidgets.QDialog.Accepted:
                    return group_choose.selected_area,group_choose.excel_sheet_name
                return None, None
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to choose area: {e}")
                return None
        if file_path:
            try:
                area_name, sheet_name = get_area_name()
                if area_name is None or sheet_name is None:
                    return
                excel_data_process = Downtime_Excel_Processor(file_path=file_path, sheet_name=sheet_name, area_name=area_name, database=self.database_process)
                self.data, self.error_frame, self.working_time = excel_data_process.read_filter_excel()
                if self.data is not None:
                    downtime_input_dialog = Downtime_Input(parent=self, database=self.database_process, data_frame=self.data, error_frame=self.error_frame, area_name=area_name, month_year=excel_data_process.month_year)
                    downtime_input_dialog.exec()
                if downtime_input_dialog.result() == QtWidgets.QDialog.Accepted:
                    self.ui.DT_data_table.setUpdatesEnabled(False)
                    self.ui.DT_data_table.setSortingEnabled(False)
                    self.DT_model.removeRows(0, self.DT_model.rowCount())
                    self.DT_model.setRowCount(len(downtime_input_dialog.data_frame))
                    self.DT_data = downtime_input_dialog.data_frame
                    self.DT_month_year = downtime_input_dialog.month_year
                    for r in range(len(self.DT_data)):
                        for c in range(len(self.DT_data.columns)):
                            item = QtGui.QStandardItem(str(self.DT_data.iat[r, c]))
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                            self.DT_model.setItem(r, c, item)
                    self.ui.DT_data_table.setUpdatesEnabled(True)
                    self.ui.DT_data_table.setSortingEnabled(True)
                    self.DT_summary_table_show(area_name, self.DT_data, self.working_time)
                    self.DT_summary_chart_show("error", self.DT_data)
                    self.safe_connect(self.ui.DT_error_chart_btn.clicked, lambda: self.DT_summary_chart_show("error", self.DT_data))
                    self.safe_connect(self.ui.DT_line_chart_btn.clicked, lambda: self.DT_summary_chart_show("line", self.DT_data))
                    self.safe_connect(self.ui.DT_machine_chart_btn.clicked, lambda: self.DT_summary_chart_show("machine", self.DT_data))
                    self.safe_connect(self.ui.DT_import_database_btn.clicked, lambda: self.DT_import_database(self.working_time))
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to upload data: {e}")

    def DT_summary_table_show(self,area,data_frame, working_time):
        try:
            area_lbl = area
            total_failure = data_frame.shape[0]
            total_loss = data_frame["total_loss_time"].sum()
            total_working_time = working_time.drop(columns=["Date"]).apply(pd.to_numeric, errors="coerce").fillna(0).to_numpy().sum()
            mttr = round(total_loss/total_failure,2) if total_failure > 0 else 0
            mtbf = round((total_working_time*60 - total_loss)/total_failure,2) if total_failure > 0 else 0
            machine_most_failure = data_frame["machine_code"].mode()[0] if not data_frame["machine_code"].mode().empty else "N/A"
            failure_code_most_frequent = data_frame["error_code"].mode()[0] if not data_frame["error_code"].mode().empty else "N/A"
            summary_data = [area_lbl, f"{total_failure} times", f"{total_loss} mins", f"{mttr} mins", f"{mtbf} mins", machine_most_failure, failure_code_most_frequent]
            self.ui.DT_summary_table.setUpdatesEnabled(False)
            self.ui.DT_summary_table.setSortingEnabled(False)
            for i in range(len(summary_data)):
                item = QtGui.QStandardItem(str(summary_data[i]))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.DT_summary_model.setItem(i, 1, item)
            self.ui.DT_summary_table.setUpdatesEnabled(True)
            self.ui.DT_summary_table.setSortingEnabled(True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to calculate summary: {e}")
    
    @QtCore.pyqtSlot()
    def DT_summary_chart_show(self, chart_type, data_frame):
        layout = self.ui.DT_summary_chart_widget.layout()
        if layout is not None:
            while layout.count():
                child = layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()
        else:
            layout = QtWidgets.QVBoxLayout(self.ui.DT_summary_chart_widget)
            self.ui.DT_summary_chart_widget.setLayout(layout)
        
        try:
            if chart_type == "error":
                self.style_button_with_shadow((self.ui.DT_error_chart_btn,self.ui.DT_line_chart_btn,self.ui.DT_machine_chart_btn))
                error_counts = data_frame["error_code"].value_counts()
                self.DT_chart_drawing(data_frame, "error_code", "total_loss_time", "Top 10 Error Codes by Loss Time", "Error Code", "Total Loss Time (mins)")
            elif chart_type == "line":
                self.style_button_with_shadow((self.ui.DT_line_chart_btn,self.ui.DT_machine_chart_btn,self.ui.DT_error_chart_btn))
                line_counts = data_frame["line"].value_counts()
                self.DT_chart_drawing(data_frame, "line", "total_loss_time", "Top 10 Lines by Loss Time", "Line", "Total Loss Time (mins)")
            elif chart_type == "machine":
                self.style_button_with_shadow((self.ui.DT_machine_chart_btn,self.ui.DT_error_chart_btn,self.ui.DT_line_chart_btn))
                machine_counts = data_frame["machine_code"].value_counts()
                self.DT_chart_drawing(data_frame, "machine_code", "total_loss_time", "Top 10 Machines by Loss Time", "Machine", "Total Loss Time (mins)")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to show chart: {e}")

    def DT_chart_drawing(self,data_frame, group_by_col, value_col, title, x_label, y_label):
        layout = self.ui.DT_summary_chart_widget.layout()
        if layout is not None:
            while layout.count():
                child = layout.takeAt(0) 
                if child.widget():
                    child.widget().deleteLater()
            layout.setContentsMargins(0, 0, 0, 0)
        else:
            layout = QtWidgets.QVBoxLayout(self.ui.DT_summary_chart_widget)
            layout.setContentsMargins(0, 0, 0, 0)
            self.ui.DT_summary_chart_widget.setLayout(layout)

        try:
            error_loss_time = data_frame.groupby(group_by_col)[value_col].sum().sort_values(ascending=False).head(10)
            widget_w = self.ui.DT_summary_chart_widget.width() - 10
            widget_h = self.ui.DT_summary_chart_widget.height() - 10
            dpi = 100
            fig_w = widget_w / dpi
            fig_h = widget_h / dpi
            
            fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=dpi)
            fig.patch.set_alpha(0.0)
            ax.set_facecolor("none")

            bars = ax.bar(x=range(len(error_loss_time)), height=error_loss_time.values, color='#3FDAA7')
            
            ax.set_xticks(range(len(error_loss_time)))
            ax.set_xticklabels(error_loss_time.index, rotation=90, ha='center',va='top', fontsize=7)
            max_val = int(error_loss_time.values.max())
            step = max(10, int(max_val // 5))
            ax.yaxis.set_visible(False)
            ax.set_ylabel(y_label, fontsize=8)
            ax.set_xlabel(x_label, fontsize=8)
            ax.set_title(title, fontsize=9, fontweight='bold')

            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_visible(False)
            for bar in bars:
                height = bar.get_height()
                if height > 0:
                    ax.text(bar.get_x() + bar.get_width()/2., height,
                            f'{int(height)}', ha='center', va='bottom', fontsize=6)
            
            fig.tight_layout()
            
            canvas = FigureCanvas(fig)
            canvas.setFixedSize(widget_w, widget_h)
            layout.addWidget(canvas)
            canvas.draw()
            plt.close(fig)
            
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to show chart: {e}")

    @QtCore.pyqtSlot()
    def DT_error_code_show(self):
        try:
            error_code_dialog = Error_code_management(parent=self, database=self.database_process)
            error_code_dialog.exec_()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to show error codes: {e}")

    @QtCore.pyqtSlot()
    def DT_import_database(self,working_time):
        if self.DT_data is None:
            QtWidgets.QMessageBox.warning(self, "No Data", "Please upload and review the data before importing to database.")
            return
        try:
            working_time_reframe = working_time.melt(
                        id_vars=["Date"],
                        var_name="Line",
                        value_name="Working Time"
                    )
            working_time_reframe = working_time_reframe[working_time_reframe["Working Time"] > 0]
            import_data_list = [
                        {
                            "machine_code": row.iloc[9],
                            "line_name": row.iloc[1],
                            "downtime_date": f"{self.DT_month_year}-{int(row.iloc[0]):02d}",
                            "downtime_start_time": row.iloc[2],
                            "downtime_start_repair_time": row.iloc[3],
                            "downtime_end_time": row.iloc[4],
                            "staff_name": row.iloc[7],
                            "error_code": row.iloc[8],
                        }
                        for _, row in self.DT_data.iterrows()
            ]
            working_time_import_list = [
                    {
                        "line_name": row["Line"],
                        "operation_date": row["Date"],
                        "operation_hours": row["Working Time"]
                    }
                    for _, row in working_time_reframe.iterrows()
                ]
            sql = '''INSERT INTO `downtime_records`
                        (`machine_id`, `line_id`, `downtime_date`, `downtime_start_time`,
                        `downtime_start_repair_time`, `downtime_end_time`, `staff_name`, `error_code`)
                    SELECT
                        (SELECT machine_id FROM `machines` WHERE machine_code = :machine_code),
                        (SELECT line_id FROM `production_lines` WHERE line_name = :line_name),
                        :downtime_date,
                        :downtime_start_time,
                        :downtime_start_repair_time,
                        :downtime_end_time,
                        :staff_name,
                        :error_code
                '''
            sql_working_time = '''INSERT INTO `line_operation_times`
                        (`line_id`, `operation_date`, `operation_hours`)
                        VALUES
                        ((SELECT line_id FROM `production_lines` WHERE line_name = :line_name),
                         :operation_date,
                         :operation_hours)
                        '''
            with self.database_process.Session() as session:
                try:
                    import_result = session.execute(text(sql), import_data_list,execution_options={"executemany": True})
                    import_working_time_result = session.execute(text(sql_working_time), working_time_import_list,execution_options={"executemany": True})
                    session.commit()
                    QtWidgets.QMessageBox.information(self, "Success", "Data has been successfully imported to the database.")
                    self.ui.DT_data_table.setUpdatesEnabled(False)
                    self.ui.DT_data_table.setSortingEnabled(False)
                    self.ui.DT_summary_table.setUpdatesEnabled(False)
                    self.ui.DT_summary_table.setSortingEnabled(False)
                    self.DT_model.removeRows(0, self.DT_model.rowCount())
                    self.DT_summary_model.removeRows(0, self.DT_summary_model.rowCount())
                    self.ui.DT_summary_chart_widget.layout().deleteLater()
                    self.ui.DT_data_table.setUpdatesEnabled(True)
                    self.ui.DT_data_table.setSortingEnabled(True)
                    self.ui.DT_summary_table.setUpdatesEnabled(True)
                    self.ui.DT_summary_table.setSortingEnabled(True)
                except Exception:
                    session.rollback()
                    raise
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to import data to database: {e}")
    
    @QtCore.pyqtSlot()
    def Problem_report_Downtime_page(self):
        self.style_button_with_shadow((self.ui.DT_problem_report_btn,self.ui.DT_dashboard_btn,self.ui.DT_data_btn,self.ui.DT_import_data_btn))
        self.ui.DT_stacked_widget.setCurrentWidget(self.ui.DT_problem_report_page)

#==========================================================================================================================


#==========================================================================================================================


#==========================================================================================================================
class Setting_windown(QtWidgets.QDialog):
    def __init__(self, parent=None, model=None):
        super().__init__(parent)
        self.parent = parent
        self.model = model
        self.ui = Ui_SettingWindown()
        self.ui.setupUi(self)
        self.setup_signals()
        self.show_setting_value()

    def setup_signals(self):
        self.ui.default_btn.clicked.connect(self.default)
        self.ui.save_btn.clicked.connect(self.save)
        self.ui.cancel_btn.clicked.connect(self.close)
    
    @QtCore.pyqtSlot() 
    def close(self):
        super().close()
    
    @QtCore.pyqtSlot() 
    def default(self):
        self.model.default_setting()
        self.parent.params = {
            "NG_Coil_Sheetname": self.model.default_NG_Coil_Sheetname,
            "date_col_NG_Coil": self.model.default_date_col_NG_Coil,
            "begin_NG_coil": self.model.default_begin_NG_coil,
            "end_NG_coil": self.model.default_end_NG_coil,
            "NG_Molding_Sheetname": self.model.default_NG_Molding_Sheetname,
            "date_col_NG_Molding": self.model.default_date_col_NG_Molding,
            "begin_NG_Molding": self.model.default_begin_NG_Molding,
            "end_NG_Molding": self.model.default_end_NG_Molding,
            "month": self.model.default_month,
            "year": self.model.default_year,
            "FG_sheet_name": self.model.default_FG_sheet_name,
            "FG_date_col": self.model.default_FG_date_col,
            "FG_line_col": self.model.default_FG_line_col,
            "Molding_lt_sheet_name": self.model.default_Molding_lt_sheet_name,
            "Coil_lt_sheet_name": self.model.default_Coil_lt_sheet_name,
            "lt_date_col": self.model.default_lt_date_col
        }
        self.show_setting_value()

    def show_setting_value(self):
        self.ui.NG_coil_sh_currLine.setText(
            self.parent.params["NG_Coil_Sheetname"])
        self.ui.begin_row_coilNG_currLine.setText(
            str(self.parent.params["begin_NG_coil"]))
        self.ui.end_row_coilNG_currLine.setText(
            str(self.parent.params["end_NG_coil"]))
        self.ui.date_col_coilNG_currLine.setText(
            str(self.parent.params["date_col_NG_Coil"]))
        self.ui.NG_Molding_sh_currLine.setText(
            self.parent.params["NG_Molding_Sheetname"])
        self.ui.begin_row_MoldingNG_currLine.setText(
            str(self.parent.params["begin_NG_Molding"]))
        self.ui.end_row_MoldingNG_currLine.setText(
            str(self.parent.params["end_NG_Molding"]))
        self.ui.date_col_MoldingNG_currLine.setText(
            str(self.parent.params["date_col_NG_Molding"]))
        self.ui.FG_Molding_sh_currLine.setText(
            self.parent.params["FG_sheet_name"])
        self.ui.line_col_MoldingFG_currLine.setText(
            str(self.parent.params["FG_line_col"]))
        self.ui.date_col_MoldingFG_currLine.setText(
            str(self.parent.params["FG_date_col"]))
        self.ui.lt_Molding_sh_currLine.setText(
            self.parent.params["Molding_lt_sheet_name"])
        self.ui.lt_Coil_sh_currLine.setText(
            self.parent.params["Coil_lt_sheet_name"])
        self.ui.lt_date_col_currLine.setText(
            str(self.parent.params["lt_date_col"]))
        self.ui.date_current.setDateTime(QtCore.QDateTime(
            QtCore.QDate(self.parent.params["year"], self.parent.params["month"], 1), QtCore.QTime(0, 0, 0)))
        self.ui.NG_coil_sh_adjLine.setText(
            self.parent.params["NG_Coil_Sheetname"])
        self.ui.begin_row_coilNG_adjLine.setText(
            str(self.parent.params["begin_NG_coil"]))
        self.ui.end_row_coilNG_adjLine.setText(
            str(self.parent.params["end_NG_coil"]))
        self.ui.date_col_coilNG_adjLine.setText(
            str(self.parent.params["date_col_NG_Coil"]))
        self.ui.NG_Molding_sh_adjLine.setText(
            self.parent.params["NG_Molding_Sheetname"])
        self.ui.begin_row_MoldingNG_adjLine.setText(
            str(self.parent.params["begin_NG_Molding"]))
        self.ui.end_row_MoldingNG_adjLine.setText(
            str(self.parent.params["end_NG_Molding"]))
        self.ui.date_col_MoldingNG_adjLine.setText(
            str(self.parent.params["date_col_NG_Molding"]))
        self.ui.FG_Molding_sh_adjLine.setText(
            self.parent.params["FG_sheet_name"])
        self.ui.line_col_MoldingFG_adjLine.setText(
            str(self.parent.params["FG_line_col"]))
        self.ui.date_col_MoldingFG_adjLine.setText(
            str(self.parent.params["FG_date_col"]))
        self.ui.lt_Molding_sh_adjLine.setText(
            self.parent.params["Molding_lt_sheet_name"])
        self.ui.lt_Coil_sh_adjLine.setText(
            self.parent.params["Coil_lt_sheet_name"])
        self.ui.lt_date_col_adjLine.setText(
            str(self.parent.params["lt_date_col"]))
        self.ui.date_adj.setDateTime(QtCore.QDateTime(
            QtCore.QDate(self.parent.params["year"], self.parent.params["month"], 1), QtCore.QTime(0, 0, 0)))
    
    @QtCore.pyqtSlot() 
    def save(self):
        self.model.NG_Coil_Sheetname = self.ui.NG_coil_sh_adjLine.text()
        self.model.begin_NG_coil = int(self.ui.begin_row_coilNG_adjLine.text())
        self.model.end_NG_coil = int(self.ui.end_row_coilNG_adjLine.text())
        self.model.date_col_NG_Coil = int(
            self.ui.date_col_coilNG_adjLine.text())
        self.model.NG_Molding_Sheetname = self.ui.NG_Molding_sh_adjLine.text()
        self.model.begin_NG_Molding = int(
            self.ui.begin_row_MoldingNG_adjLine.text())
        self.model.end_NG_Molding = int(
            self.ui.end_row_MoldingNG_adjLine.text())
        self.model.date_col_NG_Molding = int(
            self.ui.date_col_MoldingNG_adjLine.text())
        self.model.FG_sheet_name = self.ui.FG_Molding_sh_adjLine.text()
        self.model.FG_line_col = int(self.ui.line_col_MoldingFG_adjLine.text())
        self.model.FG_date_col = int(self.ui.date_col_MoldingFG_adjLine.text())
        self.model.Molding_lt_sheet_name = self.ui.lt_Molding_sh_adjLine.text()
        self.model.Coil_lt_sheet_name = self.ui.lt_Coil_sh_adjLine.text()
        self.model.lt_date_col = int(self.ui.lt_date_col_adjLine.text())
        self.model.month = int(self.ui.date_adj.date().month())
        self.parent.params = {
            "NG_Coil_Sheetname": self.model.NG_Coil_Sheetname,
            "date_col_NG_Coil": self.model.date_col_NG_Coil,
            "begin_NG_coil": self.model.begin_NG_coil,
            "end_NG_coil": self.model.end_NG_coil,
            "NG_Molding_Sheetname": self.model.NG_Molding_Sheetname,
            "date_col_NG_Molding": self.model.date_col_NG_Molding,
            "begin_NG_Molding": self.model.begin_NG_Molding,
            "end_NG_Molding": self.model.end_NG_Molding,
            "month": self.model.month,
            "year": self.model.default_year,
            "FG_sheet_name": self.model.FG_sheet_name,
            "FG_date_col": self.model.FG_line_col,
            "FG_line_col": self.model.FG_date_col,
            "Molding_lt_sheet_name": self.model.Molding_lt_sheet_name,
            "Coil_lt_sheet_name": self.model.Coil_lt_sheet_name,
            "lt_date_col": self.model.lt_date_col
        }

        self.change_aftersave()

    def change_aftersave(self):
        self.ui.NG_coil_sh_currLine.setText(
            self.ui.NG_coil_sh_adjLine.text())
        self.ui.begin_row_coilNG_currLine.setText(
            self.ui.begin_row_coilNG_adjLine.text())
        self.ui.end_row_coilNG_currLine.setText(
            self.ui.end_row_coilNG_adjLine.text())
        self.ui.date_col_coilNG_currLine.setText(
            self.ui.date_col_coilNG_adjLine.text())
        self.ui.NG_Molding_sh_currLine.setText(
            self.ui.NG_Molding_sh_adjLine.text())
        self.ui.begin_row_MoldingNG_currLine.setText(
            self.ui.begin_row_MoldingNG_adjLine.text())
        self.ui.end_row_MoldingNG_currLine.setText(
            self.ui.end_row_MoldingNG_adjLine.text())
        self.ui.date_col_MoldingNG_currLine.setText(
            self.ui.date_col_MoldingNG_adjLine.text())
        self.ui.FG_Molding_sh_currLine.setText(
            self.ui.FG_Molding_sh_adjLine.text())
        self.ui.line_col_MoldingFG_currLine.setText(
            self.ui.line_col_MoldingFG_adjLine.text())
        self.ui.date_col_MoldingFG_currLine.setText(
            self.ui.date_col_MoldingFG_adjLine.text())
        self.ui.lt_Molding_sh_currLine.setText(
            self.ui.lt_Molding_sh_adjLine.text())
        self.ui.lt_Coil_sh_currLine.setText(
            self.ui.lt_Coil_sh_adjLine.text())
        self.ui.lt_date_col_currLine.setText(
            self.ui.lt_date_col_adjLine.text())
        self.ui.date_current.setDate(self.ui.date_adj.date())


#==========================================================================================================================


#==========================================================================================================================


#==========================================================================================================================


class View_result_windown(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.database_process = self.parent.database_process
        self.ui = Ui_View_result()
        self.ui.setupUi(self)
        self.list_df = None
        self.machine_code_adjust = None
        self.line_adjust = None
        self.machine_type_adjust = None
        self.setting_condition_flag = {
            "Line": True,
            "Date": True
        }
        try:
            self.department = self.database_process.query(
                sql='''SELECT department_name FROM Departments''', params=None)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to fetch data: {str(e)}")
            super.close()

        self.ui.date_start.setDateTime(QtCore.QDateTime(
            QtCore.QDate(self.parent.params["year"], self.parent.params["month"], 1), QtCore.QTime(0, 0, 0)))
        self.ui.date_end.setDateTime(QtCore.QDateTime(
            QtCore.QDate(self.parent.params["year"], self.parent.params["month"], 1), QtCore.QTime(0, 0, 0)))

        self.ui.type_data_cbb.addItems(["Detail", "Total"])
        self.mode = self.ui.type_data_cbb.currentIndex()
        self.ui.department_cbb.addItems([dep[0] for dep in self.department])
        self.ui.department_cbb.setCurrentText("PE3")
        self.export_line()
        self.ui.line_end_cbb.setCurrentIndex(self.ui.line_end_cbb.count() - 1)
        self.add_widget()
        self.setup_signals()

    def setup_signals(self):
        self.ui.Cancel_btn.clicked.connect(self.close)
        self.ui.Export_all.clicked.connect(self.Export_select)
        self.ui.department_cbb.currentTextChanged.connect(self.export_line)
        self.ui.import_into_DB_btn.clicked.connect(self.import_into_DB)
        self.ui.Query_btn.clicked.connect(lambda: self.view_report(where="DB"))
        self.ui.DBExport_btn.clicked.connect(self.export_from_DB)
        self.ui.line_begin_cbb.currentTextChanged.connect(lambda _: self.check_setting_condition(
            begin=self.split_text_number(
                self.ui.line_begin_cbb.currentText())[0],
            end=self.split_text_number(self.ui.line_end_cbb.currentText())[0],
            button="Line"))
        self.ui.line_end_cbb.currentTextChanged.connect(lambda _: self.check_setting_condition(
            begin=self.split_text_number(
                self.ui.line_begin_cbb.currentText())[0],
            end=self.split_text_number(self.ui.line_end_cbb.currentText())[0],
            button="Line"))
        self.ui.date_start.dateChanged.connect(lambda _: self.check_setting_condition(
            begin=int(self.ui.date_start.date().toString("yyyyMMdd")),
            end=int(self.ui.date_end.date().toString("yyyyMMdd")),
            button="Date"))
        self.ui.type_data_cbb.currentTextChanged.connect(self.mode_change)

    def close(self):
        super().close()

    def get_checked_items(self):
        self.checked_items = []
        for i, checkbox_container in enumerate(self.check_box_list):
            checkbox = checkbox_container.layout().itemAt(0).widget()
            if checkbox.checkState() == QtCore.Qt.Checked:
                line_edit = self.ui.n.cellWidget(
                    i, 0).layout().itemAt(0).widget()
                self.checked_items.append(line_edit.text())

    @QtCore.pyqtSlot()
    def Export_select(self):
        self.get_checked_items()
        if self.checked_items:
            self.parent.Choose_export_machine(
                file_export=self.checked_items, where="Excel")
        else:
            self.parent.Choose_export_machine(file_export=None, where="Excel")
        self.checked_items = None

    def add_widget(self):
        if self.parent.list_df_molding_result is not None:
            self.ui.Result_table.setRowCount(24)
            self.ui.Result_table.setColumnCount(3)
            self.view_report_btn_list = []
            self.check_box_list = []
            for i in range(len(self.parent.list_df_molding_result)):
                self.ui.Result_table.setCellWidget(
                    i, 0, self.create_Textbox(self.parent.list_df_molding_result[i][0]))
                self.check_box = self.create_checkbox_widget()
                self.ui.Result_table.setCellWidget(
                    i, 1, self.check_box)
                view_report_btn = self.create_button_widget(
                    "View report", lambda _, row=i: self.view_report(row, where=None))
                self.ui.Result_table.setCellWidget(i, 2, view_report_btn)
                self.view_report_btn_list.append(view_report_btn)
                self.check_box_list.append(self.check_box)
            return
        self.ui.Export_all.setEnabled(False)
        self.ui.import_into_DB_btn.setEnabled(False)
        QtWidgets.QMessageBox.warning(
            self, "warning", "Chưa có dữ liệu, hãy nhấn Data process trước")

    def create_checkbox_widget(self):
        checkbox = QtWidgets.QCheckBox()
        checkbox.setCheckState(QtCore.Qt.Unchecked)
        checkbox.stateChanged.connect(self.export_state)
        return self.wrap_widget(checkbox)

    def create_button_widget(self, text, callback):
        button = QtWidgets.QPushButton(text)
        button.clicked.connect(callback)
        return button

    def create_Textbox(self, text):
        textbox = QtWidgets.QLabel(text=text)
        return self.wrap_widget(textbox)

    def wrap_widget(self, widget):
        container = QtWidgets.QWidget()
        layout = QtWidgets.QHBoxLayout(container)
        layout.addWidget(widget)
        layout.setAlignment(QtCore.Qt.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)
        container.setLayout(layout)
        return container
    
    @QtCore.pyqtSlot()
    def view_report(self, row=None, where=None):
        if where is None:
            self.list_df = [self.parent.list_df_molding_result,
                            self.parent.list_df_coil_result]
            self.list_df_monthly = [self.parent.list_df_molding_monthly_result,
                                    self.parent.list_df_coil_monthly_result]
            self.show_report_no = row
            self.month = self.parent.params["month"]
            self.year = self.parent.params["year"]
        else:
            if self.mode == 0:
                self.query_data_DB(table="OEE_Daily")
            else:
                self.query_data_DB(table="OEE_Monthly")
            if self.Query_data_result is None or self.Query_data_result.empty:
                QtWidgets.QMessageBox.warning(
                    self, "warning", "Chưa có dữ liệu, hãy kiểm tra lại")
                return
        # self.windown = ReportWindown(parent=self, where=where)
        # self.windown.exec_()
    
    @QtCore.pyqtSlot()
    def export_state(self):
        for i, checkbox_container in enumerate(self.check_box_list):
            checkbox = checkbox_container.layout().itemAt(0).widget()
            if checkbox.checkState() == QtCore.Qt.Checked:
                self.ui.Export_all.setMaximumWidth(150)
                self.ui.Export_all.setText("Export Select")
                return
        self.ui.Export_all.setMaximumWidth(100)
        self.ui.Export_all.setText("Export All")
    
    @QtCore.pyqtSlot()
    def export_line(self):
        self.ui.line_begin_cbb.clear()
        self.ui.line_end_cbb.clear()
        self.dep = self.ui.department_cbb.currentText()
        try:
            self._result = self.database_process.query(sql='''SELECT line_name 
                                                FROM Production_Lines 
                                                JOIN Departments ON Production_Lines.department_id = Departments.department_id
                                                WHERE Departments.department_name =:department_name ''', params={"department_name": f"{self.dep}"})
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to fetch data: {str(e)}")
        self.ui.line_begin_cbb.addItems([ln[0] for ln in self._result])
        self.ui.line_end_cbb.addItems([ln[0] for ln in self._result])

    def default_dict(self):
        self.data_export_A = {
            "line_id": [],
            "month_year_id": [],
            "machine_id": []
        }
        self.data_export_B = {
            "Day": [],
            "Working_shift": [],
            "Net_avaiable_runtime": [],
            "Downtime": [],
            "Other_stop": [],
            "Runtime": [],
            "FGs": [],
            "Defect": [],
            "A": [],
            "P": [],
            "Q": [],
            "OEE": []
        }
        self.data_export_B_monthly = {
            "Working_shift": [],
            "Net_avaiable_runtime": [],
            "Downtime": [],
            "Other_stop": [],
            "Runtime": [],
            "FGs": [],
            "Defect": [],
            "A": [],
            "P": [],
            "Q": [],
            "OEE": []
        }
    
    @QtCore.pyqtSlot()
    def import_into_DB(self):
        self.final_data = pd.DataFrame()
        self.final_data_monthly = pd.DataFrame()
        self.get_checked_items()
        self.parent.Choose_export_machine(file_export=None,  where="DB")
        # if self.Flag_edit_data is False:
        if self.parent.machine_type_for_db == "Molding":
            temp_list = self.parent.list_df_molding_result.copy()
            temp_list_monthly = self.parent.list_df_molding_monthly_result.copy()
        elif self.parent.machine_type_for_db == "Coil":
            temp_list = self.parent.list_df_coil_result
            temp_list_monthly = self.parent.list_df_coil_monthly_result
        else:
            QtWidgets.QMessageBox.warning(
                self, "Error", "Invalid machine target selected.")
            return

        if len(self.checked_items) == 0:
            self.checked_items = [temp_list[i][0]
                                  for i in range(len(temp_list))]
        for line_name in self.checked_items:
            for i in range(len(temp_list)):
                if temp_list[i][0] == line_name:
                    self.default_dict()
                    temp_list[i][1]["Day"] = temp_list[i][1].index
                    temp_list[i][1].reset_index(drop=True, inplace=True)
                    try:
                        self.data_export_A["line_id"].append(self.database_process.query('''SELECT  line_id FROM Production_Lines
                                                            WHERE line_name = :linename''', params={"linename": line_name})[0][0])

                        self.data_export_A["month_year_id"].append(self.database_process.query('''SELECT  month_year_id FROM Months_Years
                                                            WHERE month = :month and year = :year''', params={"month": self.parent.params["month"], "year": self.parent.params["year"]})[0][0])

                        machine_id = self.database_process.query('''SELECT machine_id FROM Machines
                                                            JOIN Production_Lines ON Machines.line_id = Production_Lines.line_id
                                                            WHERE line_name = :linename AND machine_name = :machinename
                                                            ''', params={"linename": line_name, "machinename": "Molding" if self.parent.machine_type_for_db == "Molding" else "Coil"})
                    except Exception as e:
                        QtWidgets.QMessageBox.critical(
                            self, "Error", f"Failed to fetch data: {str(e)}")
                        return
                    try:
                        if (line_name in self.line_adjust) and (self.machine_type_adjust[self.line_adjust.index(line_name)] == self.parent.machine_type_for_db):
                            self.data_export_A["machine_id"] = self.machine_code_adjust[self.line_adjust.index(
                                line_name)]
                    except Exception as e:
                        self.data_export_A["machine_id"] = machine_id[0][0]
                    self.data_df_A = pd.DataFrame(self.data_export_A)
                    self.data_export_B = pd.DataFrame(self.data_export_B)
                    self.data_export_B = temp_list[i][1]
                    self.data_export_B_monthly = temp_list_monthly[i][1]
                    self.data_export_monthly = pd.concat(
                        [self.data_df_A, self.data_export_B_monthly], axis=1)
                    self.data_df_A = self.data_df_A.loc[np.repeat(
                        self.data_df_A.index, len(self.data_export_B["Day"]))].reset_index(drop=True)
                    self.data_export = pd.concat(
                        [self.data_df_A, self.data_export_B], axis=1)
                    temp_list.pop(i)
                    temp_list_monthly.pop(i)
                    break
            self.final_data = pd.concat(
                [self.final_data, self.data_export], ignore_index=True)
            self.final_data_monthly = pd.concat(
                [self.final_data_monthly, self.data_export_monthly], ignore_index=True)
        self.final_data = self.final_data.drop_duplicates()
        self.final_data = self.final_data.reset_index(drop=True)
        self.final_data_monthly = self.final_data_monthly.drop_duplicates()
        self.final_data_monthly = self.final_data_monthly.reset_index(
            drop=True)
        check_data_empty = self.check_duplicate_from_DB(
            data=self.final_data, what_table="OEE_Daily")
        check_data_data_monthly_empty = self.check_duplicate_from_DB(
            data=self.final_data_monthly, what_table="OEE_Monthly")
        if check_data_empty.empty or check_data_data_monthly_empty.empty:
            reply = QtWidgets.QMessageBox.question(
                self, "warning", "Dữ liệu đã tồn tại trong DB, bạn có muốn cập nhật lại không?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.No:
                return
            else:
                self.final_data = self.check_duplicate_from_DB(
                    data=self.final_data, what_table="OEE_Daily", check_alternate=True)
                self.final_data_monthly = self.check_duplicate_from_DB(
                    data=self.final_data_monthly, what_table="OEE_Monthly", check_alternate=True)
                return
        try:
            with self.database_process.engine.begin() as conn:
                if self.final_data is not None:
                    self.final_data.to_sql(
                        "OEE_Daily", conn, if_exists="append", index=False)
                if self.final_data_monthly is not None:
                    self.final_data_monthly.to_sql(
                        "OEE_Monthly", conn, if_exists="append", index=False)
            QtWidgets.QMessageBox.information(
                self, "Success", "Data imported successfully.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to import data: {str(e)}")
        finally:
            self.checked_items = None

    def check_duplicate_from_DB(self, data=None, what_table=None , check_alternate=False):
        if not check_alternate:
            try:
                existing_records = self.database_process.query(f'''
                                                            SELECT line_id, month_year_id, machine_id
                                                            FROM {what_table}
                                                        ''', params=None)
                existing_set = set(existing_records)
                data_no_duplicate = data.loc[~data.set_index(
                    ['line_id', 'month_year_id', 'machine_id']).index.isin(existing_set)].reset_index(drop=True)
                if data_no_duplicate.empty:
                    QtWidgets.QMessageBox.information(
                        self, "Dữ liệu đã tồn tại", "Dữ liệu đã tồn tại trong DB")
                return data_no_duplicate
            except Exception as e:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Failed to import data: {str(e)}")
                return data_no_duplicate.empty
        else:
            try:
                if what_table == "OEE_Daily":
                    existing_records = self.database_process.query(f'''
                                                                SELECT line_id, month_year_id, machine_id, Day ,OEE
                                                                FROM {what_table}
                                                            ''', params=None)
                    existing_set = pd.MultiIndex.from_tuples(existing_records, names=['line_id', 'month_year_id', 'machine_id', 'Day', 'OEE'])
                    data_no_duplicate = data.loc[~data.set_index(
                        ['line_id', 'month_year_id', 'machine_id','Day','OEE']).index.isin(existing_set)].reset_index(drop=True)
                else:
                    existing_records = self.database_process.query(f'''
                                                                SELECT line_id, month_year_id, machine_id ,OEE
                                                                FROM {what_table}
                                                            ''', params=None)
                    existing_set = pd.MultiIndex.from_tuples(existing_records, names=['line_id', 'month_year_id', 'machine_id', 'OEE'])
                    data_no_duplicate = data.loc[~data.set_index(
                        ['line_id', 'month_year_id', 'machine_id','OEE']).index.isin(existing_set)].reset_index(drop=True)
                if data_no_duplicate.empty:
                    QtWidgets.QMessageBox.information(
                        self, "Không có dữ liệu mới", "Dữ liệu đã tồn tại trong DB")
                else:
                    update_query = f'''
                            UPDATE {what_table}
                            SET Working_shift = :Working_shift,
                            Net_avaiable_runtime = :Net_avaiable_runtime,
                            Downtime = :Downtime,
                            Other_stop = :Other_stop,
                            Runtime = :Runtime,
                            FGs = :FGs,
                            Defect = :Defect,
                            A = :A,
                            P = :P, 
                            Q = :Q,
                            OEE = :OEE
                            WHERE line_id = :line_id AND month_year_id = :month_year_id AND machine_id = :machine_id AND Day = :Day
                        '''
                    params = [
                            {
                                "Working_shift": row["Working_shift"],
                                "Net_avaiable_runtime": row["Net_avaiable_runtime"],
                                "Downtime": row["Downtime"],
                                "Other_stop": row["Other_stop"],
                                "Runtime": row["Runtime"],
                                "FGs": row["FGs"],
                                "Defect": row["Defect"],
                                "A": row["A"],
                                "P": row["P"],
                                "Q": row["Q"],
                                "OEE": row["OEE"],
                                "line_id": row["line_id"],
                                "month_year_id": row["month_year_id"],
                                "machine_id": row["machine_id"],
                                "Day": row["Day"]
                            }
                        for _, row in data_no_duplicate.iterrows()
                            ]
                    self.database_process.query(update_query, params)
            except Exception as e:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Failed to import data: {str(e)}")
                return
    
    @QtCore.pyqtSlot() 
    def check_setting_condition(self, begin=None, end=None, button=None):
        if begin == "" or end == "" or begin > end:
            QtWidgets.QMessageBox.warning(
                self, "warning", f"Dữ liệu chọn tại {button} sai, hãy kiểm tra lại")
            self.setting_condition_flag[button] = False
            self.ui.Query_btn.setEnabled(False)
            self.ui.DBExport_btn.setEnabled(False)
            return
        else:
            self.setting_condition_flag[button] = True
            if all(self.setting_condition_flag.values()):
                self.ui.Query_btn.setEnabled(True)
                self.ui.DBExport_btn.setEnabled(True)
            else:
                self.ui.Query_btn.setEnabled(False)
                self.ui.DBExport_btn.setEnabled(False)
    
    @QtCore.pyqtSlot()
    def mode_change(self):
        self.mode = self.ui.type_data_cbb.currentIndex()

    def query_data_DB(self, table=None):
        line_begin = self.ui.line_begin_cbb.currentText()
        line_end = self.ui.line_end_cbb.currentText()
        start_month = self.ui.date_start.date().toString("MM")
        start_year = self.ui.date_start.date().toString("yyyy")
        end_month = self.ui.date_end.date().toString("MM")
        end_year = self.ui.date_end.date().toString("yyyy")
        line_begin, line_letter = self.split_text_number(line_begin)
        line_end, _ = self.split_text_number(line_end)
        self.line_list = [f"{line_letter}0{i}" if i <
                          10 else f"{line_letter}{i}" for i in range(line_begin, line_end+1)]
        self.month_list = [month for month in range(
            int(start_month), int(end_month)+1)]
        self.year_list = [year for year in range(
            int(start_year), int(end_year)+1)]
        line_placeholders = ', '.join(
            [f":line_{i}" for i in range(len(self.line_list))])
        month_placeholders = ', '.join(
            [f":month_{i}" for i in range(len(self.month_list))])
        year_placeholders = ', '.join(
            [f":year_{i}" for i in range(len(self.year_list))])
        params = {f"line_{i}": line for i, line in enumerate(self.line_list)}
        params.update({f"month_{i}": month for i,
                      month in enumerate(self.month_list)})
        params.update({f"year_{i}": year for i,
                      year in enumerate(self.year_list)})
        try:
            if table == "OEE_Daily":
                self.Query_data_result = pd.DataFrame(self.database_process.query(f'''SELECT PL.line_name, MC.machine_name, MC.machine_id, MY.year, MY.month, OEE.Day, OEE.Working_shift, OEE.Net_avaiable_runtime, OEE.Downtime, OEE.Other_stop, OEE.Runtime, OEE.FGs, OEE.Defect, OEE.A, OEE.P, OEE.Q, OEE.OEE 
                                                                                        FROM {table} as OEE
                                                                                        JOIN `Months_Years` as MY ON OEE.month_year_id = MY.month_year_id
                                                                                        JOIN `Production_Lines` as PL ON OEE.line_id = PL.line_id
                                                                                        JOIN `Machines` as MC ON OEE.machine_id = MC.machine_id
                                                                                        WHERE PL.line_name IN ({line_placeholders}) AND MY.month IN ({month_placeholders}) AND MY.year IN ({year_placeholders})
                                                                                        ORDER BY MY.month ASC, MY.year ASC, OEE.line_id;''', params))
                self.Query_data_result.columns = [
                    'Line', 'Machine_Name', 'Machine_Code', 'Year', 'Month', 'Day', "Working_shift", "Net_avaiable_runtime", "Downtime", "Other_stop", "Runtime", "FGs", "Defect", 'A', 'P', 'Q', 'OEE']
            else:
                self.Query_data_result = pd.DataFrame(self.database_process.query(f'''SELECT PL.line_name, MC.machine_name, MC.machine_id, MY.year, MY.month, OEE.Working_shift, OEE.Net_avaiable_runtime, OEE.Downtime, OEE.Other_stop, OEE.Runtime, OEE.FGs, OEE.Defect, OEE.A, OEE.P, OEE.Q, OEE.OEE 
                                                                                        FROM {table} as OEE
                                                                                        JOIN `Months_Years` as MY ON OEE.month_year_id = MY.month_year_id
                                                                                        JOIN `Production_Lines` as PL ON OEE.line_id = PL.line_id
                                                                                        JOIN `Machines` as MC ON OEE.machine_id = MC.machine_id
                                                                                        WHERE PL.line_name IN ({line_placeholders}) AND MY.month IN ({month_placeholders}) AND MY.year IN ({year_placeholders})
                                                                                        ORDER BY MY.month ASC, MY.year ASC, OEE.line_id;''', params))

                self.Query_data_result.columns = [
                    'Line', 'Machine_Name', 'Machine_Code', 'Year', 'Month', "Working_shift", "Net_avaiable_runtime", "Downtime", "Other_stop", "Runtime", "FGs", "Defect", 'A', 'P', 'Q', 'OEE']
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to fetch data: {str(e)}")
            return

    def split_text_number(self, text):
        if not text:
            self.ui.Query_btn.setEnabled(False)
            self.ui.DBExport_btn.setEnabled(False)
            return 0, ''
        letter_part = ''.join(filter(str.isalpha, text))
        number_part = int(''.join(filter(str.isdigit, text)))
        return number_part, letter_part
    
    @QtCore.pyqtSlot()
    def export_from_DB(self):
        if self.ui.type_data_cbb.currentText() == "Detail":
            self.query_data_DB(table="OEE_Daily")
        if self.ui.type_data_cbb.currentText() == "Total":
            self.query_data_DB(table="OEE_Monthly")
        if self.Query_data_result is None or self.Query_data_result.empty:
            QtWidgets.QMessageBox.warning(
                self, "warning", "Chưa có dữ liệu, hãy kiểm tra lại")
            return
        self.parent.Choose_export_machine(file_export=None, where="DB")
        self.parent.machine_type_for_db = "Molding" if self.parent.machine_type_for_db == "Molding" else "Coil"
        save_path = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Chọn thư mục lưu file", "")
        if not save_path:
            QtWidgets.QMessageBox.warning(
                self, "Thiếu thông tin", "Vui lòng chọn thư mục để lưu file.")
            return
        try:
            for line_name in self.line_list:
                for year in self.year_list:
                    for month in self.month_list:
                        export_frame = self.Query_data_result.query(
                            f"Line == '{line_name}' and Machine_Name == '{self.parent.machine_type_for_db}' and Year == {year} and Month == {month}")
                        if not export_frame.empty:
                            export_frame.to_excel(os.path.join(
                                save_path, f"OEE_{self.parent.machine_type_for_db}_result_{line_name}_{year}_{month}.xlsx"), index=False)
                        else:
                            QtWidgets.QMessageBox.warning(
                                self, "warning", f"Không có dữ liệu cho {line_name} trong tháng {month} năm {year} của máy {self.parent.machine_type_for_db}.")
            QtWidgets.QMessageBox.information(
                self, "Hoàn tất", "File Excel đã được lưu hoàn tất.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", "Lỗi khi lưu file: {e}.")
            return

    def receivers(self, machine_code_adjust=None, machine_type_adjust=None, line_adjust=None, data_receivers = None):
        self.machine_code_adjust = machine_code_adjust
        self.machine_type_adjust = machine_type_adjust
        self.line_adjust = line_adjust
        # self.list_df

#==========================================================================================================================


#==========================================================================================================================


#==========================================================================================================================

# class ReportWindown(QtWidgets.QDialog):
#     def __init__(self, parent=None, where=None):
#         super().__init__(parent)
#         self.parent = parent
#         self.machine_info = self.parent.parent.machine_info
#         self.where = where
#         self.database_process = self.parent.database_process
#         self.setWindowTitle("Report ")
#         self.ui = Ui_Result_chart()
#         self.ui.setupUi(self)
#         self.mode = self.parent.mode
#         self.show_flag = False
#         self.Flag_data_change = False
#         self.Flag_view_type = False
#         self.machine_code_adjust = []
#         self.line_adjust = []
#         self.machine_type_adjust = []
#         self.data_update = None
#         if self.parent.machine_code_adjust is not None:
#             self.machine_code_adjust = self.parent.machine_code_adjust
#             self.line_adjust = self.parent.line_adjust
#             self.machine_type_adjust = self.parent.machine_type_adjust
#         if self.mode == 0:
#             self.ui.Mode_btn.setStyleSheet("QPushButton {\n"
#                                            "    background-color: qlineargradient(\n"
#                                            "        x1: 0, y1: 0, x2: 1, y2: 1,\n"
#                                            "        stop: 0 #83f3ff,\n"
#                                            "        stop: 1 #eeeeee\n"
#                                            "    );\n"
#                                            "    border: 2px solid #009dac;;\n"
#                                            "    border-radius: 25px;\n"
#                                            "    padding: 5px 10px;\n"
#                                            "    font-size: 14px;\n"
#                                            "}")
#         else:
#             self.ui.Mode_btn.setText("Total Mode")
#             self.ui.Mode_btn.setStyleSheet("QPushButton {\n"
#                                            "    background-color: qlineargradient(\n"
#                                            "        x1: 0, y1: 0, x2: 1, y2: 1,\n"
#                                            "        stop: 0 #71ff99,\n"
#                                            "        stop: 1 #eeeeee\n"
#                                            "    );\n"
#                                            "    border: 2px solid #009329;\n"
#                                            "    border-radius: 25px;\n"
#                                            "    padding: 5px 10px;\n"
#                                            "    font-size: 14px;\n"
#                                            "}")
#         if self.where == "DB":
#             self.ui.Mode_btn.setEnabled(False)
#             self.ui.edit_btn.setEnabled(False)
#             self.ui.machine_code_edit.setEnabled(False)
#             self.ui.edit_btn.hide()
#             self.ui.save_btn.hide()
#         self.webview = QWebEngineView()
#         self.ui.machine_cbb.addItems(["Molding", "Coil"])
#         self.ui.machine_cbb.setCurrentIndex(0)
#         self.ui.result_data_frame.hide()
#         self.ui.result_data_frame.setMinimumWidth(0)
#         self.ui.result_data_frame.setMaximumWidth(0)
#         self.setSizePolicy(QtWidgets.QSizePolicy.Expanding,
#                            QtWidgets.QSizePolicy.Expanding)
#         self.adjustSize()

#         if self.where is None:
#             self.ui.date.setDateTime(QtCore.QDateTime(
#                 QtCore.QDate(self.parent.year, self.parent.month, 1), QtCore.QTime(0, 0, 0)))
#             self.ui.line_cbb.addItems(
#                 self.parent.list_df[0][i][0] for i in range(len(self.parent.list_df[0]))
#             )
#             self.ui.line_cbb.setCurrentIndex(self.parent.show_report_no)
#             self.update_chart(line=self.parent.show_report_no)
#         else:
#             self.ui.line_cbb.addItems(
#                 self.parent.Query_data_result["Line"].drop_duplicates().values.tolist())
#             self.ui.date.setDateTime(QtCore.QDateTime(
#                 QtCore.QDate(self.parent.Query_data_result["Year"].iloc[0], self.parent.Query_data_result["Month"].iloc[0], 1), QtCore.QTime(0, 0, 0)))
#             self.update_chart()
#         self.ui.chart_layout.addWidget(self.webview)
#         self.channel = QWebChannel()
#         self.channel.registerObject("bridge", self)
#         self.webview.page().setWebChannel(self.channel)
#         self.ui.save_btn.setEnabled(False)
#         self.setup_signals()

#     def setup_signals(self):
#         self.ui.line_cbb.currentTextChanged.connect(
#             lambda: self.update_chart(line=None))
#         self.ui.machine_cbb.currentTextChanged.connect(
#             lambda: self.update_chart(line=None))
#         self.ui.back_btn.clicked.connect(
#             lambda: self.change_chart(button=self.ui.back_btn))
#         self.ui.next_btn.clicked.connect(
#             lambda: self.change_chart(button=self.ui.next_btn))
#         self.ui.date.dateChanged.connect(lambda: self.update_chart(line=None))
#         self.ui.Mode_btn.clicked.connect(self.change_mode)
#         self.ui.machine_code_edit.textChanged.connect(
#             self.adjust_machinne_code)
#         self.ui.edit_btn.clicked.connect(self.change_view)
#         self.ui.table_widget.itemChanged.connect(self.update_data_frame)
#         self.ui.save_btn.clicked.connect(self.save_data_frame)
    
#     @QtCore.pyqtSlot()
#     def update_chart(self, line = None, df = None):
#         self.df = None
#         if df is None:
#             if self.data_update is not None:
#                 reply = QtWidgets.QMessageBox.question(
#                     self,
#                     "Xác nhận",
#                     "Bạn có muốn lưu data đã chỉnh sửa không?",
#                     QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
#                     QtWidgets.QMessageBox.No
#                 )
#                 if reply == QtWidgets.QMessageBox.Yes:
#                     self.save_data_frame()
#                 else:
#                     self.data_update = None
#             self.df = self._get_chart_data(line)
#         else:
#             self.df = df
#         if self.df.empty:
#             return
#         if self.where == "DB" and self.mode == 1:
#             self.ui.machine_code_edit.setText(
#                 str(self.parent.Query_data_result["Machine_Code"].iloc[0]))
#             if self.show_flag:
#                 pie_chart_fig = self._create_pie_chart(self.df)
#                 pie_chart_fig.update_layout(
#                     margin=dict(l=50, r=50, t=50, b=50))
#                 pie_chart_html = self._wrap_html_with_qt_channel(
#                     pie_chart_fig.to_html(include_plotlyjs="cdn", full_html=False))
#                 self.pie_chart_view.setHtml(
#                     pie_chart_html, QtCore.QUrl("qrc:///"))
#                 bar_chart_fig = self._create_bar_chart(
#                     self.parent.Query_data_result)
#                 bar_chart_fig.update_layout(
#                     margin=dict(l=50, r=50, t=50, b=50))
#                 bar_chart_html = self._wrap_html_with_qt_channel(
#                     bar_chart_fig.to_html(include_plotlyjs="cdn", full_html=False))
#                 self.bar_chart_view.setHtml(
#                     bar_chart_html, QtCore.QUrl("qrc:///"))
#                 return
#             split_layout = QtWidgets.QVBoxLayout()
#             self.pie_chart_view = QWebEngineView()
#             self.bar_chart_view = QWebEngineView()
#             self.pie_chart_view.setMaximumHeight(300)
#             self.bar_chart_view.setMaximumHeight(300)
#             split_layout.addWidget(self.pie_chart_view)
#             split_layout.addWidget(self.bar_chart_view)
#             split_layout.setAlignment(QtCore.Qt.AlignBottom)
#             self.ui.chart_layout.addLayout(split_layout)
#             pie_chart_fig = self._create_pie_chart(self.df)
#             bar_chart_fig = self._create_bar_chart(
#                 self.parent.Query_data_result)
#             pie_chart_fig.update_layout(margin=dict(l=50, r=50, t=50, b=50))
#             bar_chart_fig.update_layout(margin=dict(l=50, r=50, t=50, b=50))
#             pie_chart_html = self._wrap_html_with_qt_channel(
#                 pie_chart_fig.to_html(include_plotlyjs="cdn", full_html=False))
#             self.pie_chart_view.setHtml(pie_chart_html, QtCore.QUrl("qrc:///"))
#             bar_chart_html = self._wrap_html_with_qt_channel(
#                 bar_chart_fig.to_html(include_plotlyjs="cdn", full_html=False))
#             self.bar_chart_view.setHtml(bar_chart_html, QtCore.QUrl("qrc:///"))
#         else:
#             self.query_machine_id()
#             if self.mode == 0:
#                 fig = self._create_bar_chart(self.df)
#             else:
#                 fig = self._create_pie_chart(self.df)
#             fig.update_layout(margin=dict(l=50, r=50, t=50, b=50))
#             html = self._wrap_html_with_qt_channel(
#                 fig.to_html(include_plotlyjs="cdn", full_html=False))
#             self.webview.setHtml(html, QtCore.QUrl("qrc:///"))

#     def _get_chart_data(self, line):
#         if self.where is None:
#             if self.mode == 0:
#                 self.line = self.ui.line_cbb.currentIndex() if line is None else line
#                 self.machine_type = self.ui.machine_cbb.currentIndex()

#                 if self.ui.date.date() == QtCore.QDate(self.parent.year, self.parent.month, 1):
#                     return self.parent.list_df[0][self.line][1] if not self.machine_type else self.parent.list_df[1][self.line][1]
#                 else:
#                     return self.query_data(table="OEE_Daily")
#             else:
#                 self.line = self.ui.line_cbb.currentIndex() if line is None else line
#                 self.machine_type = self.ui.machine_cbb.currentIndex()

#                 if self.ui.date.date() == QtCore.QDate(self.parent.year, self.parent.month, 1):
#                     return self.parent.list_df_monthly[0][self.line][1] if not self.machine_type else self.parent.list_df_monthly[1][self.line][1]
#                 else:
#                     return self.query_data(table="OEE_Monthly")
#         else:
#             result = self.parent.Query_data_result[
#                 (self.parent.Query_data_result["Line"] == self.ui.line_cbb.currentText()) &
#                 (self.parent.Query_data_result["Machine_Name"] == self.ui.machine_cbb.currentText()) &
#                 (self.parent.Query_data_result["Year"] == self.ui.date.date().year()) &
#                 (self.parent.Query_data_result["Month"] == self.ui.date.date().month())]
#             self.line = self.ui.line_cbb.currentIndex() if line is None else line
#             self.machine_type = self.ui.machine_cbb.currentIndex()
#             return result

#     def _create_bar_chart(self, df=None):
#         if self.where == "DB" and self.mode == 1:
#             fig = go.Figure()
#             df_fixed = pd.DataFrame({'Month': range(1, 13)})
#             df_merged = df_fixed.merge(df[(df["Line"] == self.ui.line_cbb.currentText()) & (
#                 df["Machine_Name"] == self.ui.machine_cbb.currentText())], on='Month', how='left').fillna(0)
#             df_merged["OEE"] = df_merged["OEE"].clip(lower=0, upper=1.5)
#             fig.add_trace(go.Bar(
#                 x=df_merged["Month"], y=df_merged["OEE"], name="OEE", marker_color='royalblue'))
#             fig.update_yaxes(range=[0, 1.5], fixedrange=True)
#             fig.update_layout(title_text="OEE Breakdown Over Year",
#                               height=300, width=1530, showlegend=False)
#             self.show_flag = True
#         else:
#             if self.where is None:
#                 df["Day"] = df.index
#             fig = make_subplots(rows=2, cols=2,
#                                 subplot_titles=(
#                                     "OEE", "A (Availability)", "P (Performance)", "Q (Quality)"),
#                                 vertical_spacing=0.15, horizontal_spacing=0.1)

#             metrics = ['OEE', 'A', 'P', 'Q']
#             colors = ['royalblue', 'seagreen', 'firebrick', 'mediumpurple']
#             df[metrics] = df[metrics].clip(lower=0, upper=1.5)

#             for i, (metric, color) in enumerate(zip(metrics, colors)):
#                 row, col = i // 2 + 1, i % 2 + 1
#                 fig.add_trace(go.Bar(
#                     x=df["Day"], y=df[metric], name=metric, marker_color=color), row=row, col=col)
#                 fig.update_yaxes(
#                     range=[0, 1.5], fixedrange=True, row=row, col=col)
#             fig.update_layout(title_text="OEE Breakdown Over Days",
#                               height=650, width=1530, showlegend=False)
#         return fig

#     def _create_pie_chart(self, df):
#         if df.empty:
#             return go.Figure()

#         last_row = df.iloc[-1]
#         metrics = ['OEE', 'A', 'P', 'Q']
#         titles = ['OEE', 'Availability (A)', 'Performance (P)', 'Quality (Q)']

#         fig = make_subplots(
#             rows=1, cols=4,
#             specs=[[{'type': 'domain'}, {'type': 'domain'},
#                     {'type': 'domain'}, {'type': 'domain'}]],
#             subplot_titles=titles
#         )
#         color_map = {
#             'OEE': '#636EFA',        # Xanh dương
#             'A': '#00CC96',          # Xanh lá
#             'P': '#AB63FA',          # Tím
#             'Q': '#FFA15A',           # Cam
#         }
#         remaining_color = '#EF553B'
#         for i, metric in enumerate(metrics):
#             value = last_row.get(metric, 0)
#             remainder = max(1 - value, 0)
#             fig.add_trace(go.Pie(
#                 labels=[metric, "Remaining"],
#                 values=[value, remainder],
#                 hole=0.4,
#                 name=metric,
#                 textinfo='percent',
#                 showlegend=True,
#                 sort=False,
#                 marker=dict(colors=[color_map.get(
#                     metric, '#CCCCCC'), remaining_color])
#             ), row=1, col=i + 1)
#         if self.where == "DB" and self.mode == 1:
#             fig.update_layout(
#                 title_text=None,
#                 height=300, width=1530)
#             for annotation in fig['layout']['annotations']:
#                 annotation['y'] += 0.1
#         else:
#             fig.update_layout(
#                 title_text="OEE Breakdowns",
#                 height=650, width=1530
#             )
#         return fig

#     def _wrap_html_with_qt_channel(self, plot_html):
#         return f"""
#         <html>
#         <head>
#             <script src="qrc:///qtwebchannel/qwebchannel.js"></script>
#         </head>
#         <body>
#             {plot_html}
#             <script>
#             new QWebChannel(qt.webChannelTransport, function(channel) {{
#                 const bridge = channel.objects.bridge;
#                 const plot = document.getElementsByClassName('plotly-graph-div')[0];
#                 plot.on('plotly_click', function(data) {{
#                     const point = data.points[0];
#                     const info = JSON.stringify({{
#                         x: point.x,
#                         y: point.y,
#                         series: point.data.name
#                     }});
#                     bridge.bar_clicked(info);
#                 }});
#             }});
#             </script>
#         </body>
#         </html>
#         """
    
#     @QtCore.pyqtSlot()
#     def change_chart(self, button=None):
#         if button == self.ui.back_btn:
#             if self.ui.line_cbb.currentIndex() == 0:
#                 self.ui.line_cbb.setCurrentIndex(self.ui.line_cbb.count()-1)
#             else:
#                 self.ui.line_cbb.setCurrentIndex(
#                     self.ui.line_cbb.currentIndex()-1)
#         else:
#             if self.ui.line_cbb.currentIndex() == (self.ui.line_cbb.count()-1):
#                 self.ui.line_cbb.setCurrentIndex(0)
#             else:
#                 self.ui.line_cbb.setCurrentIndex(
#                     self.ui.line_cbb.currentIndex()+1)
    
#     @QtCore.pyqtSlot()
#     def change_mode(self):
#         if not self.mode:
#             self.ui.edit_btn.setEnabled(False)
#             self.ui.Mode_btn.setText("Total Mode")
#             self.ui.Mode_btn.setStyleSheet("QPushButton {\n"
#                                            "    background-color: qlineargradient(\n"
#                                            "        x1: 0, y1: 0, x2: 1, y2: 1,\n"
#                                            "        stop: 0 #71ff99,\n"
#                                            "        stop: 1 #eeeeee\n"
#                                            "    );\n"
#                                            "    border: 2px solid #009329;\n"
#                                            "    border-radius: 25px;\n"
#                                            "    padding: 5px 10px;\n"
#                                            "    font-size: 14px;\n"
#                                            "}")
#             self.mode = 1
#         else:
#             self.ui.edit_btn.setEnabled(True)
#             self.ui.Mode_btn.setText("Detail Mode")
#             self.ui.Mode_btn.setStyleSheet("QPushButton {\n"
#                                            "    background-color: qlineargradient(\n"
#                                            "        x1: 0, y1: 0, x2: 1, y2: 1,\n"
#                                            "        stop: 0 #83f3ff,\n"
#                                            "        stop: 1 #eeeeee\n"
#                                            "    );\n"
#                                            "    border: 2px solid #009dac;\n"
#                                            "    border-radius: 25px;\n"
#                                            "    padding: 5px 10px;\n"
#                                            "    font-size: 14px;\n"
#                                            "}")
#             self.mode = 0
#         self.update_chart(line=None)

#     def query_data(self, table=None):
#         _year = self.ui.date.date().toString("yyyy")
#         _month = self.ui.date.date().toString("MM")
#         try:
#             if table == "OEE_Daily":
#                 change_date_data_query = pd.DataFrame(self.database_process.query(f'''SELECT Day, A, P, Q, OEE FROM {table}
#                                                                 JOIN Production_Lines ON {table}.line_id = Production_Lines.line_id
#                                                                 JOIN Months_Years ON {table}.month_year_id = Months_Years.month_year_id
#                                                                 JOIN Machines ON {table}.machine_id = Machines.machine_id
#                                                                 WHERE month = :month AND year = :year AND line_name = :linename AND machine_name = :machinename''', params={"month": _month, "year": _year, "linename": self.ui.line_cbb.currentText(), "machinename": self.ui.machine_cbb.currentText()}))

#                 change_date_data_query.columns = ['Day', 'A', 'P', 'Q', 'OEE']
#             else:
#                 change_date_data_query = pd.DataFrame(self.database_process.query(f'''SELECT month, A, P, Q, OEE FROM {table}
#                                                                 JOIN Production_Lines ON {table}.line_id = Production_Lines.line_id
#                                                                 JOIN Months_Years ON {table}.month_year_id = Months_Years.month_year_id
#                                                                 JOIN Machines ON {table}.machine_id = Machines.machine_id
#                                                                 WHERE month = ? AND year = ? AND line_name = ? AND machine_name = ?''', params={"month": _month, "year": _year, "linename": self.ui.line_cbb.currentText(), "machinename": self.ui.machine_cbb.currentText()}))
#                 change_date_data_query.columns = [
#                     'Month', 'A', 'P', 'Q', 'OEE']
#         except Exception as e:
#             QtWidgets.QMessageBox.critical(
#                 self, "Error", f"Fail to query: {str(e)}")
#             return
#         return change_date_data_query

#     def query_machine_id(self):
#         self.ui.machine_code_edit.clear()
#         if self.ui.line_cbb.currentText() in self.line_adjust:
#             self.ui.machine_code_edit.setStyleSheet(
#                 "QLineEdit { background-color: rgb(255,162,98); }")
#             self.ui.machine_code_edit.setText(
#                 self.machine_code_adjust[self.line_adjust.index(self.ui.line_cbb.currentText())])
#             return
#         else:
#             self.ui.machine_code_edit.setStyleSheet(
#                 "QLineEdit { background-color: rgb(255,255,255); }")
#         try:
#             self.machine_code = self.database_process.query('''SELECT machine_id FROM Machines
#                                                             JOIN Production_Lines ON Machines.line_id = Production_Lines.line_id
#                                                             WHERE line_name = :linename AND machine_name = :machinename
#                                                             ''', params={"linename": self.ui.line_cbb.itemText(self.line), "machinename": self.ui.machine_cbb.itemText(self.machine_type)})
#             self.ui.machine_code_edit.setPlaceholderText(
#                 ' + '.join(str(code[0]) for code in self.machine_code)
#             )
#         except Exception as e:
#             QtWidgets.QMessageBox.critical(
#                 self, "Error", f"Failed to fetch machine names: {e}")
#             self.ui.machine_code_edit.setPlaceholderText("")

#     @QtCore.pyqtSlot()
#     def adjust_machinne_code(self):
#         suggestions = []
#         self.machine_code = []
#         try:
#             self.machine_code = self.database_process.query('''SELECT machine_id FROM Machines
#                                                             JOIN Production_Lines ON Production_Lines.line_id = Machines.line_id
#                                                             JOIN Departments ON Departments.department_id = Production_Lines.department_id
#                                                             WHERE  department_name = :department AND machine_name = :machinename LIMIT 5
#                                                             ''', params={"department": self.parent.ui.department_cbb.currentText(), "machinename": self.ui.machine_cbb.itemText(self.machine_type)})
#             suggestions = [str(name[0]) for name in self.machine_code]
#         except Exception as e:
#             QtWidgets.QMessageBox.critical(
#                 self, "Error", f"Failed to fetch machine names: {e}")
#         completer = QtWidgets.QCompleter(suggestions, self)
#         completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
#         self.ui.machine_code_edit.setCompleter(completer)
#         if (self.ui.machine_code_edit.text() != self.ui.machine_code_edit.placeholderText()) and (self.ui.machine_code_edit.text() in suggestions):
#             self.line_adjust.append(self.ui.line_cbb.currentText())
#             self.machine_type_adjust.append(self.ui.machine_cbb.currentText())
#             self.machine_code_adjust.append(self.ui.machine_code_edit.text())
#             self.ui.machine_code_edit.setStyleSheet(
#                 "QLineEdit { background-color: rgb(255,162,98); }")
    
#     @QtCore.pyqtSlot()
#     def change_view(self):
#         if self.ui.widget.isHidden():
#             self.ui.edit_btn.setText("Data")
#             self.ui.widget_2.hide()
#             self.ui.widget.show()
#             self.ui.next_btn.setEnabled(False)
#             self.ui.back_btn.setEnabled(False)
#             self.edit_data()
#             self.Flag_view_type = True
#         else:
#             self.ui.edit_btn.setText("Chart")
#             self.ui.widget.hide()
#             self.ui.widget_2.show()
#             self.set_data_update()
#             self.update_chart(line=None, df=self.data_update)
#             self.ui.next_btn.setEnabled(True)
#             self.ui.back_btn.setEnabled(True)
#             self.Flag_view_type = False

#     def edit_data(self):
#         if self.df is not None:
#             data = self.df
#             self.ui.table_widget.setRowCount(data["Day"].shape[0])
#             for row, (day, df) in enumerate(data.iterrows()):
#                 self.ui.table_widget.setItem(
#                     row, 0, QtWidgets.QTableWidgetItem(str(day)))
#                 self.ui.table_widget.setItem(
#                     row, 1, QtWidgets.QTableWidgetItem(str(df["Working_shift"])))
#                 self.ui.table_widget.setItem(row, 2, QtWidgets.QTableWidgetItem(
#                     str(df["Net_avaiable_runtime"])))
#                 self.ui.table_widget.setItem(
#                     row, 3, QtWidgets.QTableWidgetItem(str(df["Downtime"])))
#                 self.ui.table_widget.setItem(
#                     row, 4, QtWidgets.QTableWidgetItem(str(df["Other_stop"])))
#                 self.ui.table_widget.setItem(
#                     row, 5, QtWidgets.QTableWidgetItem(str(df["Runtime"])))
#                 self.ui.table_widget.setItem(
#                     row, 6, QtWidgets.QTableWidgetItem(str(df["FGs"])))
#                 self.ui.table_widget.setItem(
#                     row, 7, QtWidgets.QTableWidgetItem(str(df["Defect"])))
#                 self.ui.table_widget.setItem(
#                     row, 8, QtWidgets.QTableWidgetItem(str(df["A"])))
#                 self.ui.table_widget.setItem(
#                     row, 9, QtWidgets.QTableWidgetItem(str(df["P"])))
#                 self.ui.table_widget.setItem(
#                     row, 10, QtWidgets.QTableWidgetItem(str(df["Q"])))
#                 self.ui.table_widget.setItem(
#                     row, 11, QtWidgets.QTableWidgetItem(str(df["OEE"])))
    
#     @QtCore.pyqtSlot()
#     def update_data_frame(self, item):
#         if item.row() < 0 or item.column() < 0 or self.Flag_view_type == False:
#             return
#         row = item.row()
#         if self.ui.table_widget.item(row, 8) is None:
#             return
#         try:
#             runtime = float(self.ui.table_widget.item(row, 2).text()) - float(self.ui.table_widget.item(row, 3).text()) - float(self.ui.table_widget.item(row, 4).text())
#             self.ui.table_widget.item(row, 5).setText(str(round(runtime, 2)))

#             A_cell = float(self.ui.table_widget.item(row, 5).text()) / ((float(self.ui.table_widget.item(row, 2).text())) + 0.000001)
#             self.ui.table_widget.item(row, 8).setText(str(round(A_cell, 2)))

#             cycle_time = self.machine_info[
#                 (self.machine_info["line_name"] == self.ui.line_cbb.currentText()) &
#                 (self.machine_info["machine_name"] == self.ui.machine_cbb.currentText())
#             ]["cycletime"].values[0]
#             P_cell = float(self.ui.table_widget.item(row, 6).text()) / (((60 / cycle_time) * float(self.ui.table_widget.item(row, 5).text())) + 0.000001)
#             self.ui.table_widget.item(row, 9).setText(str(round(P_cell, 2)))

#             Q_cell = float(self.ui.table_widget.item(row, 6).text()) / (float(self.ui.table_widget.item(row, 6).text()) + float(self.ui.table_widget.item(row, 7).text()) + 0.000001)
#             self.ui.table_widget.item(row, 10).setText(str(round(Q_cell, 2)))

#             OEE_cell = A_cell * P_cell * Q_cell
#             self.ui.table_widget.item(row, 11).setText(str(round(OEE_cell, 2)))
#             self.Flag_data_change = True
#         except Exception as e:
#             QtWidgets.QMessageBox.critical(self, "Error", f"Failed to update data: {str(e)}")
#         self.ui.save_btn.setEnabled(True)

#     def set_data_update(self):
#         if self.Flag_data_change:
#             self.data_update = []
#             for row in range(self.ui.table_widget.rowCount()):
#                 row_data = []
#                 for column in range(self.ui.table_widget.columnCount()):
#                     item = self.ui.table_widget.item(row, column)
#                     row_data.append(item.text() if item else None)
#                 self.data_update.append(row_data)
#             self.data_update = pd.DataFrame(self.data_update, columns=[
#                 "Day", "Working_shift", "Net_avaiable_runtime", "Downtime", "Other_stop", "Runtime", "FGs", "Defect", "A", "P", "Q", "OEE"])
#             self.data_update = self.data_update.set_index("Day")
#             self.data_update = self.data_update.astype(float)
#             self.update_chart(line=None,df = self.data_update)
#         else:
#             self.update_chart(line=None)

#     @QtCore.pyqtSlot()
#     def save_data_frame(self):
#         self.set_data_update()
#         if ( self.data_update is not None) or ( not self.data_update.empty ):
#             try:
#                 if self.ui.machine_cbb.currentText() == "Molding":
#                     self.parent.list_df[0][self.line] = (self.parent.list_df[0][self.line][0], self.data_update)
#                 else:
#                     self.parent.list_df[1][self.line] = (self.parent.list_df[1][self.line][0], self.data_update)
#                 QtWidgets.QMessageBox.information(self, "Success", "Dữ liệu đã được lưu thành công.")
#             except Exception as e:
#                 QtWidgets.QMessageBox.critical(self, "Error", f"Failed to save data: {str(e)}")
#         else:
#             QtWidgets.QMessageBox.warning(self, "Warning", "Không có dữ liệu để lưu.")
#         self.ui.save_btn.setEnabled(False)
#         self.data_update = None
        


#     def closeEvent(self, event):
#         if (len(self.machine_code_adjust) == 0) and (len(self.line_adjust) == 0):
#             if self.webview is not None:
#                 self.webview.deleteLater()
#             event.accept()
#         else:
#             reply = QtWidgets.QMessageBox.question(
#                 self,
#                 "Xác nhận",
#                 "Bạn có chắc muốn lưu mã máy được chỉnh sửa không?",
#                 QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
#                 QtWidgets.QMessageBox.No
#             )
#             if reply == QtWidgets.QMessageBox.Yes:
#                 self.parent.receivers(machine_code_adjust=self.machine_code_adjust,
#                                       machine_type_adjust=self.machine_type_adjust, line_adjust=self.line_adjust)
#                 if self.webview is not None:
#                     self.webview.deleteLater()
#                 event.accept()
    
#==========================Function of OEE page ==================================================================================END
#==========================Function of OEE page ==================================================================================END
#==========================Function of OEE page ==================================================================================END

class Machine_information(QtWidgets.QDialog):
    def __init__(self, database=None, code = None):
        super().__init__()
        self.database = database
        self.ui = Ui_Machine_detail()
        self.ui.setupUi(self)
        self.ui.mc_code.setText(code)
        try:
            data = self.database.query(sql = '''
                                                SELECT m.machine_name,p.line_name,d.department_name,m.maker,m.model,m.function,m.date_receipt,m.machine_status,m.image_link
                                                FROM `Machines` as m
                                                JOIN `Production_Lines` as p
                                                ON m.line_id = p.line_id
                                                JOIN `Departments` as d
                                                ON p.department_id = d.department_id
                                                WHERE m.machine_code = :code;
                                                ''',params = {'code':code}) 
            self.ui.mc_name.setText(data[0][0])
            self.ui.mc_marker.setText(data[0][3])
            self.ui.mc_model.setText(data[0][4])
            self.ui.mc_func.setText(data[0][5])
            self.ui.mc_date_receipt.setText(data[0][6])
            self.ui.mc_department.setText(data[0][2])
            self.ui.mc_line.setText(data[0][1])
            self.ui.mc_status.setText(data[0][7])
            pixmap = QtGui.QPixmap(data[0][8])
            if pixmap:
                self.ui.mc_pic.setPixmap(pixmap)
            headers = ["Date","Line","Result","Record"]
            monitor_model = QtGui.QStandardItemModel()
            monitor_model.setHorizontalHeaderLabels(headers)
            history = self.database.query(sql ='''SELECT mr.maintenance_date, p.line_name,mr.record_link
                                                    FROM `Maintenance_records` as mr
                                                    JOIN `Production_Lines` as p
                                                    ON mr.line_id = p.line_id
                                                    JOIN `Machines` as m
                                                    ON mr.machine_id = m.machine_id 
                                                    WHERE m.machine_code = :code
                                                    ORDER BY mr.maintenance_date DESC;
                                                    ''',params = {'code':code}) 
            for row in history:
                item = [QtGui.QStandardItem(str(row[0])),QtGui.QStandardItem(str(row[1])),QtGui.QStandardItem("OK")]
                monitor_model.appendRow(item)
            self.ui.mc_history.setModel(monitor_model)
            self.ui.mc_history.setColumnWidth(0,80)
            self.ui.mc_history.setColumnWidth(1,60)
            self.ui.mc_history.setColumnWidth(2,50)
            self.ui.mc_history.setColumnWidth(3,50)
            delegate_btn = ButtonDelegate(buttons=("Link",))
            self.ui.mc_history.setItemDelegateForColumn(3, delegate_btn)
            delegate_btn.ButtonClicked.connect(lambda name, idx : self.on_delegate_btn_clicked(name,idx,history))
            self.ui.mc_history.setMouseTracking(True)
            self.ui.mc_history.viewport().setMouseTracking(True) 
            self.ui.mc_history.resizeRowsToContents()
            headers2 = ["Code","Name","Stock","Safety"]
            monitor_model2 = QtGui.QStandardItemModel()
            monitor_model2.setHorizontalHeaderLabels(headers2)
            # part_list = self.database.query(sql ='''SELECT p.part_code, p.part_name
            #                                                 FROM `Machine_Partlist` AS mp
            #                                                 JOIN `Part_code` as p
            #                                                 ON mp.part_id = p.part_id
            #                                                 JOIN `Machines` as m
            #                                                 ON mp.machine_id = m.machine_id
            #                                                 WHERE m.machine_code = :code;
            #                                         ''',params = {'code':code}) 
            # for row in part_list:
            #     item = [QtGui.QStandardItem(str(row[0])),QtGui.QStandardItem(str(row[1])),QtGui.QStandardItem("1O"),QtGui.QStandardItem("1O")]
            #     monitor_model2.appendRow(item)
            # self.ui.mc_partlist.setModel(monitor_model2)
            self.ui.mc_partlist.setColumnWidth(0,80)
            self.ui.mc_partlist.setColumnWidth(1,100)
            self.ui.mc_partlist.setColumnWidth(2,40)
            self.ui.mc_partlist.setColumnWidth(3,40)
            self.ui.mc_partlist.resizeRowsToContents()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
    
    @QtCore.pyqtSlot()
    def on_delegate_btn_clicked(self,name,index,history):
        row = index.row()
        try:
            self.pdf = pdf_view(history[row][2])
            self.pdf.show()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"File not found: {history[row][2]}")

class Update_machine_info(QtWidgets.QDialog):
    def __init__(self, parent = None, code = None):
        super().__init__()
        self.parent = parent
        self.database = self.parent.database_process
        self.code = code
        self.ui = Ui_Update_machine_info()
        self.ui.setupUi(self)
        self.delete_data = []
        hearders = ["Code","Machine name","Group","Line name","Maintenance Frequency","Maker","Model","Function","Date receipt","Machine status","Image Link"]
        self.ui.machine_info_table_bf.setRowCount(len(hearders))
        self.ui.machine_info_table_af.setRowCount(len(hearders))
        self.ui.machine_info_table_af.setColumnCount(2)
        self.ui.machine_info_table_bf.setColumnCount(2)
        for r,hearder in enumerate( hearders, start= 0 ):
            item = QtWidgets.QTableWidgetItem(hearder)
            item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.ui.machine_info_table_bf.setItem(r,0,QtWidgets.QTableWidgetItem(hearder))
            self.ui.machine_info_table_af.setItem(r,0,item)
        self.ui.machine_info_table_bf.verticalHeader().setVisible(False)
        self.ui.machine_info_table_bf.horizontalHeader().setVisible(False)
        self.ui.machine_info_table_bf.setItem(0,1,QtWidgets.QTableWidgetItem(self.code))
        self.ui.machine_info_table_bf.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.machine_info_table_bf.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.machine_info_table_af.setItem(0,1,QtWidgets.QTableWidgetItem(self.code))
        self.ui.machine_info_table_af.horizontalHeader().setVisible(False)
        self.ui.machine_info_table_af.verticalHeader().setVisible(False)
        self.ui.machine_info_table_af.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.maintenance_plan_table_bf.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.maintenance_plan_table_bf.verticalHeader().setVisible(False)
        self.ui.maintenance_plan_table_bf.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.maintenance_plan_table_af.verticalHeader().setVisible(False)
        self.ui.maintenance_plan_table_af.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        try:
            self.machine_info = self.database.query(sql = '''   SELECT m.machine_name,d.department_name,p.line_name,m.maintenance_frequency,m.maker,m.model,m.function,m.date_receipt,m.machine_status,m.image_link
                                                                FROM `Machines` as m
                                                                JOIN `Production_Lines` as p
                                                                ON m.line_id = p.line_id
                                                                JOIN `Departments` as d
                                                                ON p.department_id = d.department_id
                                                                WHERE m.machine_code = :code;
                                                                ''',params = {'code':self.code})
            self.maintenance_plan = self.database.query(sql = '''SELECT my.month,mp.week,p.line_name
                                                                FROM `Maintenance_plan` as mp
                                                                JOIN `Production_Lines` as p
                                                                ON mp.line_id = p.line_id
                                                                JOIN `Machines` as m
                                                                ON mp.machine_id = m.machine_id
                                                                JOIN `Months_Years` as my
                                                                ON mp.month_year_id = my.month_year_id
                                                                WHERE m.machine_code = :code AND my.year = :year AND mp.maintenance_date IS NULL AND mp.status is NULL
                                                                GROUP BY my.month;
                                                                ''',params = {'code':self.code,'year':self.parent.year_num})
            self.register_form = self.database.query(sql = '''  SELECT mf.form_name,mf.form_link, d.department_name
                                                                FROM `Maintenance_Form_Register` as mfr
                                                                JOIN `Maintenance_form` as mf
                                                                ON mfr.form_id = mf.form_id
                                                                JOIN `Machines` as m
                                                                ON mfr.machine_id = m.machine_id
                                                                JOIN `Departments` as d
                                                                ON mf.department_id = d.department_id
                                                                WHERE m.machine_code = :code;
                                                            ''',params = {'code':self.code})
            for r,item in enumerate( self.machine_info[0], start= 1 ):
                self.ui.machine_info_table_bf.setItem(r,1,QtWidgets.QTableWidgetItem(str(item)))
            self.ui.maintenance_plan_table_bf.setRowCount(len(self.maintenance_plan))
            for row in range(len(self.maintenance_plan)):
                for col in range(len(self.maintenance_plan[row])):
                    self.ui.maintenance_plan_table_bf.setItem(row,col,QtWidgets.QTableWidgetItem(str(self.maintenance_plan[row][col])))
            if self.register_form:
                self.ui.form_type_lnedit_bf.setText(f"{self.register_form[0][0]} : {self.register_form[0][2]}")
            self.ui.form_type_lnedit_bf.setEnabled(False)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
        self.setup_signals()
    
    def setup_signals(self):
        self.ui.insert_btn.clicked.connect(self.insert_row)
        self.ui.delete_btn.clicked.connect(self.delete_row)
        self.ui.cancel_btn.clicked.connect(self.close)
        self.ui.transfer_btn.clicked.connect(self.transfer_data)
        self.ui.confirm_btn.clicked.connect(self.update_data)
        self.ui.form_type_lnedit_af.textChanged.connect(lambda text : self.on_text_changed(text=text))
        self.ui.maintenance_plan_table_af.itemChanged.connect(lambda item : self.check_line(item))
    
    @QtCore.pyqtSlot()
    def insert_row(self):
        current_row = self.ui.maintenance_plan_table_af.rowCount()
        if current_row < 12:
            self.ui.maintenance_plan_table_af.insertRow(current_row)
    
    @QtCore.pyqtSlot()
    def delete_row(self):
        current_row = self.ui.maintenance_plan_table_af.currentRow()
        try:
            if self.ui.maintenance_plan_table_af.item(current_row,0) is not None:
                code = self.ui.machine_info_table_bf.item(0,1).text()
                week = self.ui.maintenance_plan_table_bf.item(current_row,1).text()
                question = QtWidgets.QMessageBox.question(self,"Delete",f"Are you sure to delete the maintenance plan for the machine '{code}'?",QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,QtWidgets.QMessageBox.No)
                if question == QtWidgets.QMessageBox.Yes:
                    self.database.query(sql = ''' DELETE mp
                                                    FROM `Maintenance_plan` mp
                                                    JOIN `Machines` m
                                                    ON mp.machine_id = m.machine_id
                                                    JOIN `Production_Lines` p
                                                    ON mp.line_id = p.line_id
                                                    JOIN `Months_Years` AS my
                                                    ON mp.month_year_id = my.month_year_id
                                                    WHERE m.machine_code = :code AND my.year = :year AND mp.week = :week;''',
                                                params = {'code':code,'week':week,'year':self.parent.year_num})
                    self.maintenance_plan = self.database.query(sql = '''SELECT my.month,mp.week,p.line_name
                                                FROM `Maintenance_plan` as mp
                                                JOIN `Production_Lines` as p
                                                ON mp.line_id = p.line_id
                                                JOIN `Machines` as m
                                                ON mp.machine_id = m.machine_id
                                                JOIN `Months_Years` as my
                                                ON mp.month_year_id = my.month_year_id
                                                WHERE m.machine_code = :code AND my.year = :year AND mp.maintenance_date IS NULL
                                                GROUP BY my.month;
                                                ''',params = {'code':code,'year':self.parent.year_num})
                    self.ui.maintenance_plan_table_bf.clearContents()
                    self.ui.maintenance_plan_table_bf.setRowCount(len(self.maintenance_plan))
                    for row in range(len(self.maintenance_plan)):
                        for col in range(len(self.maintenance_plan[row])):
                            self.ui.maintenance_plan_table_bf.setItem(row,col,QtWidgets.QTableWidgetItem(str(self.maintenance_plan[row][col])))
                    self.ui.maintenance_plan_table_af.removeRow(current_row)
        except:
            pass
    
    @QtCore.pyqtSlot()
    def transfer_data(self):
        self.ui.machine_info_table_af.setItem(0,1,QtWidgets.QTableWidgetItem(self.code))
        for r,item in enumerate( self.machine_info[0], start= 1 ):
            self.ui.machine_info_table_af.setItem(r,1,QtWidgets.QTableWidgetItem(str(item)))
        self.ui.maintenance_plan_table_af.setRowCount(len(self.maintenance_plan))
        for row in range(len(self.maintenance_plan)):
            for col in range(len(self.maintenance_plan[row])):
                self.ui.maintenance_plan_table_af.setItem(row,col,QtWidgets.QTableWidgetItem(str(self.maintenance_plan[row][col])))
        if self.register_form:
                self.ui.form_type_lnedit_af.setText(f"{self.register_form[0][0]} : {self.register_form[0][2]}")
    
    @QtCore.pyqtSlot() 
    def on_text_changed(self,text):
        try:
            dep = self.ui.machine_info_table_af.item(2,1).text()
            self.parent.filter_suggestion(self.ui.form_type_lnedit_af,"mf.form_name, d.department_name","`Maintenance_form` as mf ", f'''JOIN `Departments` as d
                                                                                                                                                ON d.department_id = mf.department_id
                                                                                                                                                WHERE mf.form_name LIKE "%{text}%" AND d.department_name = "{dep}"
                                                                                                                                                ''')
        except:
            pass
    
    @QtCore.pyqtSlot()
    def update_data(self):
        def to_null(value):
            if value is None:
                return None
            v = str(value).strip()
            return None if v.lower() in ("none", "", "null") else v
        try:
            ui = self.ui.machine_info_table_af
            new_code = to_null(ui.item(0,1).text())
            old_code = to_null(self.ui.machine_info_table_bf.item(0,1).text())
            name = to_null(ui.item(1,1).text())
            dep = to_null(ui.item(2,1).text())
            line = to_null(ui.item(3,1).text())
            freq = to_null(ui.item(4,1).text())
            maker = to_null(ui.item(5,1).text())
            model = to_null(ui.item(6,1).text())
            function = to_null(ui.item(7,1).text())
            receipt = to_null(ui.item(8,1).text())
            status = to_null(ui.item(9,1).text())
            image = to_null(ui.item(10,1).text())
            iscorrectDep = self.database.query(sql =''' SELECT 1 FROM `Production_Lines` as p
                                                        JOIN `Departments` as d
                                                        ON p.department_id = d.department_id
                                                        WHERE p.line_name = :line AND d.department_name = :dep;
                                                        ''',params = {'line':line,'dep':dep})
            if not iscorrectDep:
                QtWidgets.QMessageBox.critical(self, "Error", f"Line name not found in your Group")
                return
            form = self.ui.form_type_lnedit_af.text().split(" : ")[0]
            iscorrectForm = self.database.query(sql =''' SELECT 1 FROM `Maintenance_form` as mf
                                                        JOIN `Departments` as d
                                                        ON mf.department_id = d.department_id
                                                        WHERE mf.form_name = :form AND d.department_name = :dep;
                                                        ''',params = {'form':form,'dep':dep})
            if not iscorrectForm:
                QtWidgets.QMessageBox.critical(self, "Error", f"Form name not found in your Group")
                return
            if dep != self.machine_info[0][1]:
                maintenance_plans = []
                for row in range(self.ui.maintenance_plan_table_af.rowCount()):
                    if self.ui.maintenance_plan_table_af.item(row,0) is not None:
                        maintenance_plans.append((self.ui.maintenance_plan_table_af.item(row,0).text(),self.ui.maintenance_plan_table_af.item(row,1).text(),self.ui.maintenance_plan_table_af.item(row,2).text()))
                payload = {'old_code':old_code,'name':name,'department': dep,'line':line,'freq':freq, 
                           'maker':maker, 'model':model, 'function':function,'receipt':receipt,'status':status,'image':image,
                           'maintenance':maintenance_plans,'form':form}
                receiver_id = self.database.query(sql = ''' SELECT u.user_id FROM `Users` as u
                                                            JOIN `Departments` as d
                                                            ON u.department_id = d.department_id
                                                            JOIN `Roles` as r
                                                            ON u.role_id = r.role_id
                                                            WHERE d.department_name = :dep AND r.role_level = "Supervisor";''',params = {'dep':dep})
                data = {
                    'type': 'update_machine',
                    'sender_id': self.parent.login_info['user_id'],
                    'receiver_id': receiver_id[0][0],
                    'title': "Change department of machine request",
                    'message': f"User {self.parent.login_info['user_name']} has requested to change the department of machine '{old_code}'",
                    'payload': payload,
                    'status': None,
                    'priority': None,
                    'expires_at': None,
                    'related_task_id': None
                }
                self.parent.send_notification(data)
                QtWidgets.QMessageBox.information(self,"Request Sent","Your request has been sent to the Supervisor, please wait for confirmation.")
            else:
                for row in range(self.ui.maintenance_plan_table_af.rowCount()):
                    if row < self.ui.maintenance_plan_table_bf.rowCount():
                        self.database.query(sql = '''   UPDATE `Maintenance_plan` as mp
                                                        JOIN `Machines` as m
                                                        ON mp.machine_id = m.machine_id
                                                        JOIN `Months_Years` as my
                                                        ON mp.month_year_id = my.month_year_id
                                                        SET 
                                                            mp.line_id =  ( SELECT p.line_id FROM  `Production_Lines` as p
                                                                            WHERE p.line_name = :line ),
                                                            mp.month_year_id = (
                                                                SELECT my2.month_year_id FROM `Months_Years` as my2 
                                                                WHERE my2.month = get_working_week_month(:year,:week) AND my2.year = :year),
                                                            mp.week = :week
                                                        WHERE m.machine_code = :old_code AND my.month = :old_month AND my.year = :year;
                                                    ''',params = {'old_code':old_code,'week':self.ui.maintenance_plan_table_af.item(row,1).text(),
                                                                'line':self.ui.maintenance_plan_table_af.item(row,2).text(),'old_month':self.ui.maintenance_plan_table_bf.item(row,0).text(),'year':self.parent.year_num})
                    else:
                        self.database.query(sql = '''   INSERT INTO `Maintenance_plan` 
                                                            (machine_id, line_id, month_year_id, quarter, week, original_week)
                                                        SELECT 
                                                            m.machine_id,
                                                            (SELECT p.line_id FROM `Production_Lines` AS p WHERE p.line_name = :line LIMIT 1),
                                                            (SELECT my.month_year_id 
                                                            FROM `Months_Years` AS my 
                                                            WHERE my.month = get_working_week_month(:year, :week)
                                                            AND my.year = :year
                                                            LIMIT 1),
                                                            :quarter,
                                                            :week,
                                                            :original_week
                                                        FROM `Machines` AS m
                                                        WHERE m.machine_code = :code;
                                                    ''',params = {'code':old_code,'line':self.ui.maintenance_plan_table_af.item(row,2).text(),'quarter':(self.parent.ui.company_week_month(self.parent.year_num,int(self.ui.maintenance_plan_table_af.item(row,1).text())) - 1) // 3 + 1,
                                                                  'week':self.ui.maintenance_plan_table_af.item(row,1).text(),'original_week':self.ui.maintenance_plan_table_af.item(row,1).text(),'year':self.parent.year_num})
                self.database.query(sql = '''   UPDATE `Maintenance_Form_Register` AS mfr
                                                SET mfr.form_id = ( SELECT mf.form_id FROM `Maintenance_form` AS mf
                                                                    JOIN `Departments` AS d
                                                                    ON mf.department_id = d.department_id
                                                                    WHERE mf.form_name = :form AND d.department_name = :dep LIMIT 1)
                                                WHERE mfr.machine_id = ( SELECT m2.machine_id FROM `Machines` AS m2 WHERE m2.machine_code = :code LIMIT 1);
                                    ''',params = {'code':old_code,'form':form,'dep':dep})

                self.database.query(sql = '''   UPDATE `Machines` AS m
                                                SET 
                                                    m.machine_code = :new_code,
                                                    m.machine_name = :name,
                                                    m.line_id = (
                                                        SELECT p2.line_id FROM `Production_Lines` AS p2 WHERE p2.line_name = :line
                                                    ),
                                                    m.maintenance_frequency = :freq,
                                                    m.maker = :maker,
                                                    m.model = :model,
                                                    m.function = :function,
                                                    m.date_receipt = :receipt,
                                                    m.machine_status = :status,
                                                    m.image_link = :image
                                                WHERE m.machine_code = :old_code;
                                                ''',params = {'old_code':old_code,'new_code':new_code,'name':name,'dep':dep,
                                                            'line':line,'freq':freq, 'maker':maker, 'model':model, 'function':function,
                                                                'receipt':receipt, 'status': status, 'image':image})
                QtWidgets.QMessageBox.information(self,"Update success","The machine data has been updated successfully.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to update data: {e}")
    
    @QtCore.pyqtSlot()
    def check_line(self,item):
        dep = self.ui.machine_info_table_af.item(2,1).text()
        col = item.column()
        row = item.row()
        line = item.text()
        if col == 2:
            isCorrectDep = self.database.query(sql =''' SELECT 1 FROM `Production_Lines` as p
                                                        JOIN `Departments` as d
                                                        ON p.department_id = d.department_id
                                                        WHERE p.line_name = :line AND d.department_name = :dep;
                                                        ''',params = {'line':line,'dep':dep})
            if not isCorrectDep:
                self.ui.maintenance_plan_table_af.itemChanged.disconnect()
                self.ui.maintenance_plan_table_af.setItem(row,col,QtWidgets.QTableWidgetItem(""))
                self.ui.maintenance_plan_table_af.itemChanged.connect(lambda item: self.check_line(item = item))
                QtWidgets.QMessageBox.critical(self, "Error", f"Line name not found in your Group")
                return

    def closeEvent(self, event):
        self.ui.machine_info_table_af.clearContents()
        self.ui.machine_info_table_bf.clearContents()
        self.ui.maintenance_plan_table_af.clearContents()
        self.ui.maintenance_plan_table_bf.clearContents()
        super().close()   
        self.deleteLater()                    

class PageRenderWorker(QtCore.QThread):
    rendered = QtCore.pyqtSignal(int, QtGui.QPixmap)

    def __init__(self, doc, page_num, zoom):
        super().__init__()
        self.doc = doc
        self.page_num = page_num
        self.zoom = zoom

    def run(self):
        try:
            page = self.doc.load_page(self.page_num)
            mat = fitz.Matrix(self.zoom, self.zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = QtGui.QImage(pix.samples, pix.width, pix.height, pix.stride,
                               QtGui.QImage.Format_RGB888)
            self.rendered.emit(self.page_num, QtGui.QPixmap.fromImage(img))
        except Exception:
            pass

class pdf_view(QtWidgets.QGraphicsView):
    def __init__(self, pdf_path):
        super().__init__()
        self.setWindowTitle("PDF Viewer")
        self.resize(1200, 900)
        self.setRenderHints(QtGui.QPainter.Antialiasing |
                            QtGui.QPainter.SmoothPixmapTransform)
        self.setDragMode(QtWidgets.QGraphicsView.ScrollHandDrag)

        self.dpi = 140                 
        self.zoom = self.dpi / 72
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
            return
        self.doc = fitz.open(pdf_path)

        self.scene = QtWidgets.QGraphicsScene()
        self.setScene(self.scene)

        self.page_cache = {}
        self.page_items = {}
        self.loading_threads = {}

        self.load_initial_pages()

    def load_initial_pages(self):
        for page_num in range(min(3, self.doc.page_count)):
            self.load_page(page_num)

    def load_page(self, page_num):
        if page_num in self.page_cache or page_num in self.loading_threads:
            return

        worker = PageRenderWorker(self.doc, page_num, self.zoom)
        worker.rendered.connect(self.insert_page)
        worker.start()

        self.loading_threads[page_num] = worker

    def insert_page(self, page_num, pixmap):
        self.page_cache[page_num] = pixmap
        item = self.scene.addPixmap(pixmap)
        item.setPos(0, page_num * (pixmap.height() + 40))
        self.page_items[page_num] = item
        self.scene.setSceneRect(self.scene.itemsBoundingRect())

        del self.loading_threads[page_num]

    def wheelEvent(self, event):
        if event.modifiers() & QtCore.Qt.ControlModifier:
            delta = event.angleDelta().y()
            self.zoom *= 1.1 if delta > 0 else 1/1.1
            self.zoom = max(0.4, min(self.zoom, 4))
            self.resetTransform()
            self.scale(self.zoom, self.zoom)

        else:
            super().wheelEvent(event)

        if self.verticalScrollBar().value() > self.verticalScrollBar().maximum() - 300:
            next_page = len(self.page_cache)
            if next_page < self.doc.page_count:
                self.load_page(next_page)

class StatusColorDelegate(QtWidgets.QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        if index.column() == 5:
            status = index.data()
            if status == "Overdue":
                option.palette.setColor(QtGui.QPalette.Text, QtGui.QColor("#ff0000"))
            elif status == "Upcoming":
                option.palette.setColor(QtGui.QPalette.Text, QtGui.QColor("#29cc00"))
            elif status == "Near due":
                option.palette.setColor(QtGui.QPalette.Text, QtGui.QColor("#b04903"))

class ButtonDelegate(QtWidgets.QStyledItemDelegate):
    ButtonClicked = QtCore.pyqtSignal(str, QtCore.QModelIndex)

    def __init__(self, buttons=("Detail", "Update"), parent=None):
        super().__init__(parent)
        self._buttons = {}
        self._hovered = None
        self._button_names = buttons  # tuple các nút muốn tạo

    def paint(self, painter, option, index):
        super().paint(painter, option, index)

        rect = option.rect
        count = len(self._button_names)
        if count == 0:
            return

        # chia đều chiều ngang cho các nút
        w = rect.width() // count - (count + 1)
        h = rect.height() - 9

        btn_rects = {}
        for i, name in enumerate(self._button_names):
            x = rect.left() + 1 + i * (w + 5)
            y = rect.top() + 4
            r = QtCore.QRect(x, y, w, h)
            btn_rects[name] = r

            self.drawButton(
                painter, r, name,
                bg="#FFFFFF", hover="#ff6600", text_color="black",
                hovered=(self._hovered == (name, index))
            )

        self._buttons[index] = btn_rects

    def drawButton(self, painter, rect, text, bg, hover, text_color, hovered=False):
        painter.save()
        painter.setBrush(QtGui.QColor(bg))
        border_color = QtGui.QColor(hover if hovered else "#ffffff")
        painter.setPen(QtGui.QPen(border_color, 1))
        painter.drawRoundedRect(rect, 2, 2)

        painter.setPen(QtGui.QColor(text_color))
        painter.setFont(QtGui.QFont("Arial", 8, QtGui.QFont.Bold))
        painter.drawText(rect, QtCore.Qt.AlignCenter, text)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QtCore.QEvent.MouseMove:
            pos = event.pos()
            btns = self._buttons.get(index, {})
            for name, rect in btns.items():
                if rect.contains(pos):
                    if self._hovered != (name, index):
                        self._hovered = (name, index)
                        option.widget.viewport().update()
                    return True
            if self._hovered:
                self._hovered = None
                option.widget.viewport().update()
            return True

        elif event.type() == QtCore.QEvent.MouseButtonRelease:
            pos = event.pos()
            btns = self._buttons.get(index, {})
            for name, rect in btns.items():
                if rect.contains(pos):
                    self.ButtonClicked.emit(name, index)
                    return True

        return super().editorEvent(event, model, option, index)  

class Print_selector(QtWidgets.QWidget):
    def __init__(self, parent=None,quantity = 0,data = None, attached_machine = None,database = None, duplicate = None):
        super().__init__(parent)
        self.ui = Ui_print_selector()
        self.ui.setupUi(self)
        self.setup_signals()
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Window)
        self.printer_process = Printer_process()
        self.printer_name = [name[2] for name in self.printer_process.printers_list ]
        self.ui.Print_select_cbb.addItems(self.printer_name)
        self.quantity = quantity
        self.data = data
        self.database = database
        self.duplicate = duplicate
        self.attached_machine = attached_machine
        self.ui.label_4.setText(f"{self.quantity} forms")
    
    def setup_signals(self):
        self.ui.Print_confirm_bt.clicked.connect(self.start_printer)
        self.ui.Print_cancel_bt.clicked.connect(self.close)
        self.ui.Print_select_cbb.currentIndexChanged.connect(self.select_printer)
    
    @QtCore.pyqtSlot()
    def select_printer(self):
        self.printer_process.choice_printer(self.ui.Print_select_cbb.currentText())
    
    @QtCore.pyqtSlot()
    def start_printer(self):
        if self.quantity <= 0 or self.data == None:
            QtWidgets.QMessageBox.critical(self,"Error", f"Nothing to print")
            return
        try:
            self.worker = WorkerThread(self.print_job, self.data)
            self.progress_window = Printer_progress(max = len(self.data),worker= self.worker)
            self.worker.progress_changed.connect(lambda value: self.progress_window.update_progress(value=value))
            self.worker.finished.connect(self.progress_window.on_finished)
            self.progress_window.show()
            self.worker.start()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Print Error: {e}")
        self.close()

    def print_job(self, data):
        total = len(data)
        for i, info in enumerate(data, start=0):
            if info[0] in self.attached_machine:
                self.printer_process.send_to_printer(input_pdf = info[-1], data = [info[0], info[1],info[3], info[4], info[8], str(info[7])],attached_machine = self.attached_machine[info[0]],file_index= i)
            else:
                self.printer_process.send_to_printer(input_pdf =  info[-1], data = [info[0], info[1],info[3], info[4], info[8], str(info[7])],file_index= i)
            if i not in self.duplicate:
                try:
                    self.database.query(sql = f''' INSERT INTO `Record_pending` (machine_id,line_id,technical,maintenance_date)
                                                    VALUES ( (SELECT machine_id FROM `Machines` WHERE machine_code = :code),
                                                            (SELECT line_id FROM `Production_Lines` WHERE line_name = :line ),
                                                            :technical , :date );''', 
                                                            params = {'code':info[0],'line':info[4],'technical':info[8],'date':info[7]})
                    if info[0] in self.attached_machine:
                        for attached_code in self.attached_machine[info[0]]:
                            self.database.query(sql = f''' INSERT INTO `Record_pending` (machine_id,line_id,technical,maintenance_date,attached_equipment)
                                                            VALUES ( (SELECT machine_id FROM `Machines` WHERE machine_code = :code),
                                                                (SELECT line_id FROM `Production_Lines` WHERE line_name = :line ),
                                                                :technical , :date ,
                                                                (SELECT machine_id FROM `Machines` WHERE machine_code = :attach_with));''', 
                                                                params = {'code':attached_code,'line':info[4],'technical':info[8],'date':info[7],'attach_with':info[0]})
                except Exception as e:
                    QtWidgets.QMessageBox.critical(self,"Error",  f"Fail to load : {e}")
            else:
                try:
                    temp = self.database.query(sql = f'''SELECT rp.rp_id,m.machine_code FROM `Record_pending` as rp
                                                                                        JOIN `Machines` as m
                                                                                        ON rp.machine_id = m.machine_id
                                                                                        WHERE rp.attached_equipment = (SELECT m2.machine_id FROM `Machines` as m2 WHERE m2.machine_code = "{info[0]}")
                                                                                        ORDER BY rp.rp_id ASC;''' )
                    if info[0] in self.attached_machine:
                        code_list = self.attached_machine[info[0]]
                        if temp:
                            attach_code_current = [code[1] for code in temp]
                            for code in code_list:
                                if code in attach_code_current:
                                    self.database.query(sql = f'''UPDATE Record_pending
                                                                        SET line_id = ( SELECT line_id 
                                                                        FROM `Production_Lines`
                                                                        WHERE line_name = "{info[4]}"),
                                                                        technical = "{info[8]}",
                                                                        maintenance_date = "{info[7]}"
                                                                        WHERE machine_id = ( SELECT machine_id FROM `Machines` WHERE machine_code = "{info[0]}" );''')

                                else:
                                    self.database.query(sql = f''' INSERT INTO `Record_pending` (machine_id,line_id,technical,maintenance_date,attached_equipment)
                                                                    VALUES ( (SELECT machine_id FROM `Machines` WHERE machine_code = :code),
                                                                    (SELECT line_id FROM `Production_Lines` WHERE line_name = :line ),
                                                                    :technical , :date ,
                                                                    (SELECT machine_id FROM `Machines` WHERE machine_code = :attach_with));''', 
                                                                    params = {'code':code,'line':info[4],'technical':info[8],'date':info[7],'attach_with':info[0]})
                            delete_code = [c for c in attach_code_current if c not in code_list]
                            if delete_code:
                                delete_ids = ','.join(f"'{x}'" for x in delete_code)
                                self.database.query(sql=f'''DELETE FROM `Record_pending` 
                                                            WHERE machine_id IN (SELECT machine_id FROM Machines WHERE machine_code IN ({delete_ids}))''')
                        else:
                            for code in code_list:
                                self.database.query(sql = f'''  INSERT INTO `Record_pending` (machine_id,line_id,technical,maintenance_date,attached_equipment)
                                                                VALUES ( (SELECT machine_id FROM `Machines` WHERE machine_code = :code),
                                                                (SELECT line_id FROM `Production_Lines` WHERE line_name = :line ),
                                                                :technical , :date ,
                                                                (SELECT machine_id FROM `Machines` WHERE machine_code = :attach_with));''', 
                                                                params = {'code':code,'line':info[4],'technical':info[8],'date':info[7],'attach_with':info[0]})
                    else:       
                        if temp:
                            self.database.query(sql = f'''  DELETE FROM `Record_pending` 
                                                            WHERE attached_equipment = (SELECT m.machine_id FROM `Machines` as m WHERE m.machine_code = "{info[0]}");''')
                except Exception as e:
                    QtWidgets.QMessageBox.critical(self,"Error",  f"Fail to load data: {e}")
                self.worker.progress_changed.emit(int(i+1))

class WorkerThread(QtCore.QThread):
    finished = QtCore.pyqtSignal(object)
    progress_changed = QtCore.pyqtSignal(int)
    result_ready = QtCore.pyqtSignal(int, list) 
    error = QtCore.pyqtSignal(str)
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self._is_running = True  
    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
            self.finished.emit(result)  
        except Exception as e:
            self.error.emit(str(e))

    def stop(self):
        self._is_running = False

class WorkerSignals(QtCore.QObject):
    finished = QtCore.pyqtSignal(object)
    progress_changed = QtCore.pyqtSignal(int)
    result_ready = QtCore.pyqtSignal(int, object)
    error = QtCore.pyqtSignal(str)

class Worker_Pool(QtCore.QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()
        self._is_running = True

    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
            self.signals.finished.emit(result)
        except Exception as e:
            self.signals.error.emit(str(e))
    def stop(self):
        self._is_running = False

class Printer_progress(QtWidgets.QWidget):
    def __init__(self, parent=None,max = 0,text = "printed",worker = None):
        super().__init__(parent)
        self.ui = Ui_printing_progress()
        self.ui.setupUi(self)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Window)
        self.ui.printing_progess.setRange(0,max)
        self.ui.printing_progess.setValue(0)
        self.text = text
        self.worker = worker
    
    @QtCore.pyqtSlot()  
    def update_progress(self, value):
        self.ui.printing_progess.setValue(value)
    
    @QtCore.pyqtSlot()  
    def on_finished(self):
        QtWidgets.QMessageBox.information(self, "Done", f"All files {self.text}!")
        self.close()

class Form_Modification(QtWidgets.QDialog):
    def __init__(self ,parent = None):
        super().__init__(parent)
        self.ui = Ui_Form_Modification()
        self.ui.setupUi(self)
        self.pdf_page = Scan_record_process()
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Window)
        self.parent = parent
        self.debounce_timers = {}
        self.setup_signals()
        self.result = []
        self.department_maintenance_form = None
        header = ["Machine code", "Machine name"]
        self.ui.apply_machine_table.setColumnCount(len(header))
        self.ui.apply_machine_table.setHorizontalHeaderLabels(header)
        self.ui.apply_machine_table.setAcceptDrops(True)
        self.ui.apply_machine_table.setDragDropMode(QtWidgets.QAbstractItemView.DropOnly)
        self.ui.apply_machine_table.dragEnterEvent = self.dragEnterEvent
        self.ui.apply_machine_table.dragMoveEvent = self.dragMoveEvent
        self.ui.apply_machine_table.dropEvent = self.dropEvent
        self.ui.apply_machine_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)


    def setup_signals(self):
        self.ui.register_form_btn.clicked.connect(lambda _: self.register_form_page())
        self.ui.update_form_btn.clicked.connect(lambda _: self.update_form_page())
        self.ui.mod_cancel_btn.clicked.connect(self.close)
        self.ui.mod_insert_btn.clicked.connect(lambda _: self.insert_row_case())
        self.ui.mode_delete_btn.clicked.connect(lambda _: self.delete_row())
        self.ui.mod_confirm_btn.clicked.connect(lambda _: self.confirm_action())
        self.ui.record_name_lnedit.textChanged.connect(lambda text: self.on_text_changed(text,isupdate=True,iscell=False))
        self.ui.load_record_form.clicked.connect(lambda _: self.load_record_info())
        self.ui.list_form_btn.clicked.connect(self.list_form_page)
        self.ui.list_form_load_btn.clicked.connect(self.load_form_data)

    @QtCore.pyqtSlot()
    def register_form_page(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.register_form_page)
        self.parent.style_button_with_shadow(button = (self.ui.register_form_btn,self.ui.update_form_btn,self.ui.list_form_btn))
        self.ui.apply_machine_table.setEnabled(True)
        self.ui.apply_machine_table.setRowCount(0)
        self.ui.apply_machine_table.clearContents()
        self.ui.label_6.setText("Apply machines")
        self.ui.apply_machine_table.setEditTriggers(QtWidgets.QAbstractItemView.AllEditTriggers)
        header = ["Machine code", "Machine name"]
        self.ui.apply_machine_table.setColumnCount(len(header))
        self.ui.apply_machine_table.setHorizontalHeaderLabels(header)
        self.ui.apply_machine_table.setColumnWidth(0,100)
        self.ui.apply_machine_table.setColumnWidth(1,310)
        self.ui.frame_14.show()
        if self.ui.register_group_cbb.count() == 0:
            self.ui.register_group_cbb.addItems([item[0] for item in self.parent.group])
            self.ui.register_group_cbb.setCurrentText(self.parent.login_info['department'])
        self.ui.register_group_cbb.currentTextChanged.connect(lambda text: setattr(self, "department_maintenance_form", text))
    
    @QtCore.pyqtSlot()
    def update_form_page(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.update_form_page)
        self.parent.style_button_with_shadow(button = (self.ui.update_form_btn,self.ui.register_form_btn,self.ui.list_form_btn))
        self.ui.apply_machine_table.clearContents()
        self.ui.apply_machine_table.setRowCount(0)
        self.ui.apply_machine_table.setEnabled(False)
        self.ui.label_6.setText("Apply machines")
        header = ["Machine code", "Machine name"]
        self.ui.apply_machine_table.setColumnCount(len(header))
        self.ui.apply_machine_table.setHorizontalHeaderLabels(header)
        self.ui.apply_machine_table.setColumnWidth(0,100)
        self.ui.apply_machine_table.setColumnWidth(1,310)
        self.ui.frame_14.show()
        self.ui.apply_machine_table.setEditTriggers(QtWidgets.QAbstractItemView.AllEditTriggers)
        self.ui.update_choice.setChecked(True)

    @QtCore.pyqtSlot()
    def insert_row_case(self):
        if self.ui.stackedWidget.currentWidget() == self.ui.register_form_page:
            self.insert_row()
        else:
            self.insert_row(isupdate = True)

    def insert_row(self , isupdate = False):
        current_row = self.ui.apply_machine_table.rowCount()
        self.ui.apply_machine_table.insertRow(current_row)
        editor = QtWidgets.QLineEdit()
        editor.setStyleSheet(''' border: none;''')
        editor.textChanged.connect(
            lambda text, r=current_row, c=0: self.on_text_changed(text, c, r, isupdate)
        )
        editor.editingFinished.connect(lambda r=current_row: self.load_data(r,isupdate))
        self.ui.apply_machine_table.setCellWidget(current_row,0,editor)
    
    @QtCore.pyqtSlot()
    def delete_row(self,r = None):
        if r is None:
            current_row = self.ui.apply_machine_table.currentRow()
            self.ui.apply_machine_table.removeRow(self.ui.apply_machine_table.currentRow())

    @QtCore.pyqtSlot() 
    def on_text_changed(self,text,c = 0,r = 0,isupdate = False, iscell = True):
        try:
            if not isupdate:
                self.parent.filter_suggestion(self.ui.apply_machine_table.cellWidget(r,c),"m.machine_code","`Machines` as m ",f'''JOIN `Production_Lines` as p
                                                                                                                            ON p.line_id = m.line_id
                                                                                                                            JOIN `Departments` as d
                                                                                                                            ON d.department_id = p.department_id
                                                                                                                            LEFT JOIN `Maintenance_Form_Register` as mfr
                                                                                                                            ON m.machine_id = mfr.machine_id
                                                                                                                            WHERE d.department_name = "{self.ui.register_group_cbb.currentText()}" AND mfr.machine_id IS NULL AND machine_code LIKE "%{text}%"
                                                                                                                            ''')
            else:
                if iscell:
                    self.parent.filter_suggestion(self.ui.apply_machine_table.cellWidget(r,c),"m.machine_code","`Machines` as m ",f'''JOIN `Production_Lines` as p
                                                                                                                                ON p.line_id = m.line_id
                                                                                                                                JOIN `Departments` as d
                                                                                                                                ON d.department_id = p.department_id
                                                                                                                                LEFT JOIN `Maintenance_Form_Register` as mfr
                                                                                                                                ON m.machine_id = mfr.machine_id
                                                                                                                                WHERE d.department_name = "{self.department_update.strip()}" AND mfr.machine_id IS NULL AND machine_code LIKE "%{text}%"''')
                else:
                    self.parent.filter_suggestion(self.ui.record_name_lnedit,"mf.form_name, d.department_name","`Maintenance_form` as mf ", f'''  JOIN `Departments` as d
                                                                                                                                                            ON d.department_id = mf.department_id
                                                                                                                                                            WHERE mf.form_name LIKE "%{text}%"
                                                                                                                                                            ''')
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to load data: {e}")
    
    @QtCore.pyqtSlot()
    def load_data(self,r,isupdate):
        try:
            code = self.ui.apply_machine_table.cellWidget(r,0).text()
            if not isupdate: 
                result = self.parent.database_process.query(sql = '''SELECT m.machine_name 
                                                                    FROM `Machines` as m
                                                                    JOIN `Production_Lines` as p
                                                                    ON m.line_id = p.line_id
                                                                    JOIN `Departments` as d
                                                                    ON p.department_id = d.department_id
                                                                    WHERE m.machine_code = :code AND d.department_name = :dep;''',params = {'code':code,'dep':self.ui.register_group_cbb.currentText()})
                if r >  ( len(self.result) -1 ):
                    self.result.append(f"'{code}'")
                else:
                    self.result[r] = f"'{code}'"
            else:
                result = self.parent.database_process.query(sql = '''SELECT m.machine_name 
                                                                    FROM `Machines` as m
                                                                    JOIN `Production_Lines` as p
                                                                    ON m.line_id = p.line_id
                                                                    JOIN `Departments` as d
                                                                    ON p.department_id = d.department_id
                                                                    WHERE m.machine_code = :code AND d.department_name = :dep;''',params = {'code':code,'dep':self.department_update.strip()})
            self.ui.apply_machine_table.setItem(r,1,QtWidgets.QTableWidgetItem(f"{result[0][0]}"))
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to load data: {e}")
            if self.ui.apply_machine_table.cellWidget(r,0):
                editor = self.ui.apply_machine_table.cellWidget(r,0)
                editor.blockSignals(True)
                editor.setText("")
                editor.blockSignals(False)
                self.ui.apply_machine_table.setItem(r,1,QtWidgets.QTableWidgetItem(""))

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            file_path = event.mimeData().urls()[0].toLocalFile()
            try:
                if file_path.endswith(".csv"):
                    df = pd.read_csv(file_path)
                elif file_path.endswith(".xlsx"):
                    df = pd.read_excel(file_path)   
                else:
                    raise ValueError("File không phải là CSV hay XLSX")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self,"Error", f"Failed to load data: {e}")
                return
            if df.empty:
                QtWidgets.QMessageBox.critical(self,"Error", f"File is empty")
                return
            else:
                temp = self.parent.database_process.query(sql = ''' SELECT m.machine_code 
                                                                    FROM `Machines` as m
                                                                    JOIN `Production_Lines` as p
                                                                    ON m.line_id = p.line_id
                                                                    JOIN `Departments` as d
                                                                    ON p.department_id = d.department_id
                                                                    LEFT JOIN `Maintenance_Form_Register` as mfr
                                                                    ON m.machine_id = mfr.machine_id
                                                                    WHERE d.department_name = :dep AND mfr.machine_id IS NULL; ''',params = {'dep':self.ui.register_group_cbb.currentText()})
                non_register_machine = [machine[0] for machine in temp]
                machine_code = df.iloc[:, 0]
                machine_hasbeen_register = []
                for r,value in enumerate(machine_code,start= 0 ):
                    if value in non_register_machine:
                        self.insert_row()
                        self.ui.apply_machine_table.cellWidget(r,0).setText(str(value))
                        self.load_data(r = r,isupdate=False)
                        event.acceptProposedAction()
                    else:
                        machine_hasbeen_register.append(value)
                if len(machine_hasbeen_register) == 0:
                    return
                else:
                    QtWidgets.QMessageBox.information(self,
                                                    "Error",
                                                    f"Machine code:\n {'\n'.join(map(str, machine_hasbeen_register))} \n has been register or not in your department.",
                                                    QtWidgets.QMessageBox.StandardButton.Ok
                                                )
        else:
            super().dropEvent(event)
    
    @QtCore.pyqtSlot()
    def load_record_info(self):
        text = self.ui.record_name_lnedit.text()
        if ( text == "" ) or ( text is None ):
            return
        self.ui.apply_machine_table.setEnabled(True)
        try:
            self.department_update = text.split(":", 1)[1]
            text = text.split(":", 1)[0]
            self.department_maintenance_form = self.department_update.strip()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Invalid form name")
            return
        try:
            self.record_info = self.parent.database_process.query(sql = '''SELECT m.machine_code, m.machine_name,mfr.register_id
                                                        FROM `Machines` as m
                                                        JOIN `Maintenance_Form_Register` as mfr
                                                        ON m.machine_id = mfr.machine_id
                                                        JOIN `Maintenance_form` as mf
                                                        ON mfr.form_id = mf.form_id
                                                        WHERE mf.form_name =  :form_name;''',params = {'form_name' :text})
            self.form_info = self.parent.database_process.query(sql = ''' SELECT form_link,form_id
                                                                    FROM `maintenance_form`    
                                                                    WHERE form_name =  :form_name;''',params = {'form_name' :text})
            num_machine = len(self.record_info)
            self.machines_registered = [machine[0] for machine in self.record_info]
            self.ui.apply_machine_table.clearContents()
            self.ui.apply_machine_table.setRowCount(0)
            self.ui.update_form_link.setText(self.form_info[0][0])
            if num_machine == 0:
                raise ValueError(f"Not see any machine has been registered for the form {text}")
            else:
                for r in range(num_machine):
                    self.insert_row(isupdate=True)
                    editor = self.ui.apply_machine_table.cellWidget(r,0)
                    editor.setText(f"{self.record_info[r][0]}")
                    self.ui.apply_machine_table.setItem(r,1,QtWidgets.QTableWidgetItem(f"{self.record_info[r][1]}"))
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to load data: {e}")
    
    @QtCore.pyqtSlot()
    def confirm_action(self):
        if self.parent.login_info["role_level"] in ["Manager","Admin"]:
                pass
        elif ( self.parent.login_info["department"] == self.department_maintenance_form ) and ( self.parent.login_info["role_level"] in ["Supervisor"]):
            pass
        else:
            QtWidgets.QMessageBox.information(self,"Permission denied","Your don't have permission to update this machine info")
            return
        if self.ui.stackedWidget.currentWidget() == self.ui.list_form_page:
            return
        if self.ui.apply_machine_table.rowCount() == 0 :
            return
        new_codes = []
        for r in range(self.ui.apply_machine_table.rowCount()):
            editor = self.ui.apply_machine_table.cellWidget(r, 0)
            new_codes.append(editor.text())
        if len(new_codes) != len(set(new_codes)):
            QtWidgets.QMessageBox.critical(self,"Error", f"There are duplicate machine codes in the table")
            return
        if self.ui.stackedWidget.currentWidget() == self.ui.register_form_page:
            values = ",".join(self.result)
            try:
                form_name = self.ui.register_form_name.text()
                form_link = self.ui.register_form_link.text()
                page_num = self.pdf_page.return_form_page(form_link)
                if not form_name or not form_link or not form_link.lower().endswith(".pdf"):
                    QtWidgets.QMessageBox.warning(self, "Error", "Please enter a valid form name and .pdf path")
                    return
                department_id = self.parent.database_process.query(sql = ''' SELECT department_id FROM `Departments` WHERE department_name = :dep''', params = {'dep': self.ui.register_group_cbb.currentText() })
                self.parent.database_process.query(
                    sql = '''INSERT INTO `Maintenance_form` (form_name, form_link, department_id, page_num)
                            VALUES (:form_name, :form_link, :department_id , :num)''',
                    params = {
                        'form_name': form_name,
                        'form_link': form_link,
                        'department_id': department_id[0][0],
                        'num': page_num
                    }
                )
                self.parent.database_process.query ( sql = f''' INSERT INTO `Maintenance_Form_Register` (machine_id, form_id)
                                                                SELECT m.machine_id, f.form_id
                                                                FROM `Machines` AS m
                                                                JOIN `Maintenance_form` AS f 
                                                                ON f.form_name = :form_name
                                                                WHERE m.machine_code IN ({values});''',params = {'form_name':self.ui.register_form_name.text()})
            except Exception as e:
                QtWidgets.QMessageBox.critical(self,"Error", f"Failed to register form: {e}")
                return
        elif self.ui.stackedWidget.currentWidget() == self.ui.update_form_page:
            if self.ui.update_choice.isChecked():
                try:
                    text = self.ui.record_name_lnedit.text()
                    form_name = text.split(":", 1)[0]
                    form_link = self.ui.update_form_link.text()
                    self.parent.database_process.query( sql = f''' UPDATE `Maintenance_form`
                                                                    SET form_name = :form_name , form_link = :form_link , page_num = :num
                                                                    WHERE form_id = :form_id
                                                                ''', params = {'form_name':form_name, 'form_link':form_link,'num':self.pdf_page.return_form_page(form_link),'form_id':self.form_info[0][1]})
                    #delete
                    if len(self.record_info) > self.ui.apply_machine_table.rowCount():
                        old_codes = [row[0] for row in self.record_info] 
                        codes_to_delete = set(old_codes) - set(new_codes)
                        for old_code in codes_to_delete:
                            self.parent.database_process.query(
                                sql = ''' DELETE FROM `Maintenance_Form_Register`
                                        WHERE form_id = :form_id
                                        AND machine_id = (SELECT machine_id FROM `Machines` WHERE machine_code = :code) ''',
                                params = {
                                    'form_id': self.form_info[0][1],
                                    'code': old_code
                                }
                            )
                        QtWidgets.QMessageBox.information(self,"Action complete","Delete maintenance form complete")
                        return
                    #update and insert       
                    for r in range(self.ui.apply_machine_table.rowCount()):
                        editor = self.ui.apply_machine_table.cellWidget(r,0)
                        new_code = editor.text()
                        #insert
                        if r >= len(self.record_info):
                            self.parent.database_process.query ( sql = f''' INSERT INTO `Maintenance_Form_Register` (machine_id, form_id)
                                                                        SELECT m.machine_id, f.form_id
                                                                        FROM `Machines` AS m
                                                                        JOIN `Maintenance_form` AS f 
                                                                        ON f.form_id = :form_id
                                                                        WHERE m.machine_code = :code;''',params = {'form_id':self.form_info[0][1],'code' : new_code})
                            continue
                        #update
                        if self.record_info[r][0] != new_code:
                            self.parent.database_process.query ( sql = ''' UPDATE `Maintenance_Form_Register`
                                                                            SET machine_id = ( SELECT machine_id FROM `Machines` WHERE `machine_code` = :code)
                                                                            WHERE register_id = :register_id ''', params = {'code':new_code ,'register_id':self.record_info[r][-1]} )

                except Exception as e:
                    QtWidgets.QMessageBox.critical(self,"Error", f"Failed to update data: {e}")
                    return
            else:
                try:
                    self.parent.database_process.query ( sql = ''' DELETE FROM `Maintenance_form` 
                                                                    WHERE form_id = :form_id''', params = {'form_id':self.record_info[0][3]})
                except Exception as e:
                    QtWidgets.QMessageBox.critical(self,"Error", f"3Failed to update data: {e}")
                    return
        QtWidgets.QMessageBox.information(self,"Action complete","Update maintenance form complete")
    
    @QtCore.pyqtSlot()
    def list_form_page(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.list_form_page)
        self.parent.style_button_with_shadow(button = (self.ui.list_form_btn,self.ui.update_form_btn,self.ui.register_form_btn))
        self.ui.apply_machine_table.clearContents()
        self.ui.apply_machine_table.setRowCount(0)
        self.ui.label_6.setText("List of record form")
        self.ui.apply_machine_table.setEnabled(True)
        self.ui.apply_machine_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        header = ["Form name", "Machine apply\nQ'ty"]
        self.ui.apply_machine_table.setColumnCount(len(header))
        self.ui.apply_machine_table.setHorizontalHeaderLabels(header)
        self.ui.apply_machine_table.setColumnWidth(1,100)
        self.ui.apply_machine_table.setColumnWidth(0,280)
        self.ui.frame_14.hide()
        if self.ui.list_form_group_cbb.count() == 0:
            self.ui.list_form_group_cbb.addItems([item[0] for item in self.parent.group])
            self.ui.list_form_group_cbb.setCurrentText(self.parent.login_info['department'])
        self.ui.list_form_group_cbb.currentTextChanged.connect(lambda text: setattr(self, "department_maintenance_form", text))

    @QtCore.pyqtSlot()
    def load_form_data(self):
        dep = self.ui.list_form_group_cbb.currentText()
        self.ui.apply_machine_table.clearContents()
        self.ui.apply_machine_table.setRowCount(0)
        try:
            result = self.parent.database_process.query(sql = '''SELECT mf.form_name,COUNT(mfr.machine_id)
                                                                FROM `maintenance_form` as mf
                                                                JOIN `maintenance_form_register` as mfr
                                                                ON mf.form_id = mfr.form_id
                                                                JOIN `departments` as d
                                                                ON mf.department_id = d.department_id
                                                                WHERE d.department_name = :dep
                                                                GROUP BY mf.form_name
                                                        ''',params = {'dep':dep})
            total_machine = 0
            if result:
                self.ui.apply_machine_table.setRowCount(len(result))
                self.ui.form_quantity_lbl.setText(f"{len(result)}")
                for r in range(len(result)):
                    self.ui.apply_machine_table.setItem(r,0,QtWidgets.QTableWidgetItem(f"{result[r][0]}"))
                    self.ui.apply_machine_table.setItem(r,1,QtWidgets.QTableWidgetItem(f"{result[r][1]}"))
                    total_machine += result[r][1]
                self.ui.form_list_machine_qty_lbl.setText(f"{total_machine}")
            else:
                self.ui.form_list_machine_qty_lbl.setText("0")
                self.ui.form_quantity_lbl.setText("0")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to load data: {e}")

    def closeEvent(self, event):
        self.ui.apply_machine_table.clearContents()
        super().close()   
        self.deleteLater()                    

class Login_Dialog(QtWidgets.QDialog):
    def __init__(self ,parent = None):
        super().__init__(parent)
        self.ui = Ui_Login()
        self.ui.setupUi(self)
        self.setup_signals()
        self.authenticated = False
        QtCore.QTimer.singleShot(100, self.init_database)

    def init_database(self):
        try:
            self.database = Database_process()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self,
                "Connection failed",
                f"Cannot connect to database:\n{e}"
            )
            self.close()

    def setup_signals(self):
        self.ui.login_btn.clicked.connect(self.login_process)
    
    @QtCore.pyqtSlot()
    def login_process(self):
        try:
            self.ui.pass_status.clear()
            self.ui.user_status.clear()
            username = self.ui.user_line.text().strip()
            password = self.ui.password_line.text().strip()
            # password = "1"
            username = "misa"
            self.ui.login_btn.setEnabled(False)
            QtWidgets.QApplication.processEvents()
            result = self.database.query(sql = ''' SELECT s.user_id,s.username,s.password_hash,r.role_level,d.department_name,s.first_name,s.last_name FROM `Users` as s 
                                                    LEFT JOIN `Departments` as d
                                                    ON s.department_id = d.department_id
                                                    JOIN `Roles` as r
                                                    ON s.role_id = r.role_id
                                                    WHERE username = :username ''', params= {'username':username})
            if not result:
                self.ui.user_status.setText("❌ Wrong username")
                self.ui.user_status.setStyleSheet("color: red;")
                return
            stored_hash = result[0][2]
            if bcrypt.checkpw(password.encode('utf-8'), stored_hash.encode('utf-8')):
                self.login_info = {
                    'user_id': result[0][0],
                    'user_name': result[0][1],
                    'role_level': result[0][3],
                    'department': result[0][4],
                    'first_name': result[0][5],
                    'last_name': result[0][6]
                }
                self.authenticated = True
                try:
                    self.database.query(
                        sql="SET @app_user = :user",
                        params={"user": self.login_info['user_name']}
                    )
                except Exception as e:
                    QtWidgets.QMessageBox.warning(self, "Warning", f"Failed to set MySQL session user: {e}")
                    return
                self.accept()
            else:
                self.ui.pass_status.setText("❌ Wrong password")
                self.ui.pass_status.setStyleSheet("color: red;")

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to login: {e}")
        finally:
            self.ui.login_btn.setEnabled(True)

class NotificationItem(QtWidgets.QWidget):
    def __init__(self, notification_content, parent=None,isYours =False):
        super().__init__()
        self.notification_content = notification_content
        self.title = self.notification_content[4]
        self.message = self.notification_content[5] 
        self.status = self.notification_content[7]
        self.comment = self.notification_content[15]
        self.parent = parent
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        lbl_title = QtWidgets.QLabel(f"<h3 style='color:#ff6600;'>{self.title}</h3>")
        lbl_title.setWordWrap(True)

        lbl_message = QtWidgets.QLabel(self.message)
        lbl_message.setWordWrap(True)
        lbl_message.setStyleSheet("color: gray; font-size: 12px;")

        receive_at =  self.notification_content[10].strftime("%Y-%m-%d %H:%M:%S")
        lbl_time = QtWidgets.QLabel(f"<i style='color: gray; font-size: 10px;'>Received at: {receive_at}</i>")
        lbl_time.setWordWrap(True)
        lbl_status = QtWidgets.QLabel(f"<b style='color: gray; font-size: 12px;'>Status: {self.status}</b>")
        frame = QtWidgets.QFrame()
        frame.setObjectName("frame")
        frame.setMinimumWidth(200)
        frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        frame.setFrameShadow(QtWidgets.QFrame.Raised)

        horizontalLayout = QtWidgets.QHBoxLayout(frame)
        horizontalLayout.setSpacing(10)
        horizontalLayout.setContentsMargins(0, 0, 0, 0)

        btn = QtWidgets.QPushButton("Details")
        btn.setStyleSheet("padding: 4px 10px; background-color: #ddd; border-radius: 5px;")
        btn.clicked.connect(lambda: self.show_details( self.notification_content, isYours))
        btn2 = QtWidgets.QPushButton("Cancel")
        btn2.setStyleSheet("padding: 4px 10px; background-color: #ddd; border-radius: 5px;")
        btn2.clicked.connect(lambda: self.cancel_request(self.notification_content, isYours) if isYours == True else self.reject_action())
        if ( isYours ) and (self.status == "ACCEPTED" or self.status == "REJECTED"):
            btn2.setText("Close")
        horizontalLayout.addWidget(btn)
        horizontalLayout.addWidget(btn2)

        layout.addWidget(lbl_title)
        layout.addWidget(lbl_message)
        layout.addWidget(lbl_time)
        layout.addWidget(lbl_status)
        if self.comment:
            lbl_comment = QtWidgets.QLabel(f"<b style='color: gray; font-size: 12px;'>Reason: {self.comment}</b>")
            layout.addWidget(lbl_comment)
        layout.addWidget(frame, alignment=QtCore.Qt.AlignRight)
    
    @QtCore.pyqtSlot()
    def show_details(self, data,isYours):
        if not isYours:
            self.parent.database_process.query(sql = ''' UPDATE `Notifications`
                                                SET status = 'READ'
                                                WHERE notification_id = :nid ''', params = {'nid': data[0]})
        if data[1] == "update_machine":
            html_table = self.json_to_html_table(json.loads(data[6]))
        detail_html = f"""
        <div style='font-family:Segoe UI; font-size:13px;'>
            <h2 style='color:#007acc; margin-bottom:8px;'>{self.title}</h2>
            <p><b>Message:</b> {self.message}</p>
            <p><b>Status:</b> {self.status}</p>
            <P><b>Reason:</b> {self.comment}</p>
            <p><b>Received at:</b> {data[10].strftime("%Y-%m-%d %H:%M:%S")}</p>
            <hr>
            <h3>Content</h3>
            {html_table}
        </div>
        """
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Notification Details")
        dlg.resize(600, 800)

        layout = QtWidgets.QVBoxLayout(dlg)
        view = QtWidgets.QTextBrowser()
        view.setOpenExternalLinks(True)
        view.setHtml(detail_html)

        layout.addWidget(view)
        frame1 = QtWidgets.QFrame()
        frame1.setObjectName("frame_1")
        frame1.setMinimumWidth(200)
        frame1.setFrameShape(QtWidgets.QFrame.StyledPanel)
        frame1.setFrameShadow(QtWidgets.QFrame.Raised)

        horizontalLayout_1 = QtWidgets.QHBoxLayout(frame1)
        horizontalLayout_1.setSpacing(10)
        horizontalLayout_1.setContentsMargins(0, 0, 0, 0)
        btn_close = QtWidgets.QPushButton("Close")
        btn_close.clicked.connect(dlg.close)
        btn_close.setStyleSheet("padding: 6px 12px; background-color: #007acc; color: white; border-radius: 5px;")
        if not isYours:
            btn_accept = QtWidgets.QPushButton("Accept")
            btn_accept.clicked.connect(self.accept_action)
            btn_accept.setStyleSheet("padding: 6px 12px; background-color: #28a745; color: white; border-radius: 5px;")
            btn_reject = QtWidgets.QPushButton("Reject")
            btn_reject.setStyleSheet("padding: 6px 12px; background-color: #dc3545; color: white; border-radius: 5px;")
            btn_reject.clicked.connect(self.reject_action)
            horizontalLayout_1.addWidget(btn_accept)
            horizontalLayout_1.addWidget(btn_reject)
        else:
            if self.status == "ACCEPTED" or self.status == "REJECTED":
                btn_close.setText("OK")
        horizontalLayout_1.addWidget(btn_close)
        layout.addWidget(frame1, alignment=QtCore.Qt.AlignRight)
        dlg.exec_()
    
    @QtCore.pyqtSlot()
    def cancel_request(self, data,isYours):
        if not isYours:
            return
        if self.status not in ["ACCEPTED","REJECTED"]:
            confirm = QtWidgets.QMessageBox.question(self, "Confirm", "Are you sure you want to cancel this request?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if confirm == QtWidgets.QMessageBox.Yes:
                try:
                    self.parent.database_process.query(sql = '''DELETE FROM `Notifications`
                                                                WHERE notification_id = :nid ''', params = {'nid': data[0]})
                    self.setParent(None)
                    self.deleteLater()
                except Exception as e:
                    QtWidgets.QMessageBox.critical(self,"Error", f"Action Error: {e}")
        else:
             self.parent.database_process.query(sql = '''   UPDATE `Notifications` 
                                                            SET lifecycle_status = "CLOSED"
                                                            WHERE notification_id = :nid ''', params = {'nid': data[0]})
        self.parent.Home_page()
        self.close()
    
    @QtCore.pyqtSlot()
    def accept_action(self):
        try:
            isNotification = self.parent.database_process.query(sql = ''' SELECT * FROM `Notifications` WHERE notification_id = :nid AND lifecycle_status = "PENDING"''', params = {'nid': self.notification_content[0]})
            if isNotification:
                if self.notification_content[1] == "update_machine":
                    content = json.loads(self.notification_content[6])
                    self.parent.database_process.query(sql = '''UPDATE `Machines` AS m
                                                                SET 
                                                                    m.machine_code = :code,
                                                                    m.machine_name = :name,
                                                                    m.line_id = (
                                                                        SELECT p2.line_id FROM `Production_Lines` AS p2 WHERE p2.line_name = :line
                                                                    ),
                                                                    m.maintenance_frequency = :freq,
                                                                    m.maker = :maker,
                                                                    m.model = :model,
                                                                    m.function = :function,
                                                                    m.date_receipt = :receipt,
                                                                    m.machine_status = :status,
                                                                    m.image_link = :image
                                                                WHERE m.machine_code = :code;''', 
                                                                params = {  'code': content.get('old_code'),
                                                                            'name': content.get('name'),
                                                                            'line': content.get('line'),
                                                                            'freq': content.get('freq'),
                                                                            'maker': content.get('maker'),
                                                                            'model': content.get('model'),
                                                                            'function': content.get('function'),
                                                                            'receipt': content.get('receipt'),
                                                                            'status': content.get('status'),
                                                                            'image': content.get('image')})
                    maintenance = content.get('maintenance')
                    params_list = []
                    for row in maintenance:
                        month, week, line = row
                        original_week = week
                        if month in ['1', '2', '3']:
                            quarter = 1
                        elif month in ['4', '5', '6']:
                            quarter = 2
                        elif month in ['7', '8', '9']:
                            quarter = 3
                        elif month in ['10', '11', '12']:
                            quarter = 4
                        else:
                            QtWidgets.QMessageBox.critical(self,"Error", f"Invalid month: {month}")
                            return
                        params_list.append({
                            'code': content.get('old_code'),
                            'line': content.get('line'),
                            'quarter': quarter,
                            'week': week,
                            'original_week': original_week,
                            'year':self.parent.year_num
                        })
                    if params_list:
                        self.parent.database_process.executemany(sql = ''' INSERT INTO `Maintenance_plan` 
                                                                (machine_id, line_id, month_year_id, quarter, week, original_week)
                                                            SELECT 
                                                                m.machine_id,
                                                                (SELECT p.line_id FROM `Production_Lines` AS p WHERE p.line_name = :line LIMIT 1),
                                                                (SELECT my.month_year_id 
                                                                FROM `Months_Years` AS my 
                                                                WHERE my.month = get_working_week_month(:year, :week)
                                                                AND my.year = :year
                                                                LIMIT 1),
                                                                :quarter,
                                                                :week,
                                                                :original_week
                                                            FROM `Machines` AS m
                                                            WHERE m.machine_code = :code
                                                            ON DUPLICATE KEY UPDATE
                                                                line_id = VALUES(line_id),
                                                                month_year_id = VALUES(month_year_id),
                                                                quarter = VALUES(quarter),
                                                                week = VALUES(week),
                                                                original_week = VALUES(original_week);''', params_list = params_list)
                self.parent.database_process.query(sql = '''UPDATE `Notifications`
                                                            SET status = 'ACCEPTED'
                                                            WHERE notification_id = :nid ''', params = {'nid': self.notification_content[0]})
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Action Error: {e}") 
        self.parent.Home_page()
        self.close()
    
    @QtCore.pyqtSlot()
    def reject_action(self):
        confirm = QtWidgets.QMessageBox.question(
            self,
            "Confirm",
            "Are you sure you want to cancel this request?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        if confirm == QtWidgets.QMessageBox.Yes:
            reason, ok = QtWidgets.QInputDialog.getText(
                self,
                "Reject Reason",
                "Please enter reason for rejection:",
                QtWidgets.QLineEdit.Normal
            )
            if not ok:
                return

            reason = reason.strip()
            if not reason:
                reason = None

            try:
                self.parent.database_process.query(
                    sql = '''
                        UPDATE `Notifications`
                        SET status = 'REJECTED', comment = :reason
                        WHERE notification_id = :nid
                    ''',
                    params = {
                        'nid': self.notification_content[0],
                        'reason': reason
                    }
                )
                QtWidgets.QMessageBox.information(self, "Success", "Request rejected successfully.")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Action Error: {e}")
                return
            self.parent.Home_page()
            self.close()

            
    def json_to_html_table(self,d: dict) -> str:
        html = [
            "<table style='border-collapse:collapse; font-size:13px; width:100%;'>"
            "<tr style='background:#007acc; color:white; text-align:left;'>"
            "<th style='padding:6px; border:1px solid #ccc;'>Field</th>"
            "<th style='padding:6px; border:1px solid #ccc;'>Value</th></tr>"
        ]
        for key, value in d.items():
            if value is None:
                value = "<i style='color:gray;'>NULL</i>"
            elif key.lower() == "maintenance" and isinstance(value, list):
                sub_table = [
                    "<table style='border-collapse:collapse; width:100%; margin-top:4px;'>"
                    "<tr style='background:#ddd; font-weight:bold; text-align:center;'>"
                    "<th style=' padding:4px; border:1px solid #bbb;'>Month</th>"
                    "<th style=' padding:4px; border:1px solid #bbb;'>Week</th>"
                    "<th style=' padding:4px; border:1px solid #bbb;'>Line</th></tr>"
                ]
                for row in value:
                    if len(row) == 3:
                        sub_table.append(
                            f"<tr style='text-align:center;'><td style='padding:4px; border:1px solid #ccc;'>{row[0]}</td>"
                            f"<td style='padding:4px; border:1px solid #ccc;'>{row[1]}</td>"
                            f"<td style='padding:4px; border:1px solid #ccc;'>{row[2]}</td></tr>"
                        )
                    else:
                        sub_table.append(
                            "<tr><td colspan='3' style='padding:4px; border:1px solid #ccc; color:gray;'>Invalid format</td></tr>"
                        )
                sub_table.append("</table>")
                value = "".join(sub_table)

            elif isinstance(value, list):
                value = "<br>".join([str(v) for v in value])
            elif isinstance(value, str) and value.startswith("\\\\"):
                value = f"<a href='{value}' style='color:#007acc; text-decoration:none;'>{value}</a>"
            html.append(
                f"<tr><td style='padding:6px; border:1px solid #ccc; font-weight:bold;'>{key.upper()}</td>"
                f"<td style='padding:6px; border:1px solid #ccc;'>{value}</td></tr>"
            )
        html.append("</table>")
        return "".join(html)

class Sync_Missing_Data(QtWidgets.QWidget):
    synced = QtCore.pyqtSignal()
    def __init__(self, parent = None,line_name = "None", data_list = []):
        super().__init__(parent)
        self.parent = parent
        self.ui = Ui_Sync_Missing_Data()
        self.data_list = data_list
        self.ui.setupUi(self)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Window)
        self.setup_signals()
        self.setWindowTitle("Sync Missing Data")
        self.ui.label_2.setText(f"Line: {line_name}")
        headers = ["Machine Code", "Page Number"]
        self.data_model = QtGui.QStandardItemModel()
        self.data_model.setHorizontalHeaderLabels(headers)
        for row in self.data_list:
            items = []
            for col in row:
                item = QtGui.QStandardItem(str(col) if col is not None else "")
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                items.append(item)
            self.data_model.appendRow(items)
        self.ui.data_table.setModel(self.data_model)

    def setup_signals(self):
        self.ui.Cancel_btn.clicked.connect(self.close)
        self.ui.Confirm_btn.clicked.connect(self.sync_data)

    @QtCore.pyqtSlot()
    def close(self):
        return super().close()

    @QtCore.pyqtSlot()
    def sync_data(self):
        try:
            for row in range(self.data_model.rowCount()):
                machine_code = self.data_model.item(row, 0).text()
                page_num = int(self.data_model.item(row, 1).text())
                self.parent.sync_missing_list[machine_code]["page_num"] = page_num
            
            self.synced.emit()
            QtWidgets.QMessageBox.information(self, "Success", "Missing data synchronized successfully.")
            self.close()

        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to sync data: {e}")

class Group_Area_Choose(QtWidgets.QDialog):
    def __init__(self, parent = None, database = None, file_path = None):
        super().__init__(parent)
        self.ui = Ui_Group_choose()
        self.ui.setupUi(self)
        self.setWindowTitle("Group and Area")
        self.database = database
        self.selected_group = None
        self.selected_area = None
        self.file_path = file_path
        self.setup_signals()
        self.load_groups()

    def setup_signals(self):
        self.ui.confirm_btn.clicked.connect(self.confirm_selection)
        self.ui.cancel_btn.clicked.connect(self.reject)
        self.ui.DT_group_input_data.currentIndexChanged.connect(self.load_area)

    @QtCore.pyqtSlot()
    def load_area(self):
        group = self.ui.DT_group_input_data.currentText()
        if not group:
            return
        try:
            result = self.database.query(sql = ''' SELECT downtime_area_name 
                                                    FROM `downtime_areas` as da
                                                    JOIN `departments` as d
                                                    ON da.department_id = d.department_id
                                                    WHERE d.department_name = :group ''', params = {'group': group})
            self.ui.DT_area_input_data.clear()
            if result:
                self.ui.DT_area_input_data.addItems([row[0] for row in result])
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load areas: {e}")
    
    def load_groups(self):
        try:
            result = self.database.query(sql = ''' SELECT d.department_name 
                                                    FROM `downtime_areas` as da
                                                    JOIN `departments` as d
                                                    ON da.department_id = d.department_id
                                                    ORDER BY d.department_name ASC ''')
            self.ui.DT_group_input_data.clear()
            if result:
                self.ui.DT_group_input_data.addItems([row[0] for row in result])
                excel_sheet = pd.ExcelFile(self.file_path).sheet_names if self.file_path else None
                sheet_name_list = [sheet for sheet in excel_sheet]
                self.ui.DT_sheet_name.addItems(sheet_name_list)
                prev = dt.datetime.now() - relativedelta(months=1)
                month_label = prev.strftime("%b").lower()
                for sheet_name in sheet_name_list:
                    if month_label in sheet_name.lower():
                        self.ui.DT_sheet_name.setCurrentText(sheet_name)
                        break
                return
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load groups: {e}")
            return

    @QtCore.pyqtSlot()
    def confirm_selection(self):
        group = self.ui.DT_group_input_data.currentText()
        area = self.ui.DT_area_input_data.currentText()
        sheet_name = self.ui.DT_sheet_name.currentText()
        if not group or not area:
            QtWidgets.QMessageBox.warning(self, "Warning", "Please select both Group and Area.")
            return
        self.selected_group = group
        self.selected_area = area
        self.excel_sheet_name = sheet_name
        self.accept()

class Downtime_Input(QtWidgets.QDialog):
    def __init__(self, parent = None, database = None, data_frame = None, error_frame = None, area_name = None, month_year = None):
        super().__init__(parent)
        self.parent = parent
        self.ui = Ui_DowntimeInputWindow()
        self.ui.setupUi(self)
        self.setWindowTitle("Downtime Data Input")
        self.database = database
        self.data_frame = data_frame
        self.error_frame = error_frame
        self.area_name = area_name
        self.month_year = month_year
        self.setup_signals()
        self.load_data()
    
    def setup_signals(self):
        self.ui.Confirm_btn.clicked.connect(self.confirm_data)
        self.ui.Cancel_btn.clicked.connect(self.reject)

    def load_data(self):
        self.ui.data_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.data_table.setSortingEnabled(False)
        self.ui.data_table.setUpdatesEnabled(False)
        self.ui.error_row_table.setSortingEnabled(False)
        self.ui.error_row_table.setUpdatesEnabled(False)
        headers =["Date", "Line", "Start\nTime","Technical\nStart","Finish\nTime","Total Loss\nTime","Wait\nTechnical","Technical\nName", "Failure\nCode", "Machine Code"]
        self.data_model = QtGui.QStandardItemModel()
        self.data_model.setHorizontalHeaderLabels(headers)
        headers =["Date", "Line", "Start\nTime","Technical\nStart","Finish\nTime","Technical\nName", "Failure\nCode", "Machine Code", "Column Error", "Error Message","Action"]
        self.error_model = QtGui.QStandardItemModel()
        self.error_model.setHorizontalHeaderLabels(headers)
        try:
            self.ui.total_row_lbl.setText(str(self.data_frame.shape[0]) if self.data_frame is not None else "0")
            self.ui.total_downtime_lbl.setText(str(self.data_frame["total_loss_time"].sum()) if self.data_frame is not None else "0")
            if self.data_frame.empty:
                return 
            else:
                for row in range(self.data_frame.shape[0]):
                    items = []
                    for col in range(self.data_frame.shape[1]):
                        value = self.data_frame.iat[row, col]
                        if col == 0:
                            value = int(value) if not pd.isna(value) else ""
                        item = QtGui.QStandardItem(str(value) if value is not None and str(value) != "NaT" else "")
                        item.setTextAlignment(QtCore.Qt.AlignCenter)
                        item.setEditable(False)
                        items.append(item)
                    self.data_model.appendRow(items)
                self.ui.data_table.setModel(self.data_model)
            if self.error_frame.empty:
                return
            else:
                for row in range(self.error_frame.shape[0]):
                    items = []
                    for col in range(self.error_frame.shape[1]):
                        value = self.error_frame.iat[row, col]
                        if col == 0:
                            value = int(value) if not pd.isna(value) else ""
                        item = QtGui.QStandardItem(str(value) if value is not None and str(value) != "NaT" else "")
                        item.setTextAlignment(QtCore.Qt.AlignCenter)
                        item.setEditable(False)
                        items.append(item)
                    self.error_model.appendRow(items)
                self.ui.error_row_table.setModel(self.error_model)
            delegate_btn = ButtonDelegate(buttons=("Edit","Delete"))
            self.ui.error_row_table.setItemDelegateForColumn(10, delegate_btn)
            self.parent.safe_connect(delegate_btn.ButtonClicked, lambda name, idx : self.on_delegate_btn_clicked(name, idx))
            self.ui.error_row_table.setMouseTracking(True)
            self.ui.error_row_table.viewport().setMouseTracking(True)
            self.ui.data_table.setSortingEnabled(True)
            self.ui.data_table.resizeColumnsToContents()
            self.ui.data_table.setUpdatesEnabled(True)
            self.ui.error_row_table.setSortingEnabled(True)
            self.ui.error_row_table.resizeColumnsToContents()
            self.ui.error_row_table.setColumnWidth(10,100)
            self.ui.error_row_table.setUpdatesEnabled(True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to load data into tables: {e}")

    def confirm_data(self):
        self.accept()

    def on_delegate_btn_clicked(self, name, idx):
        row = idx.row()
        if name == "Edit":
            self.new_data = pd.DataFrame(columns = self.data_frame.columns)
            try:
                row_data = {}
                for col in range(self.error_model.columnCount() - 1): 
                    item = self.error_model.item(row, col)
                    if item:
                        row_data[self.error_model.headerData(col, QtCore.Qt.Horizontal)] = item.text()
                
                edit_dialog = QtWidgets.QDialog(self)
                edit_dialog.setWindowTitle("Edit Row Data")
                edit_dialog.setMinimumWidth(500)
                
                layout = QtWidgets.QVBoxLayout(edit_dialog)
                
                form_fields = {}
                for col in range(self.error_model.columnCount() - 1):
                    header = self.error_model.headerData(col, QtCore.Qt.Horizontal)
                    
                    label = QtWidgets.QLabel(f"{header}:")
                    line_edit = QtWidgets.QLineEdit()
                    line_edit.setText(self.error_model.item(row, col).text())
                    
                    form_fields[header] = line_edit
                    
                    h_layout = QtWidgets.QHBoxLayout()
                    h_layout.addWidget(label, 1)
                    h_layout.addWidget(line_edit, 2)
                    layout.addLayout(h_layout)
                
                button_layout = QtWidgets.QHBoxLayout()
                confirm_btn = QtWidgets.QPushButton("Confirm")
                cancel_btn = QtWidgets.QPushButton("Cancel")
                
                button_layout.addWidget(confirm_btn)
                button_layout.addWidget(cancel_btn)
                layout.addLayout(button_layout)
                
                def save_changes():
                    new_row = []
                    pd.concat([self.new_data, pd.DataFrame([row_data])], ignore_index=True, sort=False)
                    for col, (header, line_edit) in enumerate(form_fields.items()):
                        if col < 8:
                            new_row.append(line_edit.text())
                    try:
                        start_time = pd.to_datetime(new_row[2], format='%H:%M')
                        technical_start = pd.to_datetime(new_row[3], format='%H:%M')
                        finish_time = pd.to_datetime(new_row[4], format='%H:%M')
                        
                        total_loss_time = int((finish_time - start_time).total_seconds() / 60)
                        wait_technical_time = int((technical_start - start_time).total_seconds() / 60)
                    except Exception:
                        QtWidgets.QMessageBox.warning(self, "Warning", "Invalid format. Please check again.")
                        return
                    new_row.insert(5, total_loss_time)
                    new_row.insert(6, wait_technical_time)
                    self.new_data.loc[len(self.new_data)] = new_row
                    edit_dialog.accept()
                confirm_btn.clicked.connect(save_changes)
                cancel_btn.clicked.connect(edit_dialog.reject)
                edit_dialog.exec_()
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to edit row: {e}")
            if edit_dialog.result() == QtWidgets.QDialog.Accepted:
                if not self.new_data.empty:
                    self.data_frame = pd.concat([self.data_frame, self.new_data], ignore_index=True, sort=False)
                    self.error_frame.drop(self.error_frame.index[row], inplace=True)
                    for r in range(self.new_data.shape[0]):
                        items = []
                        for c in range(self.new_data.shape[1]):
                            value = self.new_data.iat[r, c]
                            item = QtGui.QStandardItem(str(value) if value is not None and str(value) != "NaT" else "")
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                            item.setEditable(False)
                            items.append(item)
                        self.data_model.appendRow(items)
                    self.error_model.removeRow(row)
                    self.ui.total_row_lbl.setText(str(self.data_frame.shape[0]))
                    self.ui.total_downtime_lbl.setText(str(self.data_frame["total_loss_time"].sum()))
        elif name == "Delete":
            self.ui.error_row_table.model().removeRow(row)

class Error_code_management(QtWidgets.QDialog):
    def __init__(self, parent = None, database = None):
        super().__init__(parent)
        self.parent = parent
        self.database = database
        self.ui = Ui_Error_Code_Management()
        self.ui.setupUi(self)
        self.new_errors_code = []
        self.remove_error_code = []
        self.setWindowTitle("Error Code Management")
        self.load_error_codes()
        self.setup_signals()

    def setup_signals(self):
        self.ui.Group_cbb.currentIndexChanged.connect(self.load_area)
        self.ui.Area_cbb.currentIndexChanged.connect(self.load_process)
        self.ui.Load_btn.clicked.connect(self.filter_error_codes)
        self.ui.Cancel_btn.clicked.connect(self.reject)
        self.ui.Confirm_btn.clicked.connect(self.update_changes)
        self.ui.Insert_btn.clicked.connect(self.insert_row)
        self.ui.Delete_btn.clicked.connect(self.delete_row)

    @QtCore.pyqtSlot()
    def load_error_codes(self):
        try:
            groups = ["All"] + [row[0] for row in self.parent.group]
            self.ui.Group_cbb.addItems(groups)
            self.ui.Group_cbb.setCurrentText("All")
            self.load_area()
            self.ui.error_list_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
            self.ui.error_list_table.setSortingEnabled(False)
            self.ui.error_list_table.setUpdatesEnabled(False)
            result = self.database.query(sql = ''' SELECT ecl.error_code, ecl.error_description, ecl.reason, ecl.recommended_action, ecl.process, da.downtime_area_name
                                                    FROM `error_codes_list` ecl
                                                    JOIN `downtime_areas` da ON ecl.downtime_area_id = da.downtime_area_id
                                                    ORDER BY ecl.error_code ASC, ecl.process COLLATE utf8mb4_unicode_ci ASC;''')
            headers = ["Error Code", "Description", "Reason", "Recommended\nAction", "Process", "Downtime\nArea"]
            self.error_model = QtGui.QStandardItemModel()
            self.error_model.setHorizontalHeaderLabels(headers)
            self.add_item_to_error_list(result, self.error_model)
            self.ui.error_list_table.setModel(self.error_model)
            self.ui.error_list_table.resizeColumnsToContents()
            self.ui.error_list_table.setColumnWidth(3,self.ui.error_list_table.columnWidth(3)-15)
            self.ui.error_list_table.setColumnWidth(2,self.ui.error_list_table.columnWidth(2)-15)
            self.ui.error_list_table.resizeRowsToContents()
            self.ui.error_list_table.setSortingEnabled(True)
            self.ui.error_list_table.setUpdatesEnabled(True)
            self.ui.error_list_table.setAlternatingRowColors(True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to load error codes: {e}")

    def add_item_to_error_list(self,data,model):
        model.removeRows(0, model.rowCount())
        self.new_errors_code = []
        self.remove_error_code = []
        for row_idx, row in enumerate(data):
            for col_idx, col in enumerate(row):
                item = QtGui.QStandardItem(str(col) if col is not None else "")
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                item.setEditable(False)
                model.setItem(row_idx, col_idx, item)

    @QtCore.pyqtSlot()
    def filter_error_codes(self):
        try:
            selected_group = self.ui.Group_cbb.currentText()
            area = self.ui.Area_cbb.currentText()
            process = self.ui.Process_cbb.currentText()
            search_char = self.ui.Search_lnedit.text().strip()
            filter_scripts = ""
            if selected_group == "All" and area == "All" and process == "All" and not search_char:
                filter_scripts = ""
                params =  None
            else:
                conditions = []
                params = {}
                if selected_group != "All":
                    conditions.append("d.department_name = :group")
                    params['group'] = selected_group
                if area != "All":
                    conditions.append("da.downtime_area_name = :area")
                    params['area'] = area
                if process != "All":
                    conditions.append("ecl.process = :process")
                    params['process'] = process
                if search_char:
                    conditions.append("(ecl.error_code LIKE :search OR ecl.error_description LIKE :search OR ecl.reason LIKE :search OR ecl.recommended_action LIKE :search)")
                    params['search'] = f"%{search_char}%"
                filter_scripts = "WHERE " + " AND ".join(conditions)
            result = self.database.query(sql = f''' SELECT ecl.error_code, ecl.error_description, ecl.reason, ecl.recommended_action, ecl.process, da.downtime_area_name
                                                FROM `error_codes_list` ecl
                                                JOIN `downtime_areas` da ON ecl.downtime_area_id = da.downtime_area_id
                                                JOIN `departments` d ON da.department_id = d.department_id
                                                {filter_scripts}
                                                ORDER BY ecl.error_code ASC, ecl.process COLLATE utf8mb4_unicode_ci ASC;''', params = params)
            self.ui.error_list_table.setSortingEnabled(False)
            self.ui.error_list_table.setUpdatesEnabled(False)
            self.add_item_to_error_list(result, self.error_model)
            self.ui.error_list_table.setSortingEnabled(True)
            self.ui.error_list_table.setUpdatesEnabled(True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to filter error codes: {e}")

    @QtCore.pyqtSlot()
    def load_area(self):
        try:
            selected_group = self.ui.Group_cbb.currentText()
            if selected_group == "All":
                result = self.database.query(sql = ''' SELECT DISTINCT da.downtime_area_name
                                                        FROM `downtime_areas` da
                                                        ORDER BY da.downtime_area_name COLLATE utf8mb4_unicode_ci ASC;''')
            else:
                result = self.database.query(sql = ''' SELECT DISTINCT da.downtime_area_name
                                                        FROM `downtime_areas` da
                                                        JOIN `departments` d ON da.department_id = d.department_id
                                                        WHERE d.department_name = :group
                                                        ORDER BY da.downtime_area_name COLLATE utf8mb4_unicode_ci ASC;''', params = {'group': selected_group})
            self.ui.Area_cbb.clear()
            if result:
                self.ui.Area_cbb.addItems(["All"] + [row[0] for row in result])
                self.ui.Area_cbb.setCurrentText("All")
                self.load_process()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to load areas: {e}")

    @QtCore.pyqtSlot()
    def load_process(self):
        try:
            area = self.ui.Area_cbb.currentText()
            if area == "All":
                result = self.database.query(sql = ''' SELECT DISTINCT process 
                                                        FROM `error_codes_list` ecl
                                                        JOIN `downtime_areas` da ON ecl.downtime_area_id = da.downtime_area_id
                                                        ORDER BY process COLLATE utf8mb4_unicode_ci ASC;''')
            else:
                result = self.database.query(sql = ''' SELECT DISTINCT process 
                                                        FROM `error_codes_list` ecl
                                                        JOIN `downtime_areas` da ON ecl.downtime_area_id = da.downtime_area_id
                                                        WHERE da.downtime_area_name = :area
                                                        ORDER BY process COLLATE utf8mb4_unicode_ci ASC;''', params = {'area': area})
            self.ui.Process_cbb.clear()
            if result:
                self.ui.Process_cbb.addItems(["All"] + [row[0] for row in result])
                self.ui.Process_cbb.setCurrentText("All")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to load processes: {e}")
    
    @QtCore.pyqtSlot()
    def delete_row(self):
        if self.error_model.rowCount() == 0:
            QtWidgets.QMessageBox.warning(self, "Warning", "No rows available to delete.")
            return
        index = self.ui.error_list_table.currentIndex()
        if not index.isValid():
            QtWidgets.QMessageBox.warning(self, "Warning", "Please select a row to delete.")
            return
        row = index.row()
        if row in self.new_errors_code:
            self.new_errors_code.remove(row)
        else:
            self.remove_error_code.append((row,self.error_model.item(row, 0).text()))
        self.error_model.removeRow(row)
        self.new_errors_code = [r-1 if r > row else r for r in self.new_errors_code]
        self.remove_error_code = [(r-1 if r > row else r, code) for r, code in self.remove_error_code]


    @QtCore.pyqtSlot()
    def insert_row(self):
        try:
            self.error_model.insertRow(self.error_model.rowCount())
            for col in range(self.error_model.columnCount()):
                item = QtGui.QStandardItem("")
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                item.setEditable(True)
                self.error_model.setItem(self.error_model.rowCount() - 1, col, item)
            self.ui.error_list_table.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.EditKeyPressed)
            self.new_errors_code.append(self.error_model.rowCount() - 1)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to insert new error code: {e}")
    
    @QtCore.pyqtSlot()
    def update_changes(self):
        if not self.new_errors_code and not self.remove_error_code:
            self.accept()
            return
        try:
            question = ""
            if self.new_errors_code:
                question += f"You are going to add {len(self.new_errors_code)} new error code(s):\n{', '.join([self.error_model.item(row, 0).text() for row in self.new_errors_code])}\n"
            if self.remove_error_code:
                question += f"You are going to remove {len(self.remove_error_code)} error code(s):\n{', '.join([error_code for row, error_code in self.remove_error_code])}\n"
            error_code_list_remove = []
            reply = QtWidgets.QMessageBox.question(self, "Confirm", f'''Are you sure you want to apply changes? \n{question}''', QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.No:
                return
            for row,error_code in self.remove_error_code:
                error_code_list_remove.append({'code': error_code})
            if error_code_list_remove:
                self.database.executemany(sql = ''' DELETE FROM `error_codes_list` WHERE error_code = :code ''', params_list = error_code_list_remove)
            for row in self.new_errors_code:
                error_code = self.error_model.item(row, 0).text()
                description = self.error_model.item(row, 1).text()
                reason = self.error_model.item(row, 2).text()
                recommended_action = self.error_model.item(row, 3).text()
                process = self.error_model.item(row, 4).text()
                area_name = self.error_model.item(row, 5).text()
                if not error_code or not area_name:
                    QtWidgets.QMessageBox.warning(self, "Warning", f"Error Code and Downtime Area cannot be empty. Please check row {row+1}.")
                    return
                self.database.query(sql = ''' INSERT INTO `error_codes_list` (error_code, error_description, reason, recommended_action, process, downtime_area_id)
                                            VALUES (:code, :description, :reason, :action, :process,
                                            (SELECT downtime_area_id FROM `downtime_areas` WHERE downtime_area_name = :area LIMIT 1))
                                            ON DUPLICATE KEY UPDATE
                                            error_description = VALUES(error_description),
                                            reason = VALUES(reason),
                                            recommended_action = VALUES(recommended_action),
                                            process = VALUES(process),
                                            downtime_area_id = VALUES(downtime_area_id);''', 
                                            params = {
                                                'code': error_code,
                                                'description': description,
                                                'reason': reason,
                                                'action': recommended_action,
                                                'process': process,
                                                'area': area_name
                                            })
            QtWidgets.QMessageBox.information(self,"Success", "Changes updated successfully.")
            self.accept()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self,"Error", f"Failed to update changes: {e}")
            return
        
def main():
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass
    except Exception:
        pass

    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    app = QtWidgets.QApplication(sys.argv)
    while True:
        login = Login_Dialog()
        if login.exec() != QtWidgets.QDialog.Accepted or not login.authenticated:
            break  
        window = OEEAppWindow(login.login_info)
        window.show()
        window.ui.Home_btn.setStyleSheet("""
            #Home_btn {
                background-color: rgba(0, 0, 255, 0.07);
                border: none;
                border-top: 1px solid rgba(0, 0, 255, 1);
                border-bottom: 1px solid rgba(0, 0, 255, 1);
            }
        """)

        QtCore.QTimer.singleShot(100, window._init_database)
        QtCore.QTimer.singleShot(110, window.Home_page)
        app.exec_()
        if getattr(window, "logout_triggered", False):
            continue
        break  

    sys.exit(0)

if __name__ == "__main__":
    main()
  
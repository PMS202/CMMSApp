# -*- coding: utf-8 -*-
from matplotlib import colors

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
from UI.OEE_Edit_data import UI_OEE_Edit_Data
from UI.OEE_other_data import UI_OEE_Other_Data
from UI.Downtime_detail_report import UI_DT_detail_report
from UI.OEE_mini_card import Ui_OEE_mini_card
from KPI.KPI_widget import ArcGauge, PetalCard
from Database.MariaDB import Database_process
from Stock_control.stock_delegate import StockItemDelegate, ImageCache
from Stock_control.image_loader import ImageLoaderRunnable
from Maintenance.printer import Printer_process
from Maintenance.scan_qrcode import Scan_record_process
from Maintenance.attached_equipment import DynamicSuggestion
from Downtimes.Excel_processing import Downtime_Excel_Processor
from UI.Report_input import Ui_ReportInput
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
import fitz
import shutil
import datetime as dt
import re
from decimal import Decimal
from dateutil.relativedelta import relativedelta
from pyqtspinner.spinner import Qt, WaitingSpinner
from sqlalchemy import text
from scipy.interpolate import interp1d, PchipInterpolator
from scipy.ndimage import gaussian_filter1d
import calendar
import subprocess
import traceback

STRICT_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class OEEAppWindow(QtWidgets.QMainWindow):
    def __init__(self, login_info=None):
        super().__init__()
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding,
                           QtWidgets.QSizePolicy.Expanding)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowIcon(QtGui.QIcon(resource_path("Icons/Tokin-logo.ico")))
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
        self.qty_week = self.ui.company_week_number(
            dt.date(self.year_num, 12, 31))
        self.spinner = WaitingSpinner(
            self, center_on_parent=True, disable_parent_when_spinning=True, speed=1.1)
        self.spinner.roundness = 70.0
        self.spinner.line_length = 30
        self.spinner.line_width = 10
        self.spinner.inner_radius = 40
        self.spinner.number_of_lines = 100
        self.spinner.color = QtGui.QColor(68, 60, 113)

    def setup_signals(self):
        self.ui.Home_btn.clicked.connect(self.Home_page)
        self.ui.OEE_btn.clicked.connect(self.OEE_page)
        self.ui.Maintenance_btn.clicked.connect(self.Maintenance_page)
        self.ui.KPI_btn.clicked.connect(self.KPI_monitoring_page)
        self.ui.Stock_btn.clicked.connect(self.Stock_control_page)
        self.ui.Downtime_btn.clicked.connect(self.Downtime_page)
        self.ui.Main_Home_btn.clicked.connect(self.Mainten_Home_page)
        self.ui.Main_Input_record_btn.clicked.connect(self.Mainten_Input_page)
        self.ui.Main_Print_record_btn.clicked.connect(self.Mainten_Print_page)
        self.ui.Main_detail_plan_btn.clicked.connect(
            self.Mainten_Detail_plan_page)
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
        self.ui.profile_btn.clicked.connect(lambda _: self.ui.frame_60.show(
        ) if self.ui.frame_60.isHidden() else self.ui.frame_60.hide())
        self.ui.return_home_btn.clicked.connect(lambda _: self.return_home())
        self.ui.user_info_btn.clicked.connect(lambda _: self.user_info())
        self.ui.change_password_btn.clicked.connect(
            lambda _: self.change_password(form_user_info=False))
        self.ui.change_password_inside_btn.clicked.connect(
            lambda _: self.change_password(form_user_info=True))

    def _init_database(self):
        try:
            self.database_process = Database_process()
            self.group = self.database_process.query(
                sql=''' SELECT department_name FROM `Departments` ''')
        except ConnectionError as e:
            QtWidgets.QMessageBox.critical(self, "Error", str(e))
            self.close()

    def safe_connect(self, signal, slot):
        try:
            signal.disconnect()
        except TypeError:
            pass
        signal.connect(slot)

# ==========================Function of Maintenance page ==================================================================================BEGIN
# ==========================Function of Maintenance page ==================================================================================BEGIN
# ==========================Function of Maintenance page ==================================================================================BEGIN

    def expand_windown_animation(self, is_expand=False):
        size_animation = QtCore.QPropertyAnimation(self, b"size")
        size_animation.setDuration(250)
        size_animation2 = QtCore.QPropertyAnimation(
            self.ui.func_frame, b"size")
        size_animation2.setDuration(250)
        size_animation3 = QtCore.QPropertyAnimation(
            self.ui.main_stacked, b"size")
        size_animation3.setDuration(250)
        size_animation4 = QtCore.QPropertyAnimation(
            self.ui.Mainten_widget, b"size")
        size_animation4.setDuration(250)
        size_animation5 = QtCore.QPropertyAnimation(
            self.ui.Mainten_frame, b"size")
        size_animation5.setDuration(250)
        size_animation6 = QtCore.QPropertyAnimation(
            self.ui.Maintenance_stacked, b"size")
        size_animation6.setDuration(250)
        pos_animation = QtCore.QPropertyAnimation(self, b"pos", self)
        pos_animation.setDuration(250)
        pos_animation.setEasingCurve(QtCore.QEasingCurve.OutCubic)
        if is_expand:
            size_animation.setStartValue(QtCore.QSize(932, 545))
            size_animation.setEndValue(QtCore.QSize(1500, 870))
            size_animation2.setStartValue(QtCore.QSize(121, 551))
            size_animation2.setEndValue(QtCore.QSize(121, 850))
            size_animation3.setStartValue(QtCore.QSize(811, 551))
            size_animation3.setEndValue(QtCore.QSize(1379, 850))
            size_animation4.setStartValue(QtCore.QSize(811, 551))
            size_animation4.setEndValue(QtCore.QSize(1379, 850))
            size_animation5.setStartValue(QtCore.QSize(811, 471))
            size_animation5.setEndValue(QtCore.QSize(1379, 770))
            size_animation6.setStartValue(QtCore.QSize(811, 431))
            size_animation6.setEndValue(QtCore.QSize(1379, 730))
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
            size_animation.setStartValue(QtCore.QSize(1500, 870))
            size_animation.setEndValue(QtCore.QSize(932, 545))
            size_animation2.setStartValue(QtCore.QSize(121, 850))
            size_animation2.setEndValue(QtCore.QSize(121, 551))
            size_animation3.setStartValue(QtCore.QSize(1379, 850))
            size_animation3.setEndValue(QtCore.QSize(811, 551))
            size_animation4.setStartValue(QtCore.QSize(1379, 850))
            size_animation4.setEndValue(QtCore.QSize(811, 551))
            size_animation5.setStartValue(QtCore.QSize(1379, 770))
            size_animation5.setEndValue(QtCore.QSize(811, 471))
            size_animation6.setStartValue(QtCore.QSize(1379, 730))
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

    def set_stylesheet_change_page(self, button: tuple):
        button[0].setStyleSheet('''
                                    QPushButton {
                                        background-color: rgba(0, 0, 255, 0.07);
                                        border: none;                     
                                        border-top: 1px solid rgba(0, 0, 255, 1);
                                        border-bottom: 1px solid rgba(0, 0, 255, 1);
                                    }
        ''')
        for i in range(1, len(button)):
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
        self.set_stylesheet_change_page((self.ui.Home_btn, self.ui.OEE_btn, self.ui.Maintenance_btn,
                                        self.ui.KPI_btn,
                                        self.ui.Stock_btn, self.ui.Downtime_btn))
        if self.is_expanded:
            self.is_expanded = False
            self.expand_windown_animation(self.is_expanded)
        if self.login_info is not None:
            self.ui.welcome_label.setText(
                f"Hello {self.login_info['first_name']} {self.login_info['last_name']}")
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
                                                        ''', params={"id": self.login_info['user_id']})
            self.ui.notification_listwidget.clear()
            for note in notifications:
                item_widget = NotificationItem(
                    notification_content=note, parent=self, isYours=False)
                list_item = QtWidgets.QListWidgetItem()
                list_item.setSizeHint(item_widget.sizeHint())
                self.ui.notification_listwidget.addItem(list_item)
                self.ui.notification_listwidget.setItemWidget(
                    list_item, item_widget)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load notifications: {e}")

        try:
            your_requests = self.database_process.query(sql='''
                                                            SELECT * FROM `Notifications`
                                                            WHERE sender_id = :id AND lifecycle_status NOT IN ('CLOSED')
                                                            ORDER BY created_at DESC
                                                        ''', params={"id": self.login_info['user_id']})
            self.ui.your_request_listwidget.clear()
            for note in your_requests:
                item_widget = NotificationItem(
                    notification_content=note, parent=self, isYours=True)
                list_item = QtWidgets.QListWidgetItem()
                list_item.setSizeHint(item_widget.sizeHint())
                self.ui.your_request_listwidget.addItem(list_item)
                self.ui.your_request_listwidget.setItemWidget(
                    list_item, item_widget)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load notifications: {e}")
        self.safe_connect(self.ui.update_first_name_btn.clicked, lambda _: self.update_user_info(
            update_column="first_name", update_content=self.ui.first_name_lnedit.text()))
        self.safe_connect(self.ui.update_last_name_btn.clicked, lambda _: self.update_user_info(
            update_column="last_name", update_content=self.ui.last_name_lnedit.text()))
        self.safe_connect(self.ui.update_password_btn.clicked, lambda _: self.update_user_info(
            update_column="password_hash", update_content=self.ui.confirm_password_lnedit.text()))
        self.safe_connect(self.ui.logout_btn.clicked, self.logout)

    @QtCore.pyqtSlot()
    def return_home(self):
        self.ui.frame_60.hide()
        self.ui.frame_58.setMaximumWidth(16777215)
        self.ui.frame_59.setMaximumWidth(16777215)
        self.ui.frame_61.setMaximumWidth(0)
        self.ui.change_password_frame.setMaximumWidth(0)
        self.ui.horizontalLayout_41.setContentsMargins(0, 0, 0, 0)
        self.ui.horizontalLayout_41.setSpacing(0)

    @QtCore.pyqtSlot()
    def user_info(self):
        self.ui.frame_60.hide()
        self.ui.frame_58.setMaximumWidth(0)
        self.ui.frame_59.setMaximumWidth(0)
        self.ui.frame_61.setMaximumWidth(350)
        self.ui.change_password_frame.setMaximumWidth(0)
        self.ui.horizontalLayout_41.setContentsMargins(0, 0, 400, 0)
        self.ui.horizontalLayout_41.setSpacing(0)

    @QtCore.pyqtSlot()
    def change_password(self, form_user_info=False):
        if not form_user_info:
            self.ui.frame_60.hide()
            self.ui.frame_58.setMaximumWidth(0)
            self.ui.frame_59.setMaximumWidth(0)
            self.ui.frame_61.setMaximumWidth(0)
            self.ui.change_password_frame.setMaximumWidth(16777215)
            self.ui.horizontalLayout_41.setContentsMargins(250, 0, 250, 0)
            self.ui.horizontalLayout_41.setSpacing(0)
        else:
            self.ui.frame_58.setMaximumWidth(0)
            self.ui.frame_59.setMaximumWidth(0)
            self.ui.frame_61.setMaximumWidth(350)
            self.ui.change_password_frame.setMaximumWidth(16777215)
            self.ui.horizontalLayout_41.setContentsMargins(0, 0, 150, 0)
            self.ui.horizontalLayout_41.setSpacing(10)

    @QtCore.pyqtSlot()
    def update_user_info(self, update_column, update_content):
        try:
            if update_column != "password_hash":
                self.database_process.query(f''' UPDATE `Users` 
                                                SET {update_column} = :update_content 
                                                WHERE user_id = :id''', params={'update_content': update_content,
                                                                                'id': self.login_info['user_id']})
            else:
                result = self.database_process.query(sql='''  SELECT password_hash FROM `Users`
                                                                WHERE user_id = :id''', params={'id': self.login_info['user_id']})
                if bcrypt.checkpw(self.ui.current_password_lnedit.text().strip().encode('utf-8'), result[0][0].encode('utf-8')):
                    if self.ui.new_password_lnedit.text() == self.ui.confirm_password_lnedit.text():
                        self.database_process.query(f''' UPDATE `Users` 
                                                SET {update_column} = :update_content 
                                                WHERE user_id = :id''', params={'update_content': bcrypt.hashpw(self.ui.confirm_password_lnedit.text().strip().encode('utf-8'), bcrypt.gensalt()).decode('utf-8'),
                                                                                'id': self.login_info['user_id']})
                    else:
                        QtWidgets.QMessageBox.warning(
                            self, "Wrong Password", "Incorrect confirmation for the New password")
                        return
                else:
                    QtWidgets.QMessageBox.warning(
                        self, "Wrong Password", "Incorrect Current password")
                    return
            QtWidgets.QMessageBox.information(
                self, "Update success", "Information updated successfully")
            self.logout(needconfirm=False)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to update data: {e}")

    @QtCore.pyqtSlot()
    def logout(self, needconfirm=True):
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
# ==========================Function of KPI page ====================================================================================
# ==========================Function of KPI page ====================================================================================
# ==========================Function of KPI page ====================================================================================

    def KPI_monitoring_page(self):
        self.ui.main_stacked.setCurrentWidget(self.ui.KPI_monitoring_page)
        self.set_stylesheet_change_page((self.ui.KPI_btn, self.ui.OEE_btn, self.ui.Home_btn, self.ui.Maintenance_btn,
                                        self.ui.Stock_btn, self.ui.Downtime_btn))
        if not self.is_expanded:
            self.is_expanded = True
            self.expand_windown_animation(self.is_expanded)
        self.spinner.start()
        BG_COLOR = "#eef1f5"
        CARD_BG = "#ffffff"
        TEXT_DARK = "#1e2733"
        TEXT_GRAY = "#6b7684"
        TEAL = "#22b6c9"
        GREEN = "#2fbf5a"
        RED = "#e2483d"
        AMBER = "#e8a33d"
        TRACK = "#e6eaf0"
        BLUE_LOGO = "#1447a3"
        
        def fetch_kpi_data():
            kpi_maintenance_data = self.database_process.query(sql='''SELECT
                                                                    (SELECT "Total" ) as department_name,
                                                                    COUNT(CASE WHEN mp.status IN ('Ontime','Overdue') THEN 1 END) AS total,
                                                                    COUNT(CASE WHEN mp.status = "Overdue" THEN 1 END) AS overdue
                                                                    FROM Maintenance_plan mp
                                                                    JOIN `Months_Years` as my ON my.month_year_id = mp.month_year_id
                                                                    WHERE my.year = :year;  ''', params={"year": self.year_num})
            kpi_oee_data = self.database_process.query(sql='''SELECT AVG(`Availability_percentage`) AS avg_availability, 
                                                                AVG(`Performance_percentage`) AS avg_performance, 
                                                                AVG(`Quality_percentage`) AS avg_quality, 
                                                                AVG(`OEE_percentage`) AS avg_oee
                                                                FROM oee_report
                                                                WHERE YEAR(`production_date`) = :year;''', params={"year": self.year_num})
            kpi_dt_data = self.database_process.query(sql=f'''SELECT SUM(Total_Loss) AS total_loss, SUM(Repair_Time) AS total_repair_time, COUNT(*) AS total_records
                                                                FROM `downtime_report` AS dr
                                                                WHERE YEAR(dr.Date) = :year;''', params={"year": self.year_num})
            working_time = self.database_process.query(sql=f''' SELECT SUM(lot.operation_hours)*60 - COALESCE(SUM(lot.setup_time),0) - COALESCE(SUM(lot.break_time),0) AS total_planed_time 
                                                                FROM `line_operation_times` as lot
                                                                JOIN downtime_areas_production_lines as dapl ON lot.line_id = dapl.line_id
                                                                JOIN downtime_areas as da ON dapl.downtime_area_id = da.downtime_area_id
                                                                WHERE YEAR(lot.operation_date) = :year
                                                                ''', params={ "year": self.year_num})   
            return {"kpi_maintenance_data": kpi_maintenance_data, "kpi_oee_data": kpi_oee_data, "kpi_dt_data": kpi_dt_data, "working_time": working_time}
             
        def h_line(text, color=TEXT_DARK, size=12, bold=False, align=Qt.AlignLeft):
            lbl = QtWidgets.QLabel(text)
            weight = "700" if bold else "400"
            lbl.setStyleSheet(f"color:{color}; font-size:{size}px; font-weight:{weight};")
            lbl.setAlignment(align)
            return lbl

        def build_maintenance_card(percentage=100):
            card = PetalCard(notch="br")
            card.add_header(icon_link=resource_path("Icons/maintenance.ico"), title="MAINTENANCE KPI", align_right=False)
            row = QtWidgets.QHBoxLayout()
            row.setSpacing(18)
            info = QtWidgets.QVBoxLayout()
            info.setSpacing(6)
            info.addStretch(1)
            info.addWidget(h_line("Maintenance Achievement", TEXT_DARK, 13, True))
            if percentage < 100:
                badge = QtWidgets.QLabel("❌  Missed target")
                badge.setStyleSheet(
                    "background-color:#fdecea; color:#e2483d; font-size:12px; font-weight:600;"
                    "border-radius:11px; padding:4px 12px;")
                info.addWidget(h_line(f"KPI: 100% / Actual: {percentage:.0f}%", RED, 14, True))
                card.set_shadow_color(RED)
            else:
                badge = QtWidgets.QLabel("✔  Reached target")
                badge.setStyleSheet(
                    "background-color:#e5f7ea; color:#2fbf5a; font-size:12px; font-weight:600;"
                    "border-radius:11px; padding:4px 12px;")
                info.addWidget(h_line(f"KPI: 100% / Actual: {percentage:.0f}%", GREEN, 14, True))
                card.set_shadow_color(GREEN)
            badge.setFixedWidth(160)
            badge.setAlignment(Qt.AlignCenter)
            info.addWidget(badge)
            row.addLayout(info)
            row.addStretch(1)
            card.vbox.addLayout(row)
            def make_shades(color_hex):
                c = QtGui.QColor(color_hex)
                light = c.lighter(135).name()  
                dark  = c.darker(110).name()  
                return light, dark
            flash_card = QtWidgets.QWidget()
            flash_card.setObjectName("flash_card")
            flash_card.setStyleSheet("""
                #flash_card {
                    background-color: transparent;
                    border-radius: 16px;
                    border: 1px solid #ECECEC;
                }
            """)
            shadow = QtWidgets.QGraphicsDropShadowEffect()
            shadow.setBlurRadius(35)
            shadow.setXOffset(0)
            shadow.setYOffset(6)
            shadow.setColor(QtGui.QColor(0, 0, 0, 40)) 
            flash_card.setGraphicsEffect(shadow)

            flash_card_layout = QtWidgets.QVBoxLayout(flash_card)
            flash_card_layout.setContentsMargins(24, 24, 24, 24)
            flash_card_layout.setSpacing(14)

            if percentage < 100:
                color_hex = "#E2483D"  
            else:
                color_hex = "#2FBF5A"   
            light_hex, dark_hex = make_shades(color_hex)

            val_layout = QtWidgets.QHBoxLayout()
            val_layout.setSpacing(10)

            lbl_value = QtWidgets.QLabel(f"{percentage:.0f}%")
            lbl_value.setStyleSheet(f"""
                color: {color_hex};
                font-weight: 800;
                font-size: 64px;
                font-family: 'Segoe UI', Arial, sans-serif;
                border: none;
                background: transparent;
            """)

            lbl_target = QtWidgets.QLabel("KPI: 100%")
            lbl_target.setStyleSheet("""
                color: #9E9E9E;
                font-size: 20px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                border: none;
                background: transparent;
            """)
            lbl_target.setAlignment(Qt.AlignBottom | Qt.AlignRight)

            val_layout.addWidget(lbl_value)
            val_layout.addStretch()         
            val_layout.addWidget(lbl_target)

            progress_bar = QtWidgets.QProgressBar()
            progress_bar.setValue(int(percentage))
            progress_bar.setTextVisible(False)
            progress_bar.setFixedHeight(40) 
            progress_bar.setStyleSheet(f"""
                QProgressBar {{
                    border: none;
                    border-radius: 11px;
                    background-color: #F0F1F3;
                }}
                QProgressBar::chunk {{
                    border-radius: 11px;
                    background-color: qlineargradient(
                        x1:0, y1:0, x2:1, y2:0,
                        stop:0 {light_hex},
                        stop:1 {dark_hex}
                    );
                }}
            """)

            flash_card_layout.addLayout(val_layout)
            flash_card_layout.addWidget(progress_bar)
            card.vbox.addWidget(flash_card)
            card.vbox.addStretch(1)
            card.vbox.setContentsMargins(10, 10, 100, 10)
            return card
        def build_oee_card(oee_value=0,a_value=0,p_value=0,q_value=0):
            card = PetalCard(notch="bl")
            card.add_header(icon_link=resource_path("Icons/OEE.ico"), title = "OVERALL EQUIPMENT EFFECTIVENESS", align_right=True)
            content_widget = QtWidgets.QWidget()
            content_widget.setStyleSheet("background-color:transparent;")
            content_layout = QtWidgets.QVBoxLayout(content_widget)
            content_layout.setContentsMargins(0, 0, 10, 0)
            content_layout.setSpacing(20)
            gauge = ArcGauge(
                value=oee_value*100, max_value=100, color=GREEN,
                ticks=[0, 25, 50, 75, 100],
                start_angle=180, span_angle=-180, minor_ticks=40,
                center_lines=[
                    {"text": "OEE:", "size": 12, "color": TEXT_GRAY},
                    {"text": f"{(oee_value*100):.2f}%", "size": 18, "color": TEXT_DARK, "bold": True},
                     {"text": "KPI = 70%", "size": 12, "color": "#CC7D05", "bold": True},
                ], ratio = 1.2
            )
            content_layout.addWidget(gauge, stretch=1)
            metrics_widget = QtWidgets.QWidget()
            metrics_layout = QtWidgets.QVBoxLayout(metrics_widget)
            metrics_layout.setSpacing(8)
            metrics_layout.setContentsMargins(100, 0, 0, 0)
            metrics_layout.addStretch()
            metrics_layout.addWidget(MetricMiniBar("A", a_value))
            metrics_layout.addWidget(MetricMiniBar("P", p_value))
            metrics_layout.addWidget(MetricMiniBar("Q", q_value))
            metrics_layout.addStretch()
            content_layout.addWidget(metrics_widget, stretch=1)
            card.vbox.addWidget(content_widget)
            card.vbox.addStretch(1)
            if oee_value < 0.7:
                card.set_shadow_color(RED)
            else:
                card.set_shadow_color(GREEN)
            return card
        def build_mttr_card(mttr_value=0,mttr_time="00:00:00"):
            card = PetalCard(notch="tr")
            card.add_header(icon_link=resource_path("Icons/MTTR.ico"), title = "MEAN TIME TO REPAIR", align_right=False)
            diff = ((mttr_value - 15) / 15) * 100
            if diff > 0:
                color_code = RED
                comment = f"▲ {diff:.1f}% vs target\n(KPI = 15m)"
            else:
                color_code = GREEN
                comment = f"▼ {abs(diff):.1f}% vs target\n(KPI = 15m)"
            gauge = ArcGauge(
                value=mttr_value, max_value=45, color=color_code,
                ticks=[0, 5, 10, 15, 20, 25, 30, 35, 40, 45],
                start_angle=225, span_angle=-270, minor_ticks=54,
                center_lines=[
                    {"text": "MTTR", "size": 14, "color": TEXT_DARK, "bold": True},
                    {"text": mttr_time, "size": 18, "color": TEXT_DARK, "bold": True},
                ],ratio = 1
            )
            card.vbox.addWidget(gauge, 1)
            diff = ((mttr_value - 15) / 15) * 100
            delta = QtWidgets.QLabel(comment)
            delta.setStyleSheet(f"color:{color_code}; font-size:12px; font-weight:600;")
            card.set_shadow_color(color_code)
            delta.setAlignment(Qt.AlignCenter)
            card.vbox.addWidget(delta)
            return card
        def build_mtbf_card(mtbf_value=0,mtbf_time="00:00:00"):
            card = PetalCard(notch="tl")
            card.add_header(icon_link=resource_path("Icons/MTBF.ico"), title = "MEAN TIME BETWEEN FAILURES", align_right=True)
            diff = ((mtbf_value - 70) / 70) * 100
            target_mtbf_time = self.change_time_format(time_value = 70, input_unit = "m", output_unit = True)
            if diff > 0:
                color_code = GREEN
                comment = f"▲ {diff:.1f}% vs target\n(KPI = {target_mtbf_time})"
            else:
                color_code = RED
                comment = f"▼ {abs(diff):.1f}% vs target\n(KPI = {target_mtbf_time})"
            gauge = ArcGauge(
                value=mtbf_value/60, max_value=24, color=color_code,
                ticks=[0, 4, 8, 12, 16, 20, 24],
                start_angle=225, span_angle=-270, minor_ticks=54,
                center_lines=[
                    {"text": "MTBF", "size": 14, "color": TEXT_DARK, "bold": True},
                    {"text": mtbf_time, "size": 18, "color": TEXT_DARK, "bold": True},
                ], ratio = 1
            )
            card.vbox.addWidget(gauge, 1)
            delta = QtWidgets.QLabel(comment)
            delta.setStyleSheet(f"color:{color_code}; font-size:12px; font-weight:600;")
            card.set_shadow_color(color_code)
            delta.setAlignment(Qt.AlignCenter)
            card.vbox.addWidget(delta)
            return card
        class CenterBadge(QtWidgets.QLabel):
            def __init__(self, parent, icon_path, diameter=120):
                super().__init__(parent)
                self.diameter = diameter
                self.setFixedSize(diameter, diameter)
                pix = QtGui.QPixmap(icon_path).scaled(
                    diameter - 50, diameter - 50,
                    Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.setPixmap(pix)
                self.setAlignment(Qt.AlignCenter)

                self.setStyleSheet(f"""
                    background-color: #ffffff;
                    border-radius: {diameter // 2}px;
                    border: 1px solid #e6eaf0;
                """)

                shadow = QtWidgets.QGraphicsDropShadowEffect(self)
                shadow.setBlurRadius(30)
                shadow.setColor(QtGui.QColor(0, 0, 0, 45))
                shadow.setOffset(0, 0)
                self.setGraphicsEffect(shadow)

        def on_kpi_data_fetched(data):
            while self.ui.KPI_monitoring_gridLayout.count():
                item = self.ui.KPI_monitoring_gridLayout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            kpi_maintenance_data = data["kpi_maintenance_data"]
            kpi_oee_data = data["kpi_oee_data"]
            kpi_dt_data = data["kpi_dt_data"]
            working_time = data["working_time"]
            total_maintenance = kpi_maintenance_data[0][1] if kpi_maintenance_data else 0
            overdue_maintenance = kpi_maintenance_data[0][2] if kpi_maintenance_data else 0
            maintenance_percentage = (total_maintenance - overdue_maintenance) / total_maintenance * 100 if total_maintenance > 0 else 0
            a_avg = kpi_oee_data[0][0] if kpi_oee_data else 0
            p_avg = kpi_oee_data[0][1] if kpi_oee_data else 0
            q_avg = kpi_oee_data[0][2] if kpi_oee_data else 0
            oee_avg = kpi_oee_data[0][3] if kpi_oee_data else 0
            total_loss = float(kpi_dt_data[0][0]) if kpi_dt_data else 0
            total_repair_time = float(kpi_dt_data[0][1]) if kpi_dt_data else 0
            count_records = int(kpi_dt_data[0][2]) if kpi_dt_data else 0
            planned_time = float(working_time[0][0]) if working_time else 0
            # mttr_value = total_repair_time / count_records if kpi_dt_data and count_records > 0 else total_repair_time
            mttr_value = 10.7667
            mttr_time = self.change_time_format(time_value = mttr_value, input_unit = "m", output_unit = True)
            # mtbf_value = (planned_time - total_loss) / count_records if kpi_dt_data and count_records > 0 else planned_time - total_loss
            mtbf_value = 149.1333
            mtbf_time = self.change_time_format(time_value = mtbf_value, input_unit = "m", output_unit = True)
            self.ui.KPI_monitoring_gridLayout.addWidget(build_maintenance_card(maintenance_percentage), 0, 0)
            self.ui.KPI_monitoring_gridLayout.addWidget(build_oee_card(oee_value=oee_avg, a_value=a_avg, p_value=p_avg, q_value=q_avg), 0, 1)
            self.ui.KPI_monitoring_gridLayout.addWidget(build_mttr_card(mttr_value=mttr_value, mttr_time=mttr_time), 1, 0)
            self.ui.KPI_monitoring_gridLayout.addWidget(build_mtbf_card(mtbf_value=mtbf_value, mtbf_time=mtbf_time), 1, 1)
            grid = self.ui.KPI_monitoring_gridLayout
            grid.setColumnStretch(0, 1)
            grid.setColumnStretch(1, 1)
            grid.setRowStretch(0, 1)
            grid.setRowStretch(1, 1)
            grid.setHorizontalSpacing(10)  
            grid.setVerticalSpacing(10)
            container = grid.parentWidget()
            if not hasattr(self, "center_badge") or self.center_badge is None:
                self.center_badge = CenterBadge(
                    container, resource_path("Icons/factory.ico"), diameter=190)
                container.installEventFilter(self) 

            self.center_badge.show()
            self.center_badge.raise_()
            QtCore.QTimer.singleShot(0, self._reposition_center_badge)
        try:
            if hasattr(self, 'KPI_dashboard_worker') and self.KPI_dashboard_worker.isRunning():
                self.spinner.stop()
                return
            self.KPI_dashboard_worker = WorkerThread(fetch_kpi_data)
            self.KPI_dashboard_worker.finished.connect(lambda res: on_kpi_data_fetched(res))
            self.KPI_dashboard_worker.start()
            self.spinner.stop()
        except Exception as e:
            self.spinner.stop()
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to refresh KPI Overview page: {e}")
            

    def _reposition_center_badge(self):
        if not getattr(self, "center_badge", None):
            return
        grid = self.ui.KPI_monitoring_gridLayout
        it00 = grid.itemAtPosition(0, 0)
        it11 = grid.itemAtPosition(1, 1)
        if not (it00 and it11 and it00.widget() and it11.widget()):
            return
    
        g00 = it00.widget().geometry()
        g11 = it11.widget().geometry()
    
        cx = (g00.right() + g11.left()) // 2
        cy = (g00.bottom() + g11.top()) // 2
    
        d = self.center_badge.diameter
        self.center_badge.move(cx - d // 2, cy - d // 2)
        self.center_badge.raise_()

    def eventFilter(self, obj, event):
        if (obj is self.ui.KPI_monitoring_gridLayout.parentWidget()
                and event.type() == QtCore.QEvent.Resize):
            self._reposition_center_badge()
        return super().eventFilter(obj, event)
# ==========================END of KPI page ==========================================================================================
# ==========================END of KPI page ==========================================================================================
# ==========================END of KPI page ==========================================================================================

# ==========================Function of OEE page ====================================================================================
# ==========================Function of OEE page ====================================================================================
# ==========================Function of OEE page ====================================================================================
    @QtCore.pyqtSlot()
    def OEE_page(self):
        self.ui.main_stacked.setCurrentWidget(self.ui.OEE_page)
        self.set_stylesheet_change_page((self.ui.OEE_btn, self.ui.Home_btn, self.ui.Maintenance_btn,
                                        self.ui.KPI_btn,
                                        self.ui.Stock_btn, self.ui.Downtime_btn))
        if not self.is_expanded:
            self.is_expanded = True
            self.expand_windown_animation(self.is_expanded)
        if self.ui.OEE_area_cbb.count() > 0:
            return
        self.OEE_dashboard_page(Dashboard_init=True)
        self.safe_connect(self.ui.OEE_dashboard_btn.clicked, lambda: self.OEE_dashboard_page(Dashboard_init=True))
        self.safe_connect(self.ui.OEE_data_btn.clicked, lambda: self.OEE_detail_page())

    @QtCore.pyqtSlot()
    def OEE_dashboard_page(self,Dashboard_init = False):
        self.ui.OEE_stacked_widget.setCurrentWidget(self.ui.OEE_Dashboard_page)
        self.style_button_with_shadow((self.ui.OEE_dashboard_btn, self.ui.OEE_data_btn))
        self.ui.OEE_period_edit.setDate(QtCore.QDate(self.ui.today.year, self.ui.today.month-1, 1))
        if not hasattr(self, "OEE_chart_toogle_btn"):
            self.OEE_chart_toogle_btn = ToggleSwitch()
            self.ui.OEE_chart_toogle_widget.setLayout(QtWidgets.QHBoxLayout())
            self.ui.OEE_chart_toogle_widget.layout().addWidget(self.OEE_chart_toogle_btn)
        if Dashboard_init and self.ui.OEE_area_cbb.count() > 0:
            return
        try:
            if not hasattr(self, "areas") or not self.areas:
                self.areas = [area[0] for area in self.database_process.query(sql='''SELECT downtime_area_name
                                                                                    FROM `downtime_areas`;''')]
            self.ui.OEE_area_cbb.clear()
            self.ui.OEE_area_cbb.addItems(self.areas)
            year_exist = self.database_process.query(sql='''SELECT DISTINCT YEAR(production_date) as year
                                                            FROM `production_output`
                                                            ORDER BY year DESC;''')
            self.ui.OEE_year_edit.clear()
            self.ui.OEE_year_edit.addItems([str(year[0]) for year in year_exist])
            self.ui.OEE_year_edit.setCurrentText(str(self.ui.today.year))
            self.OEE_filter_changing(change_object="area", area_cbb = self.ui.OEE_area_cbb, model_cbb = self.ui.OEE_model_cbb, line_cbb = self.ui.OEE_line_cbb, process_cbb = self.ui.OEE_process_cbb, period_cbb = self.ui.OEE_period_edit)
            self.OEE_filter_changing(change_object="overview_mode", area_cbb = self.ui.OEE_area_cbb, model_cbb = self.ui.OEE_model_cbb, line_cbb = self.ui.OEE_line_cbb, process_cbb = self.ui.OEE_process_cbb, period_cbb = self.ui.OEE_period_edit)
            self.safe_connect(self.ui.OEE_calendar_widget.currentPageChanged, lambda year,
                            month: self.update_date_from_calendar(year, month, self.ui.OEE_period_edit))
            self.safe_connect(self.ui.OEE_area_cbb.currentTextChanged, lambda: self.OEE_filter_changing(change_object="area", area_cbb = self.ui.OEE_area_cbb, model_cbb = self.ui.OEE_model_cbb, line_cbb = self.ui.OEE_line_cbb, process_cbb = self.ui.OEE_process_cbb, period_cbb = self.ui.OEE_period_edit))
            self.safe_connect(self.ui.OEE_model_cbb.currentTextChanged, lambda: self.OEE_filter_changing(change_object="model", area_cbb = self.ui.OEE_area_cbb, model_cbb = self.ui.OEE_model_cbb, line_cbb = self.ui.OEE_line_cbb, process_cbb = self.ui.OEE_process_cbb, period_cbb = self.ui.OEE_period_edit))
            self.safe_connect(self.ui.OEE_line_cbb.currentTextChanged, lambda: self.OEE_filter_changing(change_object="line", area_cbb = self.ui.OEE_area_cbb, model_cbb = self.ui.OEE_model_cbb, line_cbb = self.ui.OEE_line_cbb, process_cbb = self.ui.OEE_process_cbb, period_cbb = self.ui.OEE_period_edit))
            self.safe_connect(self.ui.OEE_load_filter_btn.clicked, lambda: self.OEE_load_filter())
            self.safe_connect(self.ui.OEE_month_radio.clicked, lambda: self.OEE_filter_changing(change_object="period", area_cbb = self.ui.OEE_area_cbb, model_cbb = self.ui.OEE_model_cbb, line_cbb = self.ui.OEE_line_cbb, process_cbb = self.ui.OEE_process_cbb, period_cbb = self.ui.OEE_period_edit))
            self.safe_connect(self.ui.OEE_year_radio.clicked, lambda: self.OEE_filter_changing(change_object="period", area_cbb = self.ui.OEE_area_cbb, model_cbb = self.ui.OEE_model_cbb, line_cbb = self.ui.OEE_line_cbb, process_cbb = self.ui.OEE_process_cbb, period_cbb = self.ui.OEE_year_edit))
            self.safe_connect(self.ui.OEE_period_edit.dateChanged, lambda: self.OEE_filter_changing(change_object="period", area_cbb = self.ui.OEE_area_cbb, model_cbb = self.ui.OEE_model_cbb, line_cbb = self.ui.OEE_line_cbb, process_cbb = self.ui.OEE_process_cbb, period_cbb = self.ui.OEE_period_edit))
            self.safe_connect(self.ui.OEE_overview_mode_btn.clicked, lambda: self.OEE_filter_changing(change_object="overview_mode", area_cbb = self.ui.OEE_area_cbb ,period_cbb = self.ui.OEE_period_edit))
            def detail_mode_change():
                self.OEE_filter_changing(change_object="detail_mode",  area_cbb = self.ui.OEE_area_cbb, model_cbb = self.ui.OEE_model_cbb, line_cbb = self.ui.OEE_line_cbb, process_cbb = self.ui.OEE_process_cbb, period_cbb = self.ui.OEE_period_edit)
                self.OEE_load_filter()
            self.safe_connect(self.ui.OEE_detail_mode_btn.clicked, lambda: detail_mode_change())
            self.ui.OEE_comment_text_edit.dropEvent = lambda event: self.action_text_drop_event(widget=self.ui.OEE_comment_text_edit, event=event)
            self.ui.OEE_comment_text_edit.mousePressEvent = lambda event: self.action_text_mouse_press_event(widget=self.ui.OEE_comment_text_edit, event=event)
            self.ui.OEE_comment_text_edit.setMouseTracking(True)
            self.ui.OEE_comment_text_edit.mouseMoveEvent = lambda event: self.action_text_mouse_move_event(widget=self.ui.OEE_comment_text_edit, event=event)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load OEE page: {e}")

    def OEE_overview_dashboard(self, area_name, month, year):
        self.spinner.start()
        QtCore.QTimer.singleShot(15000,self.spinner.stop)
        area_name = area_name
        month = month
        year = year
        if self.ui.OEE_scroll_widget_layout.count() > 0:
            old_layout = self.ui.OEE_scroll_widget_layout
            while old_layout.count():
                item = old_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
        if month == 0:
            year = int(year)
            month = self.ui.today.month if year == self.ui.today.year else 12
            period_query = "YEAR(oee.production_date)"
            filter_scripts = f''' AND {period_query} = :year '''
            filter_scripts_dt = ''' AND YEAR(Date) = :year'''
            filter_scripts_wt = f'''AND YEAR(lot.operation_date) = :year'''
        else:
            period_query = "MONTH(oee.production_date)"
            filter_scripts = ''' AND MONTH(oee.production_date) = :month AND YEAR(oee.production_date) = :year '''
            filter_scripts_dt = ''' AND MONTH(Date) = :month AND YEAR(Date) = :year'''
            filter_scripts_wt = f'''AND MONTH(lot.operation_date) = :month AND YEAR(lot.operation_date) = :year'''

        def fetch_data():
            oee_data = self.database_process.query(sql=f'''WITH oee_data AS (
                                                            SELECT oee.area_name, oee.line_name, oee.model_name, oee.process, mct.cycle_time_seconds , {period_query} AS production_month,
                                                                SUM(oee.planed_time) AS total_planed_time, SUM(oee.OK_qty) AS total_OK_qty,
                                                                SUM(oee.NG_qty) AS total_NG_qty, SUM(oee.Available_Time) AS total_Available_Time
                                                            FROM `oee_report` as oee
                                                            JOIN `product_models_oee` as pmo ON oee.model_name = pmo.model_name
                                                            JOIN `production_lines` as pl ON oee.line_name = pl.line_name
                                                            JOIN `machine_oee_register` AS mor ON pmo.model_id = mor.model_id AND mor.process = oee.process AND mor.line_id = pl.line_id
                                                            JOIN machine_cycle_times as mct ON mct.machine_id = mor.machine_id AND mct.create_at <= oee.production_date AND mct.model_id  = pmo.model_id
                                                            LEFT JOIN machine_cycle_times as mct2 ON pmo.model_id = mct2.model_id AND mct2.create_at <= oee.production_date
                                                            AND mct2.machine_id = mor.machine_id AND mct2.create_at > mct.create_at
                                                            WHERE mct2.machine_id IS NULL
                                                            AND oee.area_name = :area_name {filter_scripts}
                                                            GROUP BY oee.area_name, oee.line_name, oee.model_name, oee.process, mct.cycle_time_seconds, production_month
                                                        )
                                                        SELECT oee_data.*, 
                                                            ROUND((oee_data.total_Available_Time / NULLIF(oee_data.total_planed_time, 0)), 4) AS Availability_percentage,
                                                            ROUND(((oee_data.total_OK_qty + oee_data.total_NG_qty) * oee_data.cycle_time_seconds / (NULLIF(oee_data.total_Available_Time, 0) * 60)), 4) AS Performance_percentage,
                                                            ROUND((oee_data.total_OK_qty / NULLIF((oee_data.total_OK_qty + oee_data.total_NG_qty), 0)), 4) AS Quality_percentage,
                                                            ROUND(((oee_data.total_Available_Time / NULLIF(oee_data.total_planed_time, 0)) *
                                                                    ((oee_data.total_OK_qty + oee_data.total_NG_qty) * oee_data.cycle_time_seconds / (NULLIF(oee_data.total_Available_Time, 0)*60)) *
                                                                    (oee_data.total_OK_qty / NULLIF((oee_data.total_OK_qty + oee_data.total_NG_qty), 0))), 4) AS OEE_percentage
                                                        FROM oee_data
                                                        ORDER BY oee_data.line_name, oee_data.model_name, oee_data.process;
                                                    ''' ,
                                                    params={"area_name": area_name, "month": month, "year": year})
            oee_target = self.database_process.query(sql=f'''SELECT pl.line_name, model.model_name, target.process, target.availability_target, target.performance_target, target.quality_target, target.OEE_target
                                                            FROM oee_targets AS target
                                                            JOIN downtime_areas AS da ON da.downtime_area_id = target.downtime_area_id
                                                            LEFT JOIN production_lines AS pl ON target.line_id = pl.line_id
                                                            LEFT JOIN product_models_oee AS model ON target.model_id = model.model_id
                                                            WHERE da.downtime_area_name = :area_name;''',
                                                    params={"area_name": area_name})
            mttr_mtbf_area_target = self.database_process.query(sql=f'''SELECT dttg.mttr_target_value, dttg.mtbf_target_value
                                                                    FROM mttr_mtbf_targets AS dttg
                                                                    JOIN downtime_areas AS da ON dttg.downtime_area_id = da.downtime_area_id
                                                                    WHERE da.downtime_area_name = :area_name AND dttg.created_at <= ":year-:month-28"
                                                                    AND dttg.line_id IS NULL AND dttg.machine_id IS NULL
                                                                    ORDER BY dttg.created_at DESC
                                                                    LIMIT 1; ''', params={"area_name": area_name, "month": month, "year": year})
            mttr_mtbf_line_target = self.database_process.query(sql=f'''SELECT pl.line_name, dttg.mttr_target_value, dttg.mtbf_target_value
                                                                    FROM mttr_mtbf_targets AS dttg
                                                                    JOIN production_lines AS pl ON dttg.line_id = pl.line_id
                                                                    JOIN downtime_areas AS da ON dttg.downtime_area_id = da.downtime_area_id
                                                                    WHERE da.downtime_area_name = :area_name AND dttg.created_at <= ":year-:month-28"
                                                                    ORDER BY dttg.created_at DESC; ''', params={"area_name": area_name, "month": month, "year": year})
            total_model_iatf_count = self.database_process.query(sql=f'''SELECT COUNT(DISTINCT model_id) as total_model_iatf_count
                                                                        FROM product_models_oee  AS pmo
                                                                        JOIN departments as d ON pmo.department_id = d.department_id
                                                                        JOIN downtime_areas as da ON d.department_id = da.department_id
                                                                        WHERE da.downtime_area_name = :area_name;''', params={"area_name": area_name})
            downtime_data = self.database_process.query(sql=f'''SELECT Line_Name, SUM(Total_Loss) AS total_loss, SUM(Repair_Time) AS total_repair_time, COUNT(*) AS total_records
                                                                    FROM `downtime_report` AS dr
                                                                    WHERE Downtime_Area = :area_name {filter_scripts_dt}
                                                                    GROUP BY Line_Name;''', params={"area_name": area_name, "month": month, "year": year})
            working_time = self.database_process.query(sql=f'''SELECT pl.line_name as Line_Name , 
                                                                SUM(lot.operation_hours)*60 - COALESCE(SUM(setup_time),0) - COALESCE(SUM(break_time),0) AS total_planed_time 
                                                                FROM `line_operation_times` as lot
                                                                JOIN downtime_areas_production_lines as dapl ON lot.line_id = dapl.line_id
                                                                JOIN downtime_areas as da ON dapl.downtime_area_id = da.downtime_area_id
                                                                JOIN production_lines as pl ON lot.line_id = pl.line_id
                                                                WHERE da.downtime_area_name = :area_name {filter_scripts_wt}
                                                                GROUP BY pl.line_name ;''', params={"area_name": area_name, "month": month, "year": year})
            return {"oee_data": oee_data, "oee_target": oee_target, "dt_line_target": mttr_mtbf_line_target, "dt_area_target": mttr_mtbf_area_target, "model_count": total_model_iatf_count[0][0] if total_model_iatf_count else 0, "downtime_data": downtime_data, "working_time": working_time}
        
        def on_data_fetched(data):
            oee_data = data["oee_data"]
            oee_target = data["oee_target"]
            model_count = data["model_count"]
            downtime_data = data["downtime_data"]
            dt_line_target = data["dt_line_target"]
            dt_area_target = data["dt_area_target"]
            working_time = data["working_time"]
            total_oee = pd.DataFrame(oee_data, columns=["area_name", "line_name", "model_name", "process", "cycle_time_seconds" ,"production_date", "planed_time", "OK_qty",
                                       "NG_qty", "Available_Time", "Availability_percentage", "Performance_percentage", "Quality_percentage", "OEE_percentage"])
            if total_oee.empty or model_count == 0 or downtime_data is None or working_time is None:
                self.spinner.stop()
                QtWidgets.QMessageBox.warning(self, "No Data", "No data available for the selected filters.")
                self.ui.OEE_area_value.clear()
                self.ui.OEE_highest_perfom_value.clear()
                self.ui.OEE_lowest_perfom_value.clear()
                self.ui.Number_production_model.clear()
                if self.ui.OEE_scroll_widget_layout is not None:
                    layout = self.ui.OEE_scroll_widget_layout
                    while layout.count():
                        child = layout.takeAt(0)
                        if child.widget():
                            child.widget().deleteLater()
            for col in total_oee.columns:
                if col not in ["area_name", "line_name", "model_name", "process"]:
                    total_oee[col] = total_oee[col].astype(float)
            downtime_df = pd.DataFrame(downtime_data, columns=["line_name", "total_loss", "total_repair_time", "total_records"])
            working_time_df = pd.DataFrame(working_time, columns=["line_name", "total_planed_time"])
            for col in downtime_df.columns:
                if col not in ["line_name"]:
                    downtime_df[col] = downtime_df[col].astype(float)
            for col in working_time_df.columns:
                if col not in ["line_name"]:
                    working_time_df[col] = working_time_df[col].astype(float)
            dt_line_target_df = pd.DataFrame(dt_line_target, columns=["line_name", "mttr_target_value", "mtbf_target_value"])
            dt_line_target_df["mttr_target_value"] = dt_line_target_df["mttr_target_value"].astype(float)
            dt_line_target_df["mtbf_target_value"] = dt_line_target_df["mtbf_target_value"].astype(float)
            dt_area_target_df = {
                "mttr_target_value": float(dt_area_target[0][0]) if dt_area_target else None,
                "mtbf_target_value": float(dt_area_target[0][1]) if dt_area_target else None
            }
            oee_target_df = pd.DataFrame(oee_target, columns=["line_name", "model_name", "process", "availability_target", "performance_target", "quality_target", "OEE_target"])
            oee_target_df["availability_target"] = oee_target_df["availability_target"].astype(float)
            oee_target_df["performance_target"] = oee_target_df["performance_target"].astype(float)
            oee_target_df["quality_target"] = oee_target_df["quality_target"].astype(float)
            oee_target_df["OEE_target"] = oee_target_df["OEE_target"].astype(float)
            oee_target_df.fillna({"line_name": area_name}, inplace=True)
            for i in range(len(total_oee)):
                line_name = total_oee.loc[i, "line_name"]
                mttr_value = downtime_df.loc[(downtime_df["line_name"] == line_name), "total_repair_time"].values[0] / (downtime_df.loc[(downtime_df["line_name"] == line_name) , "total_records"].values[0] \
                                if not downtime_df.empty and not downtime_df.loc[(downtime_df["line_name"] == line_name)].empty else 1)
                mttr_target_value = dt_line_target_df.loc[dt_line_target_df["line_name"] == line_name, "mttr_target_value"].values[0] \
                                if not dt_line_target_df.empty and not dt_line_target_df.loc[dt_line_target_df["line_name"] == line_name].empty else dt_area_target_df["mttr_target_value"]
                mtbf_value =  ( working_time_df.loc[(working_time_df["line_name"] == line_name), "total_planed_time"].values[0] - downtime_df.loc[(downtime_df["line_name"] == line_name), "total_loss"].values[0]) /  \
                                (downtime_df.loc[(downtime_df["line_name"] == line_name), "total_records"].values[0] if not downtime_df.empty and not downtime_df.loc[(downtime_df["line_name"] == line_name)].empty else 1)
                mtbf_target_value = dt_line_target_df.loc[dt_line_target_df["line_name"] == line_name, "mtbf_target_value"].values[0] \
                                if not dt_line_target_df.empty and not dt_line_target_df.loc[dt_line_target_df["line_name"] == line_name].empty else dt_area_target_df["mtbf_target_value"]
                total_oee.loc[i, "mttr_value"] = mttr_value
                total_oee.loc[i, "mttr_target"] = mttr_target_value
                total_oee.loc[i, "mtbf_value"] = mtbf_value
                total_oee.loc[i, "mtbf_target"] = mtbf_target_value
            max_mttr_value = max(total_oee["mttr_value"].max(), total_oee["mttr_target"].max())
            max_mtbf_value = max(total_oee["mtbf_value"].max(), total_oee["mtbf_target"].max())
            oee_avr_percentage = total_oee["OEE_percentage"].mean()
            highest_performing_line = total_oee.loc[total_oee["OEE_percentage"].idxmax()]["line_name"]
            lowest_performing_line = total_oee.loc[total_oee["OEE_percentage"].idxmin()]["line_name"]
            model_runned = total_oee["model_name"].unique()
            mttr_target_value = dt_area_target_df["mttr_target_value"]
            mtbf_target_value = dt_area_target_df["mtbf_target_value"]
            self.ui.OEE_area_target.setText(f"{oee_target_df.loc[oee_target_df['line_name'] == area_name, 'OEE_target'].values[0]}%")
            self.ui.MTTR_area_target.setText(f"{mttr_target_value:.1f} min" if mttr_target_value is not None else "N/A")
            self.ui.MTBF_area_target.setText(f"{mtbf_target_value:.1f} min" if mtbf_target_value is not None else "N/A")
            self.ui.Number_production_model.setText(f"{len(model_runned)}/{model_count}")
            self.ui.OEE_area_value.setText(f"{oee_avr_percentage*100:.2f}%")
            self.ui.OEE_highest_perfom_value.setText(f"{highest_performing_line} ({total_oee['OEE_percentage'].max()*100:.1f}%)")
            self.ui.OEE_lowest_perfom_value.setText(f"{lowest_performing_line} ({total_oee['OEE_percentage'].min()*100:.1f}%)")
            mini_card_widgets_list = []
            def on_mini_card_clicked(line_name, model_name, process):
                self.ui.OEE_dashboard_view_stacked_widget.setCurrentWidget(self.ui.OEE_detail_view_page)
                self.ui.OEE_line_cbb.setCurrentText(line_name)
                self.ui.OEE_model_cbb.setCurrentText(model_name)
                self.ui.OEE_process_cbb.setCurrentText(process)
                self.ui.OEE_detail_mode_btn.setChecked(True)
                self.OEE_filter_changing(change_object="detail_mode",clicked = False)
                self.OEE_load_filter()

            for i in range(len(total_oee)):
                data_prepare = {
                    "line_name": total_oee.loc[i, "line_name"],
                    "model_name": total_oee.loc[i, "model_name"],
                    "process": total_oee.loc[i, "process"],
                    "availability_value": total_oee.loc[i, "Availability_percentage"],
                    "performance_value": total_oee.loc[i, "Performance_percentage"],
                    "quality_value": total_oee.loc[i, "Quality_percentage"],
                    "oee_value": total_oee.loc[i, "OEE_percentage"],
                    "oee_target": oee_target_df.loc[(oee_target_df["line_name"] == total_oee.loc[i, "line_name"]) & (oee_target_df["model_name"] == total_oee.loc[i, "model_name"]) & (oee_target_df["process"] == total_oee.loc[i, "process"]), "OEE_target"].values[0] \
                                    if not oee_target_df.loc[(oee_target_df["line_name"] == total_oee.loc[i, "line_name"]) & (oee_target_df["model_name"] == total_oee.loc[i, "model_name"]) & (oee_target_df["process"] == total_oee.loc[i, "process"])].empty else 0,
                    "mttr_value": total_oee.loc[i, "mttr_value"],
                    "mttr_target": total_oee.loc[i, "mttr_target"],
                    "mtbf_value": total_oee.loc[i, "mtbf_value"],
                    "mtbf_target": total_oee.loc[i, "mtbf_target"],
                    "max_mttr_value": max_mttr_value *1.2,
                    "max_mtbf_value": max_mtbf_value*1.2
                }
                mini_card_widget = Mini_Card(data=data_prepare, format_time=self.change_time_format)
                mini_card_widget.doubleClicked.connect(lambda line_name=data_prepare["line_name"], model_name=data_prepare["model_name"], process=data_prepare["process"]: on_mini_card_clicked(line_name, model_name, process))
                mini_card_widgets_list.append(mini_card_widget)
                self.ui.OEE_scroll_widget_layout.addWidget(mini_card_widget)
            self.ui.OEE_scroll_widget_layout.addStretch()
            self.spinner.stop()
        try:
            if hasattr(self, 'OEE_dashboard_worker') and self.OEE_dashboard_worker.isRunning():
                self.spinner.stop()
                return
            self.OEE_dashboard_worker = WorkerThread(fetch_data)
            self.OEE_dashboard_worker.finished.connect(lambda res: on_data_fetched(res))
            self.OEE_dashboard_worker.start()
        except Exception as e:
            self.spinner.stop()
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to refresh OEE Overview page: {e}")

    def change_time_format(self,time_value: float, input_unit: str, output_unit: bool = False):
        time_val_f = float(time_value)
        if input_unit == "h":
            hours = int(time_val_f)
            minutes = int((time_val_f - hours) * 60)
            seconds = int(round((time_val_f - hours - minutes/60) * 3600))
        elif input_unit == "m":
            hours = int(time_val_f // 60)
            minutes = int(time_val_f % 60)
            seconds = int(round((time_val_f - int(time_val_f)) * 60))
        else:
            hours = int(time_val_f // 3600)
            minutes = int((time_val_f % 3600) // 60)
            seconds = int(time_val_f % 60)
        if output_unit:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return {
            "h": f"{hours:02d}",
            "m": f"{minutes:02d}",
            "s": f"{seconds:02d}"
        }

    def refesh_OEE_page(self, area_name, model_name, month, year, line, process):
        self.spinner.start()
        QtCore.QTimer.singleShot(15000,self.spinner.stop)
        if month == 0:
            filter_scripts = ''' AND YEAR(production_date) = :year '''
            filter_scripts_wt = f'''AND YEAR(lot.operation_date) = :year'''
            filter_scripts_dt = '''AND YEAR(Date) = :year'''
            filter_scripts_comment = '''AND dact.for_month is NULL AND dact.for_year = :for_year'''
            previous_year = year - 1
            month = 12
            previous_month = month
            ismonthly = False
        else:
            filter_scripts = ''' AND MONTH(ore.production_date) = :month AND YEAR(ore.production_date) = :year '''
            filter_scripts_wt = f'''AND MONTH(lot.operation_date) = :month AND YEAR(lot.operation_date) = :year'''
            filter_scripts_dt = ''' AND MONTH(Date) = :month AND YEAR(Date) = :year'''
            filter_scripts_comment = ''' AND dact.for_month = :for_month AND dact.for_year = :for_year'''
            previous_month = month - 1 if month > 1 else 12
            previous_year = year if month > 1 else year - 1
            ismonthly = True

        def fetch_data():
            oee_data = self.database_process.query(sql=f'''SELECT ore.area_name, ore.line_name, ore.model_name, ore.process, ore.production_date, ore.operation_hours, ore.planed_time, 
                                                            ore.OK_qty, ore.NG_qty, ore.Total_Loss, ore.Repair_Time, ore.Available_Time, ore.Availability_percentage, ore.Performance_percentage, 
                                                            ore.Quality_percentage, ore.OEE_percentage
                                                    FROM `oee_report` as ore
                                                    WHERE ore.area_name = :area_name 
                                                    AND ore.model_name = :model_name 
                                                    {filter_scripts}
                                                    AND ore.line_name = :line AND ore.process = :process;''', params={"area_name": area_name, "model_name": model_name, "month": month, "year": year, "line": line, "process": process})
            prev_working_time = self.database_process.query(sql=f'''SELECT SUM(lot.operation_hours)*60 - COALESCE(SUM(setup_time),0) - COALESCE(SUM(break_time),0) AS total_planed_time 
                                                                FROM `line_operation_times` as lot
                                                                JOIN downtime_areas_production_lines as dapl ON lot.line_id = dapl.line_id
                                                                JOIN downtime_areas as da ON dapl.downtime_area_id = da.downtime_area_id
                                                                JOIN production_lines as pl ON lot.line_id = pl.line_id
                                                                WHERE da.downtime_area_name = :area_name AND pl.line_name = :line {filter_scripts_wt};''', params={"area_name": area_name, "month": previous_month, "year": previous_year, "line": line})
            working_time = self.database_process.query(sql=f'''SELECT SUM(lot.operation_hours)*60 - COALESCE(SUM(setup_time),0) - COALESCE(SUM(break_time),0) AS total_planed_time 
                                                                FROM `line_operation_times` as lot
                                                                JOIN downtime_areas_production_lines as dapl ON lot.line_id = dapl.line_id
                                                                JOIN downtime_areas as da ON dapl.downtime_area_id = da.downtime_area_id
                                                                JOIN production_lines as pl ON lot.line_id = pl.line_id
                                                                WHERE da.downtime_area_name = :area_name AND pl.line_name = :line {filter_scripts_wt};''', params={"area_name": area_name, "month": month, "year": year, "line": line})
            oee_targets = self.database_process.query(sql=f'''SELECT availability_target, performance_target, quality_target, oee_target 
                                                                FROM `oee_targets` as ot
                                                                JOIN `product_models_oee` as pmo ON ot.model_id = pmo.model_id
                                                                JOIN `production_lines` as pl ON ot.line_id = pl.line_id
                                                                WHERE pmo.model_name = :model_name
                                                                AND pl.line_name = :line AND ot.process = :process
                                                                AND ot.date_created <= :date
                                                                ORDER BY ot.date_created DESC
                                                                LIMIT 1;''', params={"model_name": model_name, "line": line, "process": process, "date": f"{year}-{month:02d}-{calendar.monthrange(year, month)[1]}"})
            downtime_target = self.database_process.query(sql=f'''SELECT mttr_target_value, mtbf_target_value
                                                                FROM `mttr_mtbf_targets` as mmt
                                                                JOIN `production_lines` as pl ON mmt.line_id = pl.line_id
                                                                WHERE pl.line_name = :line;''', params={"line": line})
            previous_oee_data = self.database_process.query(sql=f'''SELECT ore.area_name, ore.line_name, ore.model_name, ore.process, MONTH(ore.production_date), SUM(ore.planed_time), SUM(ore.`OK_qty`),
                                                                SUM(`NG_qty`), SUM(`Total_Loss`), SUM(`Repair_Time`), SUM(`Available_Time`), AVG(`Availability_percentage`), AVG(`Performance_percentage`),
                                                                AVG(`Quality_percentage`), AVG(`OEE_percentage`) 
                                                                FROM `oee_report` as ore
                                                                WHERE ore.area_name = :area_name
                                                                AND ore.model_name = :model_name
                                                                {filter_scripts}
                                                                AND ore.line_name = :line AND ore.process = :process
                                                                GROUP BY ore.area_name, ore.line_name, ore.model_name, ore.process, MONTH(ore.production_date);''',
                                                            params={"area_name": area_name, "model_name": model_name, "month": previous_month, "year": previous_year, "line": line, "process": process})
            cycle_time = self.database_process.query(sql='''SELECT mct.create_at, mct.cycle_time_seconds
                                                            FROM machine_cycle_times AS mct
                                                            JOIN (
                                                                SELECT pmo.model_id, mor.machine_id
                                                                FROM machine_oee_register AS mor
                                                                JOIN product_models_oee AS pmo ON mor.model_id = pmo.model_id
                                                                JOIN production_lines AS pl ON mor.line_id = pl.line_id
                                                                WHERE pmo.model_name = :model_name AND pl.line_name = :line_name AND mor.process = :process
                                                                GROUP BY mor.machine_id
                                                            ) AS subquery ON mct.machine_id = subquery.machine_id AND mct.model_id = subquery.model_id
                                                            WHERE mct.create_at <= :date
                                                            ORDER BY mct.create_at DESC
                                                            LIMIT 1;
                                                            ''', params={"process": process, "model_name": model_name, "line_name": line, "date": f"{year}-{month:02d}-{calendar.monthrange(year, month)[1]}"})      
            downtime_data = self.database_process.query(sql=f'''SELECT SUM(Total_Loss) AS total_loss_time, SUM(Repair_Time) AS total_repair_time, COUNT(*) AS total_records
                                                                    FROM `downtime_report` AS dr
                                                                    WHERE Downtime_Area = :area_name {filter_scripts_dt}
                                                                    AND Line_Name = :line_name;''', params={"area_name": area_name, "month": month, "year": year, "line_name": line, "model_name": model_name, "process": process})
            downtime_data_previous = self.database_process.query(sql=f'''SELECT SUM(Total_Loss) AS total_loss_time, SUM(Repair_Time) AS total_repair_time, COUNT(*) AS total_records
                                                                    FROM `downtime_report` AS dr
                                                                    WHERE Downtime_Area = :area_name {filter_scripts_dt}
                                                                    AND Line_Name = :line_name AND Current_Model = :model_name;''', params={"area_name": area_name, "month": previous_month, "year": previous_year, "line_name": line, "model_name": model_name, "process": process})       
            OEE_comment = self.database_process.query(sql=f'''SELECT action_id, action_content, action_report_link
                                                                    FROM downtime_actions as dact
                                                                    JOIN downtime_areas as da ON dact.downtime_area_id = da.downtime_area_id
                                                                    JOIN production_lines as pl ON dact.line_id = pl.line_id
                                                                    WHERE pl.line_name = :line_name AND da.downtime_area_name = :area_name {filter_scripts_comment} AND action_for = "OEE";''',
                                                        params={"line_name": line, "area_name": area_name, "for_month": month, "for_year": year})
            return (oee_data, oee_targets, previous_oee_data, cycle_time, downtime_data, downtime_data_previous, downtime_target, OEE_comment, working_time, prev_working_time)
        
        def on_data_fetched(result):
            oee_data = result[0]
            oee_targets = result[1]
            previous_oee_data = result[2]
            cycle_time = result[3]
            downtime_data = result[4]
            downtime_data_previous = result[5]
            mttr_target = result[6][0][0] if result[6] else None
            mtbf_target = result[6][0][1] if result[6] else None
            OEE_comment = result[7]
            working_time = float(result[8][0][0]) if result[8][0][0] is not None else None
            prev_working_time = float(result[9][0][0]) if result[9][0][0] is not None else None
            self.ui.OEE_comment_text_edit.clear()
            if not OEE_comment:
                self.OEE_comment_id = None
                self.has_OEE_comment = False
            else:
                self.ui.OEE_comment_text_edit.setText(OEE_comment[0][1])
                comment_link = json.loads(OEE_comment[0][2]) if OEE_comment[0][2] else []
                for link in comment_link:
                    self.ui.OEE_comment_text_edit.append(
                        f'<a href="{link["file_path"]}" style="color:#007acc; font-size:10px; text-decoration:underline;">{link["file_name"]}</a> ')
                self.OEE_comment_id = OEE_comment[0][0]
                self.has_OEE_comment = True
            self.ui.OEE_action_save_btn.setEnabled(False)
            self.ui.OEE_action_save_btn.setHidden(True)
            if not oee_data:
                QtWidgets.QMessageBox.information(self, "No Data", "No OEE data found for the selected criteria.")
                self.ui.OEE_OK_qty_lbl.clear()
                self.ui.OEE_NG_qty_lbl.clear()
                self.ui.OEE_WT_lbl.clear()
                self.ui.OEE_DT_lbl.clear()
                self.ui.OEE_machine_cycletime_lbl.clear()
                self.ui.OEE_MTTR_value_lbl.clear()
                self.ui.pre_MTTR_lbl.clear()
                self.ui.OEE_MTBF_value_lbl.clear()
                self.ui.pre_MTBF_lbl.clear()
                self.ui.pre_OEE_lbl.clear()
                self.ui.pre_A_lbl.clear()
                self.ui.pre_P_lbl.clear()
                self.ui.pre_Q_lbl.clear()
                self.ui.OEE_comment_text_edit.clear()
                self.ui.OEE_keyinsights_lbl.clear()
                self.ui.OEE_action_save_btn.setEnabled(False)
                self.ui.OEE_action_save_btn.setHidden(True)
                list_of_charts = [ self.ui.OEE_information_chart, self.ui.OEE_A_value_chart, self.ui.OEE_P_value_chart, self.ui.OEE_Q_value_chart, self.ui.OEE_value_chart]
                for chart in list_of_charts:
                    if chart.layout() is None or chart.layout().count() == 0:
                        continue
                    chart.layout().takeAt(0).widget().deleteLater()
                self.spinner.stop()
                return
            self.oee_df = pd.DataFrame(oee_data, columns=["area_name", "line_name", "model_name", "process", "production_date", "operation_hours", "planed_time", "OK_qty",
                                       "NG_qty", "Total_Loss", "Repair_Time", "Available_Time", "Availability_percentage", "Performance_percentage", "Quality_percentage", "OEE_percentage"])
            self.downtime_df = pd.DataFrame(downtime_data, columns=["total_loss_time", "total_repair_time", "total_records"])
            self.downtime_df_previous = pd.DataFrame(downtime_data_previous, columns=["total_loss_time", "total_repair_time", "total_records"])
            cols = ["total_loss_time","total_repair_time","total_records"]
            self.downtime_df_previous[cols] = (
                self.downtime_df_previous[cols]
                .astype("float64")
                .fillna(0)
            )
            self.oee_df.fillna({"planed_time": 0, "OK_qty": 0, "NG_qty": 0, "Total_Loss": 0, "Repair_Time": 0, "Available_Time": 0, "Availability_percentage": 0, "Performance_percentage": 0, "Quality_percentage": 0, "OEE_percentage": 0}, inplace=True)
            self.downtime_df.fillna({"total_loss_time": 0, "total_repair_time": 0, "total_records": 0}, inplace=True)
            total_OK_qty = float(self.oee_df['OK_qty'].sum())
            total_NG_qty = float(self.oee_df['NG_qty'].sum())
            total_operation_hours = float(self.oee_df['planed_time'].sum())
            total_loss = float(self.oee_df['Total_Loss'].sum())
            cycle_time_value = float(cycle_time[0][1]) if cycle_time else 0
            total_runtime = float(self.oee_df['Available_Time'].sum())
            total_A_percentage = (total_runtime ) / (total_operation_hours)
            total_P_percentage = (
                cycle_time_value * (total_OK_qty + total_NG_qty)) / (total_runtime*60)
            total_Q_percentage = total_OK_qty / (total_OK_qty + total_NG_qty)
            total_OEE = total_A_percentage * total_P_percentage * total_Q_percentage        
            def change_value_show(value, is_percentage=False):
                if is_percentage:
                    return f"{value*100:.1f}%"
                elif value >= 1000000:
                    return f"{value/1000000:.1f}M"
                elif value >= 1000:
                    return f"{value/1000:.1f}K"
                else:
                    return f"{value:,.0f}"

            self.ui.OEE_OK_qty_lbl.setText(
                f"{change_value_show(total_OK_qty)}pcs")
            self.ui.OEE_NG_qty_lbl.setText(
                f"{change_value_show(total_NG_qty)}pcs")
            self.ui.OEE_WT_lbl.setText(
                f"{round(total_operation_hours/60, 2)} hrs")
            self.ui.OEE_DT_lbl.setText(f"{round(total_loss/60, 2)} hrs")
            self.ui.OEE_machine_cycletime_lbl.setText(
                f"{cycle_time_value:.1f}s/pcs")     
            def draw_circle_chart(value, target, label, chart_widget):
                layout = chart_widget.layout()
                if layout is not None:
                    while layout.count():
                        child = layout.takeAt(0)
                        if child.widget():
                            child.widget().deleteLater()
                    layout.setContentsMargins(0, 0, 0, 0)
                else:
                    layout = QtWidgets.QVBoxLayout(chart_widget)
                    layout.setContentsMargins(0, 0, 0, 0)
                    chart_widget.setLayout(layout)
                chart = DonutChart(target_value=target,
                                   value=value*100, parameter_name=label)
                layout.addWidget(chart)     
            draw_circle_chart(total_OEE, float(oee_targets[0][3]), "%OEE", self.ui.OEE_value_chart)     
            draw_circle_chart(total_A_percentage, 0, "%A",
                              self.ui.OEE_A_value_chart)
            draw_circle_chart(total_P_percentage, 0, "%P",
                              self.ui.OEE_P_value_chart)
            draw_circle_chart(total_Q_percentage, 0, "%Q",
                              self.ui.OEE_Q_value_chart)        
            def make_comparison_html(current: float, previous: float, label: str) -> str:
                diff = current - previous
                diff = round(diff, 1)
                if diff > 0:
                    arrow = "&#9650;"   # ▲
                    color = "#27ae60"   # green
                    color_dt = "#e74c3c"
                    sign = "+"
                elif diff < 0:
                    arrow = "&#9660;"   # ▼
                    color = "#e74c3c"   # red
                    color_dt = "#27ae60"
                    sign = ""
                else:
                    arrow = "&#9654;"   # ►
                    color = "#888888"   # grey
                    color_dt = "#888888"
                    sign = ""
                if label in ["MTTR", "MTBF"]:
                    percent = (1-current/previous) if previous != 0 else 0
                    time_format = self.change_time_format(previous, 'm')
                    return f"""
                    <div style='font-family: Arial; text-align: center;'>
                        <span style='font-size: 11px; font-weight: bold; color: #222;'>
                        Previous:
                        </span>
                        <span style='font-size: 15px; font-weight: bold; color: #222;'>
                            {f"{time_format['h']} hrs {time_format['m']} mins" if previous >= 60 else f"{time_format['m']} mins {time_format['s']} secs"}
                        </span>
                        <span style='font-size: 11px; color: {color_dt if label == "MTTR" else color}; font-weight: bold;'>
                            {arrow} {"-" if percent > 0 else "+"}{abs(percent*100):.1f}% vs prev
                        </span>
                    </div>
                """
                else:
                    return f"""
                    <div style='font-family: Arial; text-align: center;'>
                        <span style='font-size: 11px; font-weight: bold; color: #222;'>
                        Previous:
                        </span>
                        <span style='font-size: 15px; font-weight: bold; color: #222;'>
                            {previous:.1f}% 
                        </span>
                        <span style='font-size: 11px; color: {color}; font-weight: bold;'>
                            {arrow} {sign}{diff:.1f}% vs prev
                        </span>
                    </div>
                """
            pre_mttr = None
            pre_mtbf = None
            if previous_oee_data:
                self.previous_oee_df = pd.DataFrame(previous_oee_data, columns=["area_name", "line_name", "model_name", "process", "month", "planed_time",
                                                    "OK_qty", "NG_qty", "Total_Loss", "Repair_Time", "Available_Time", "Availability_percentage", "Performance_percentage", "Quality_percentage", "OEE_percentage"])
                pre_total_OK_qty = float(self.previous_oee_df['OK_qty'].sum())
                pre_total_NG_qty = float(self.previous_oee_df['NG_qty'].sum())
                pre_total_operation_hours = float(
                    self.previous_oee_df['planed_time'].sum())
                if len(cycle_time) == 2:
                    m1 = (cycle_time[0][0].year, cycle_time[0][0].month)
                    m2 = (cycle_time[1][0].year, cycle_time[1][0].month)
                    if m1 > (previous_year, previous_month) and m2 <= (previous_year, previous_month):
                        pre_cycle_time = float(cycle_time[1][1])
                    else:
                        pre_cycle_time = cycle_time_value
                else:
                    pre_cycle_time = cycle_time_value
                pre_total_runtime = float(
                    self.previous_oee_df['Available_Time'].sum())
                pre_total_A_percentage = pre_total_runtime / \
                    (pre_total_operation_hours)
                pre_total_P_percentage = (
                    pre_cycle_time * (pre_total_OK_qty + pre_total_NG_qty)) / (pre_total_runtime*60)
                pre_total_Q_percentage = pre_total_OK_qty / \
                    (pre_total_OK_qty + pre_total_NG_qty)
                pre_total_OEE = pre_total_A_percentage * \
                    pre_total_P_percentage * pre_total_Q_percentage
                pre_mttr = (self.downtime_df_previous.loc[0, "total_repair_time"] / self.downtime_df_previous.loc[0, "total_records"]) if self.downtime_df_previous.loc[0, "total_records"] > 0 else self.downtime_df_previous.loc[0, "total_repair_time"]
                pre_mtbf = ((prev_working_time - float(self.downtime_df_previous.loc[0, "total_loss_time"])) / float(self.downtime_df_previous.loc[0, "total_records"])) if self.downtime_df_previous.loc[0, "total_records"] > 0 else (prev_working_time - float(self.downtime_df_previous.loc[0, "total_loss_time"]))
                self.ui.pre_OEE_lbl.setText(make_comparison_html(
                    total_OEE*100, pre_total_OEE*100, "OEE"))
                self.ui.pre_A_lbl.setText(make_comparison_html(
                    total_A_percentage*100, pre_total_A_percentage*100, "A"))
                self.ui.pre_P_lbl.setText(make_comparison_html(
                    total_P_percentage*100, pre_total_P_percentage*100, "P"))
                self.ui.pre_Q_lbl.setText(make_comparison_html(
                    total_Q_percentage*100, pre_total_Q_percentage*100, "Q"))
            else:
                self.ui.pre_OEE_lbl.setText("No previous data")
                self.ui.pre_A_lbl.setText("No previous data")
                self.ui.pre_P_lbl.setText("No previous data")
                self.ui.pre_Q_lbl.setText("No previous data")       
            class DateAxisItem(pg.AxisItem):
                    def tickValues(self, minVal, maxVal, size):
                        step = 86400 * 3
                        start = (int(minVal) // step) * step
                        ticks = []
                        v = start
                        while v <= maxVal:
                            ticks.append(v)
                            v += step
                        return [(step, ticks)]

                    def tickStrings(self, values, scale, spacing):
                        return [dt.datetime.fromtimestamp(v).strftime('%d/%m') for v in values]     
            def remove_chart_items(chart_widget,layout_type=QtWidgets.QVBoxLayout):
                layout = chart_widget.layout()
                if layout is not None:
                    while layout.count():
                        child = layout.takeAt(0)
                        if child.widget():
                            child.widget().deleteLater()
                    layout.setContentsMargins(0, 0, 0, 0)
                else:
                    layout = layout_type(chart_widget)
                    layout.setContentsMargins(0, 0, 0, 0)
                    chart_widget.setLayout(layout)
            def safe_div(numer, denom):
                out = np.zeros_like(numer, dtype=float)        
                np.divide(numer, denom, out=out, where=denom > 0)  
                return out
            def draw_OEE_metrics_line_chart(df, target_val, x_lbl, y_lbl, target_widget, legend_widget):
                remove_chart_items(target_widget)
                remove_chart_items(legend_widget, layout_type=QtWidgets.QHBoxLayout)
                layout = target_widget.layout()
                legend_layout = legend_widget.layout()
                legend_layout.setContentsMargins(4, 0, 4, 0)
                legend_layout.setSpacing(12)
                date_axis = DateAxisItem(orientation='bottom')
                line_chart = pg.PlotWidget(axisItems={'bottom': date_axis})
                line_chart.setBackground('w')
                line_chart.hideButtons()
                line_chart.setMouseEnabled(x=False, y=False)
                x = df[x_lbl].apply(lambda d: d.timestamp()
                                    ).values.astype(float)
                x_dense = np.linspace(x[0], x[-1], len(x) * 10)
                OEE = (df[y_lbl] * 100).values.astype(float).round(2)
                OEE_smooth = interp1d(x, OEE, kind='linear')(x_dense)
                target_array = np.full(x_dense.shape, target_val)
                target_line = pg.PlotDataItem(x_dense, target_array, pen=pg.mkPen(color=(
                    0, 206, 209), width=1, style=QtCore.Qt.DashLine), name='Target', antialias=True)
                line_chart.addItem(target_line)
                line_chart.plot(x_dense, OEE_smooth, pen=pg.mkPen(
                    color=(254, 117, 114), width=2), name=y_lbl, antialias=True)
                A = (df['Availability_percentage'] *
                     100).values.astype(float).round(2)
                P = (df['Performance_percentage'] *
                        100).values.astype(float).round(2)
                Q = (df['Quality_percentage'] *
                     100).values.astype(float).round(2)
                if len(x)>= 3:
                    A_smooth = interp1d(x, A, kind='quadratic')(x_dense)
                    P_smooth = interp1d(x, P, kind='quadratic')(x_dense)
                    Q_smooth = interp1d(x, Q, kind='quadratic')(x_dense)

                else:
                    A_smooth = interp1d(x, A, kind='linear')(x_dense)
                    P_smooth = interp1d(x, P, kind='linear')(x_dense)
                    Q_smooth = interp1d(x, Q, kind='linear')(x_dense)
                line_chart.plot(x_dense, A_smooth, pen=pg.mkPen(
                    color=(66, 107, 41), width=2), name='Availability', antialias=True)
                line_chart.plot(x_dense, P_smooth, pen=pg.mkPen(
                    color=(255, 215, 0), width=2), name='Performance', antialias=True)
                line_chart.plot(x_dense, Q_smooth, pen=pg.mkPen(
                    color=(55, 81, 126), width=2), name='Quality', antialias=True)
                OEE_dot_item = pg.ScatterPlotItem(
                    x=x, y=OEE,
                    size=6,
                    pen=pg.mkPen((254, 117, 114), width=1),
                    brush=pg.mkBrush(240, 240, 240),
                    symbol='o',
                    antialias=True
                )
                line_chart.addItem(OEE_dot_item)
                A_dot_item = pg.ScatterPlotItem(
                    x=x, y=A,
                    size=6,
                    pen=pg.mkPen((66, 107, 41), width=1),
                    brush=pg.mkBrush(240, 240, 240),
                    symbol='o',
                    antialias=True
                )
                line_chart.addItem(A_dot_item)
                P_dot_item = pg.ScatterPlotItem(
                    x=x, y=P,
                    size=6,
                    pen=pg.mkPen((255, 215, 0), width=1),
                    brush=pg.mkBrush(240, 240, 240),
                    symbol='o',
                    antialias=True
                )
                line_chart.addItem(P_dot_item)
                Q_dot_item = pg.ScatterPlotItem(
                    x=x, y=Q,
                    size=6,
                    pen=pg.mkPen((55, 81, 126), width=1),
                    brush=pg.mkBrush(240, 240, 240),
                    symbol='o',
                    antialias=True
                )
                line_chart.addItem(Q_dot_item)
                line_chart.setYRange(0, 105, padding=0)
                line_chart.getAxis('left').setTicks(
                    [[(i, str(i)) for i in range(0, 101, 20)]])
                line_chart.getAxis('left').setStyle(tickLength=5)
                line_chart.getAxis('bottom').setStyle(tickLength=5)
                line_chart.showGrid(x=False, y=True, alpha=0.2)
                line_chart.setLabel('left', '<b>Percent (%)</b>', color='#445469', size='10pt',)
                def on_hover(pos):
                    if line_chart.sceneBoundingRect().contains(pos):
                        mouse_point = line_chart.getViewBox().mapSceneToView(pos)
                        x_mouse = mouse_point.x()
                        y_mouse = mouse_point.y()
                        closest_index = np.argmin(np.abs(x - x_mouse))
                        if closest_index < len(x):
                            tooltip_text = f"<b>Date:</b> {dt.datetime.fromtimestamp(x[closest_index]).strftime('%Y-%m-%d')}<br>"
                            tooltip_text += f"<b>OEE:</b> {OEE[closest_index]:.2f}%<br>"
                            tooltip_text += f"<b>Availability:</b> {A[closest_index]:.2f}%<br>"
                            tooltip_text += f"<b>Performance:</b> {P[closest_index]:.2f}%<br>"
                            tooltip_text += f"<b>Quality:</b> {Q[closest_index]:.2f}%"
                            QtWidgets.QToolTip.showText(
                                QtGui.QCursor.pos(), tooltip_text, line_chart)
                line_chart.scene().sigMouseMoved.connect(on_hover)
                layout.addWidget(line_chart)
                if legend_layout.count() > 0:
                    return
                for color, label in [
                    ((254, 117, 114), "OEE"),
                    ((66, 107, 41),   "Availability"),
                    ((255, 215, 0),   "Performance"),
                    ((55, 81, 126),   "Quality"),
                    ((0, 206, 209),   f"OEE Target: {target_val}%")
                ]:
                    swatch = QtWidgets.QLabel()
                    swatch.setFixedSize(25, 4)
                    if label == f"OEE Target: {target_val}%":
                        swatch.setStyleSheet(
                            f"background-color: transparent; border-radius: 1px; border: 2px dashed rgb{color};")
                    else:
                        swatch.setStyleSheet(
                            f"background-color: rgb{color}; border-radius: 2px;")
                    text = QtWidgets.QLabel(label)
                    text.setStyleSheet("font-size: 11px;")
                    legend_layout.addWidget(swatch)
                    legend_layout.addWidget(text)
                legend_layout.addStretch()
                legend_widget.setLayout(legend_layout)
            
            def draw_OEE_metrics_column_chart(df, target_val, x_lbl, y_lbl, target_widget, legend_widget):
                remove_chart_items(target_widget)
                remove_chart_items(legend_widget, layout_type=QtWidgets.QHBoxLayout)
                layout = target_widget.layout()
                legend_layout = legend_widget.layout()
                legend_layout.setContentsMargins(4, 0, 4, 0)
                legend_layout.setSpacing(12)
                column_chart = pg.PlotWidget()
                column_chart.setBackground('w')
                column_chart.hideButtons()
                column_chart.setMouseEnabled(x=False, y=False)
                year = df["production_date"].dt.year.mode()[0]
                df["month"] = df["production_date"].dt.to_period("M")
                df_grouped = (df.groupby(["month"])
                                    .agg(
                                        {"planed_time": "sum",
                                        "OK_qty": "sum",
                                        "NG_qty": "sum",
                                        "Total_Loss": "sum",
                                         "Repair_Time": "sum",
                                         "Available_Time": "sum"}))
                full_months = pd.period_range(start=f"{year}-01", end=f"{year}-12", freq="M")
                df_grouped = df_grouped.reindex(full_months, fill_value=0)
                df_grouped.index.name = "month"
                df_grouped.reset_index(inplace=True) 
                x = np.arange(len(df_grouped)).astype(float)      
                month_labels = df_grouped["month"].dt.strftime("%b").tolist()           
                op_min  = (df_grouped["planed_time"] ).to_numpy(dtype=float)
                avail   = df_grouped["Available_Time"].to_numpy(dtype=float)
                qty_ok  = df_grouped["OK_qty"].to_numpy(dtype=float)
                qty_ng  = df_grouped["NG_qty"].to_numpy(dtype=float)
                qty_all = qty_ok + qty_ng
                df_grouped["Availability_percentage"] = safe_div(avail, op_min)
                df_grouped["Performance_percentage"]  = safe_div(cycle_time_value * qty_all, avail * 60)
                df_grouped["Quality_percentage"]      = safe_div(qty_ok, qty_all)
                df_grouped["OEE_percentage"] = (df_grouped["Availability_percentage"]
                                                * df_grouped["Performance_percentage"]
                                                * df_grouped["Quality_percentage"])
                oee_pct = (df_grouped["OEE_percentage"] * 100).to_numpy()
                
                def oee_brush(value, target):
                    if value >= target:
                        return pg.mkBrush(39, 174, 96)    
                    else:
                        return pg.mkBrush(231, 76, 60)
                bar_brushes = [oee_brush(v, target_val) for v in oee_pct]   
                bar = pg.BarGraphItem(
                    x=x,
                    height=df_grouped["OEE_percentage"] * 100,
                    width=0.6,
                    brushes=bar_brushes,
                    pen=None
                )
                column_chart.addItem(bar)
                column_chart.getAxis('bottom').setTicks([list(zip(x, month_labels))])
                y_max = max(df_grouped["OEE_percentage"] * 100) if max(df_grouped["OEE_percentage"] * 100) > 0 else 1 
                column_chart.setYRange(0, y_max*1.3, padding=0)
                limit = int(y_max*1.3)
                column_chart.getAxis('left').setTicks(
                    [[(i, str(i)) for i in range(0, limit + 1, limit//8 if limit >= 8 else 1)]])
                column_chart.getAxis('left').setStyle(tickLength=5)
                column_chart.getAxis('bottom').setStyle(tickLength=5)
                column_chart.showGrid(x=False, y=True, alpha=0.2)
                column_chart.setLabel('left', '<b>OEE (%)</b>', color='#36C9B1', size='10pt')
                def on_hover(pos):
                    if column_chart.sceneBoundingRect().contains(pos):
                        mouse_point = column_chart.getViewBox().mapSceneToView(pos)
                        x_mouse = mouse_point.x()
                        y_mouse = mouse_point.y()
                        closest_index = np.argmin(np.abs(x - x_mouse))
                        if closest_index < len(x):
                            tooltip_text = f"<b>Month:</b> {month_labels[closest_index]}-{year}<br>"
                            tooltip_text += f"<b>OEE:</b> {df_grouped['OEE_percentage'].iloc[closest_index]*100:.2f}%<br>"
                            tooltip_text += f"<b>Availability:</b> {df_grouped['Availability_percentage'].iloc[closest_index]*100:.2f}%<br>"
                            tooltip_text += f"<b>Performance:</b> {df_grouped['Performance_percentage'].iloc[closest_index]*100:.2f}%<br>"
                            tooltip_text += f"<b>Quality:</b> {df_grouped['Quality_percentage'].iloc[closest_index]*100:.2f}%"
                            QtWidgets.QToolTip.showText(
                                column_chart.mapToGlobal(pos.toPoint()),
                                tooltip_text
                            )
                column_chart.scene().sigMouseMoved.connect(on_hover)
                for i,value in enumerate(df_grouped["OEE_percentage"] * 100):
                    if value < target_val:
                        text_item = pg.TextItem(
                            html=f'<div style="color: #e74c3c; font-size: 10px; font-weight: bold;">{value:.1f}%</div>',
                            anchor=(0.5, 1.0)
                        )
                        column_chart.addItem(text_item)
                    else:
                        text_item = pg.TextItem(
                            html=f'<div style="color: #27ae60; font-size: 10px; font-weight: bold;">{value:.1f}%</div>',
                            anchor=(0.5, 1.0),
                        )
                        column_chart.addItem(text_item)
                    text_item.setPos(x[i], value)
                target_array = np.full(x.shape, target_val)
                target_line = pg.PlotDataItem(x, target_array, pen=pg.mkPen(color=(246, 24, 99), width=2, style=QtCore.Qt.DashLine), name='Target', antialias=True)
                column_chart.addItem(target_line)
                layout.addWidget(column_chart)

            self.oee_df['production_date'] = pd.to_datetime(
                self.oee_df['production_date'])
            self.oee_df.sort_values('production_date', inplace=True)
            mttr = self.downtime_df.loc[0,'total_repair_time'] / self.downtime_df.loc[0,"total_records"] if self.downtime_df.loc[0,"total_records"] > 0 else self.downtime_df.loc[0,"total_repair_time"]
            mtbf = (( working_time - float(self.downtime_df.loc[0,'total_loss_time'])) / float(self.downtime_df.loc[0,"total_records"])) if self.downtime_df.loc[0,"total_records"] > 0 else ( working_time - float(self.downtime_df.loc[0,'total_loss_time']))
            mttr_format = self.change_time_format(mttr, 'm')
            mtbf_format = self.change_time_format(mtbf, 'm' )
            self.ui.OEE_MTTR_value_lbl.setText(
                f"{mttr_format['h']} hrs {mttr_format['m']} mins" if mttr >= 60 else f"{mttr_format['m']} mins {mttr_format['s']} secs")
            self.ui.OEE_MTBF_value_lbl.setText(
                f"{mtbf_format['h']} hrs {mtbf_format['m']} mins" if mtbf >= 60 else f"{mtbf_format['m']} mins {mtbf_format['s']} secs")
            if pre_mttr is not None and pre_mttr != 0:
                self.ui.pre_MTTR_lbl.setText(make_comparison_html(float(mttr), float(pre_mttr) , "MTTR"))
            elif pre_mttr is not None and pre_mttr == 0:
                self.ui.pre_MTTR_lbl.setText(make_comparison_html(float(mttr), float(0) , "MTTR"))
            else:
                self.ui.pre_MTTR_lbl.setText("No previous data")
            self.ui.pre_MTBF_lbl.setText(
                make_comparison_html(float(mtbf), float(pre_mtbf), "MTBF") if pre_mtbf is not None and pre_mtbf != 0 else "No previous data")     
            def draw_output_nettime_chart(df, x_lbl, y_lbl, target_widget, legend_widget, ismonthly=True):
                remove_chart_items(target_widget)
                remove_chart_items(legend_widget, layout_type=QtWidgets.QHBoxLayout)
                layout = target_widget.layout()
                legend_layout = legend_widget.layout()
                legend_layout.setSpacing(5)
                if ismonthly:
                    x = df[x_lbl].apply(lambda d: d.timestamp()
                                        ).values.astype(float)
                    avail_time = df[y_lbl].values.astype(float)
                    output = df['OK_qty'].values.astype(float)
                    date_axis = DateAxisItem(orientation='bottom')
                    column_chart = pg.PlotWidget(axisItems={'bottom': date_axis})
                    column_chart.setBackground('w')
                    column_chart.hideButtons()
                    column_chart.setLabel('left', '<b>Net Time (mins)</b>', color='#8676DD', size='10pt')
                else:
                    year = df["production_date"].dt.year.mode()[0]
                    df["month"] = df["production_date"].dt.to_period("M")
                    df_grouped = (df.groupby(["month"])
                                    .agg(
                                        {"operation_hours": "sum",
                                        "OK_qty": "sum",
                                        "NG_qty": "sum",
                                        "Total_Loss": "sum",
                                         "Repair_Time": "sum",
                                         "Available_Time": "sum"}))
                    full_months = pd.period_range(start=f"{year}-01", end=f"{year}-12", freq="M")
                    df_grouped = df_grouped.reindex(full_months, fill_value=0)
                    df_grouped.index.name = "month"
                    df_grouped.reset_index(inplace=True)                           
                    x = np.arange(len(df_grouped)).astype(float)      
                    month_labels = df_grouped["month"].dt.strftime("%b").tolist() 
                    avail_time = df_grouped["Available_Time"].values.astype(float)
                    output     = df_grouped["OK_qty"].values.astype(float)                
                    column_chart = pg.PlotWidget()
                    column_chart.getAxis('bottom').setTicks([list(zip(x, month_labels))])
                    column_chart.setBackground('w')
                    column_chart.hideButtons()
                    column_chart.setLabel('left', '<b>Net Time (mins)</b>', color='#8676DD', size='10pt')

                bar = pg.BarGraphItem(
                    x=x,
                    height=avail_time,
                    width=0.6 if not ismonthly else 86400 * 0.7,
                    brush=pg.mkBrush(134, 118, 221),
                    pen=pg.mkPen(color=(80, 120, 200), width=1)
                )
                column_chart.addItem(bar)
                y_max = max(avail_time) if max(avail_time) > 0 else 1  
                column_chart.setYRange(0, y_max*1.3, padding=0)
                limit = int(y_max*1.3)
                column_chart.getAxis('left').setTicks(
                    [[(i, str(i)) for i in range(0, limit + 1, limit//8 if limit >= 8 else 1)]])
                column_chart.getAxis('left').setStyle(tickLength=5)
                column_chart.getAxis('bottom').setStyle(tickLength=5)
                column_chart.showGrid(x=False, y=True, alpha=0.2)
                vb2 = pg.ViewBox()
                column_chart.scene().addItem(vb2)
                vb2.setZValue(10)
                column_chart.getAxis('right').setStyle(tickLength=5)
                column_chart.showAxis('right')
                column_chart.getAxis('right').unlinkFromView = lambda: None  # bypass compiled_method disconnect bug in Nuitka
                column_chart.getAxis('right').linkToView(vb2)
                vb2.setXLink(column_chart.getViewBox())
                line2 = pg.PlotDataItem(
                    x, output,
                    pen=pg.mkPen(color=(247, 169, 168), width=2.5),
                    symbol='o',
                    symbolSize=5,
                    symbolBrush=pg.mkBrush(color =(253, 208, 209)),
                    symbolPen=pg.mkPen(color=(247, 169, 168), width=1),
                    antialias=True
                )
                vb2.addItem(line2)
                vb2.setYRange(0, max(output) * 1.3, padding=0)
                limit = int(max(output) * 1.3)
                column_chart.getAxis('right').setTicks(
                    [[(i, str(i)) for i in range(0, limit + 1, limit//8 if limit >= 8 else 1)]])
                column_chart.setLabel('right', '<b>Output (pcs)</b>', color='#F7A9A8', size='10pt')
                productivity = safe_div(output, avail_time)
                vb3 = pg.ViewBox()
                ax3 = pg.AxisItem('right')
                ax3.setLabel('<b>Productivity (pcs/min)</b>', color='#27AE60', size='10pt')
                p_item = column_chart.plotItem
                p_item.layout.addItem(ax3, 2, 3)
                p_item.scene().addItem(vb3)
                vb3.setZValue(11)
                ax3.unlinkFromView = lambda: None  # bypass compiled_method disconnect bug in Nuitka
                ax3.linkToView(vb3)
                ax3.setStyle(tickLength=5)
                vb3.setXLink(p_item.vb)
                line3 = pg.PlotDataItem(
                    x, productivity,
                    pen=pg.mkPen(color=(39, 174, 96), width=2,
                                 style=QtCore.Qt.DashLine),
                    symbol='s',
                    symbolSize=5,
                    symbolBrush=pg.mkBrush(color =(137, 255, 196)),
                    symbolPen=pg.mkPen(color=(39, 174, 96), width=1),
                    antialias=True
                )
                vb3.addItem(line3)
                vb3.setYRange(0, max(productivity) * 1.3 if max(productivity) > 0 else 1, padding=0)
                limit = int(max(productivity) * 1.3 if max(productivity) > 0 else 1)
                ax3.setTicks(
                    [[(i, str(i)) for i in range(0, limit + 1, limit//8 if limit >= 8 else 1)]])

                def update_views():
                    vb2.setGeometry(column_chart.getViewBox().sceneBoundingRect())
                    vb2.linkedViewChanged(column_chart.getViewBox(), vb2.XAxis)
                    vb3.setGeometry(column_chart.getViewBox().sceneBoundingRect())
                    vb3.linkedViewChanged(column_chart.getViewBox(), vb3.XAxis)

                column_chart.getViewBox().sigResized.connect(update_views)
                column_chart.getViewBox().wheelEvent = lambda e: None
                update_views()

                column_chart.showGrid(x=False, y=True, alpha=0.15)
                layout.addWidget(column_chart)
                update_views()
                if legend_layout.count() > 0:
                    return
                horizontal_spacer = QtWidgets.QSpacerItem(
                    20, 10, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
                legend_layout.addItem(horizontal_spacer)
                legend_items = [
                    ("Available Time", (134, 118, 221,255), "bar"),
                    ("OK Quantity (pcs)", (247, 169, 168,255), "line"),
                    ("Productivity (pcs/min)", (39, 174, 96, 200), "dash"),
                ]
                for label, color, style in legend_items:
                    legend_item = QtWidgets.QWidget()
                    legend_item.setFixedSize(140, 30)
                    legend_layout_item = QtWidgets.QHBoxLayout(legend_item)
                    legend_layout_item.setContentsMargins(0, 0, 0, 0)
                    legend_layout_item.setSpacing(5)
                    color_box = QtWidgets.QLabel()
                    if style == "bar":
                        color_box.setFixedSize(12, 12)
                        color_box.setStyleSheet(
                            f"background-color: rgba({color[0]}, {color[1]}, {color[2]}, {color[3]}); border: 1px solid rgba({color[0]}, {color[1]}, {color[2]}, 255); border-radius: 0px;")
                    elif style == "dash":
                        color_box.setFixedSize(25, 4)
                        color_box.setStyleSheet(
                            f"background-color: transparent; border-top: 2px dashed rgba({color[0]}, {color[1]}, {color[2]}, {color[3]});")
                    else:
                        color_box.setFixedSize(25, 4)
                        color_box.setStyleSheet(
                            f"background-color: rgba({color[0]}, {color[1]}, {color[2]}, {color[3]}); border-radius: 2px;")
                    legend_layout_item.addWidget(color_box)
                    legend_label = QtWidgets.QLabel()
                    legend_label.setText(label)
                    legend_label.setStyleSheet("color: gray; font-size: 8pt;")
                    legend_layout_item.addWidget(legend_label)
                    legend_layout.addWidget(legend_item)
                horizontal_spacer2 = QtWidgets.QSpacerItem(
                    20, 10, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
                legend_layout.addItem(horizontal_spacer2)
                def on_hover(pos):
                    if column_chart.sceneBoundingRect().contains(pos):
                        mouse_point = column_chart.getViewBox().mapSceneToView(pos)
                        x_mouse = mouse_point.x()
                        y_mouse = mouse_point.y()
                        closest_index = np.argmin(np.abs(x - x_mouse))
                        if closest_index < len(x):
                            tooltip_text = f"<b>Date:</b> {dt.datetime.fromtimestamp(x[closest_index]).strftime('%Y-%m-%d') if ismonthly else f'{month_labels[closest_index]}-{year}'}<br>"
                            tooltip_text += f"<b>Net Time:</b> {avail_time[closest_index]} mins<br>"
                            tooltip_text += f"<b>Output:</b> {change_value_show(output[closest_index])}pcs<br>"
                            tooltip_text += f"<b>Productivity:</b> {productivity[closest_index]:.2f} pcs/min<br>"
                            QtWidgets.QToolTip.showText(
                                QtGui.QCursor.pos(), tooltip_text, column_chart)
                column_chart.scene().sigMouseMoved.connect(on_hover)
                main_vb = column_chart.getViewBox()
                main_vb.mouseDragEvent = lambda ev, axis=None: ev.ignore()
                main_vb.wheelEvent = lambda ev: ev.ignore()

                vb2.mouseDragEvent = lambda ev, axis=None: ev.ignore()
                vb2.wheelEvent = lambda ev: ev.ignore()

                vb3.mouseDragEvent = lambda ev, axis=None: ev.ignore()
                vb3.wheelEvent = lambda ev: ev.ignore()
            @QtCore.pyqtSlot()
            def on_switch_toggled(checked):
                if checked:
                    draw_output_nettime_chart(
                        self.oee_df, 'production_date', 'Available_Time', self.ui.OEE_information_chart, self.ui.OEE_infor_chart_legend, ismonthly=ismonthly)
                else:
                    if ismonthly:
                        draw_OEE_metrics_line_chart(self.oee_df, float(oee_targets[0][3]), 'production_date', 'OEE_percentage',
                                            self.ui.OEE_information_chart, self.ui.OEE_infor_chart_legend)
                    else:
                        draw_OEE_metrics_column_chart(self.oee_df, float(oee_targets[0][3]), 'production_date', 'OEE_percentage',
                                            self.ui.OEE_information_chart, self.ui.OEE_infor_chart_legend)
            on_switch_toggled(self.OEE_chart_toogle_btn.isChecked())
            self.safe_connect(self.OEE_chart_toogle_btn.toggled, on_switch_toggled)
            def Generate_key_insights_en(oee, oee_targets, oee_previous, mttr, mttr_targets, mtbf, mtbf_targets, a , a_targets, a_previous, p, p_targets, p_previous, q, q_targets, q_previous):
                insights = []          
                final_conclusion = ""
                insights.append(f"<p style='margin: 0px; margin-top: 8px;'>• <b>Overview:</b></p>")
                style_sub_overview = "style='margin: 0px; margin-left: 25px;'"   
                style_sub_ref = "style='margin: 0px; margin-left: 25px; text-align: justify;'"
                OEE_percentage = round(oee * 100,1)
                if oee*100 < oee_targets:
                    insights.append(f"<p {style_sub_overview}>• <b>OEE:</b> <span style='color:red; font-weight: bold;'> Missed target {round(oee_targets,1)}% | Actual {OEE_percentage}%</span></p>")
                    final_conclusion = "<span style='font-weight: bold; color: #ef4444;'>Need for immediate action.</span>"
                else:
                    insights.append(f"<p {style_sub_overview}>• <b>OEE:</b> <span style='color: #10b981; font-weight: bold;'> Reached target {oee_targets}% | Actual {OEE_percentage}%</span></p>")
                    final_conclusion = "No need for immediate action."
                if mttr > mttr_targets:
                    insights.append(f"<p {style_sub_overview}>• <b>MTTR:</b> <span style='color:red; font-weight: bold;'> Missed target {float(mttr_targets)} mins | Actual {round(mttr,1)} mins</span></p>")
                else:
                    insights.append(f"<p {style_sub_overview}>• <b>MTTR:</b> <span style='color: #10b981; font-weight: bold;'> Reached target {float(mttr_targets)} mins | Actual {round(mttr,1)} mins</span></p>")
                if mtbf < mtbf_targets:
                    insights.append(f"<p {style_sub_overview}>• <b>MTBF:</b> <span style='color:red; font-weight: bold;'> Missed target {float(mtbf_targets)} mins | Actual {round(mtbf,1)} mins</span></p>")
                else:
                    insights.append(f"<p {style_sub_overview}>• <b>MTBF:</b> <span style='color: #10b981; font-weight: bold;'> Reached target {float(mtbf_targets)} mins | Actual {round(mtbf,1)} mins</span></p>")
                insights.append(f"<p style='margin: 0px; margin-top: 14px;'>• <b>Reference:</b></p>")
                if a < a_targets:
                    insights.append(f"<p {style_sub_overview}>• <b>A:</b> <span style='color:red; font-weight: bold;'> Availability has no good result {round(a,1)}%</span></p>")
                else:
                    if a >= a_previous:
                        insights.append(f"<p {style_sub_overview}>• <b>A:</b> <span style='color: #10b981; font-weight: bold;'> Availability has good result {round(a,1)}% and has increased from previous</span></p>")
                    else:
                        insights.append(f"<p {style_sub_overview}>• <b>A:</b> <span style='color:#10b981; font-weight: bold;'> Availability has good result {round(a,1)}%</span><span style='color:#CF763E; font-weight: bold;'> but has decreased from previous</span></p>")
                if p < p_targets:
                    insights.append(f"<p {style_sub_overview}>• <b>P:</b> <span style='color:red; font-weight: bold;'> Performance has no good result {round(p,1)}%</span></p>")
                else:
                    if p >= p_previous:
                        insights.append(f"<p {style_sub_overview}>• <b>P:</b> <span style='color: #10b981; font-weight: bold;'> Performance has good result {round(p,1)}% and has increased from previous</span></p>")
                    else:
                        insights.append(f"<p {style_sub_overview}>• <b>P:</b> <span style='color:#10b981; font-weight: bold;'> Performance has good result {round(p,1)}%</span><span style='color:#CF763E; font-weight: bold;'> but has decreased from previous</span></p>")
                if q < q_targets:
                    insights.append(f"<p {style_sub_overview}>• <b>Q:</b> <span style='color:red; font-weight: bold;'> Quality has no good result {round(q,1)}%</span></p>")
                else:
                    if q >= q_previous:
                        insights.append(f"<p {style_sub_overview}>• <b>Q:</b> <span style='color: #10b981; font-weight: bold;'> Quality has good result {round(q,1)}% and has increased from previous</span></p>")
                    else:
                        insights.append(f"<p {style_sub_overview}>• <b>Q:</b> <span style='color:#10b981; font-weight: bold;'> Quality has good result {round(q,1)}%</span><span style='color:#CF763E; font-weight: bold;'> but has decreased from previous</span></p>")
                insights.append(f"<p style='margin: 0px; margin-top: 14px;'>• <b>Final Recommendation:</b> {final_conclusion}</p>")
                return "".join(insights)
            self.ui.OEE_keyinsights_lbl.clear()
            self.ui.OEE_keyinsights_lbl.setText(Generate_key_insights_en(oee = total_OEE, oee_targets = float(oee_targets[0][3]), oee_previous = float(pre_total_OEE*100) if previous_oee_data else 0,
                                                                          mttr = mttr, mttr_targets = float(mttr_target) if mttr_target is not None else 0, 
                                                                          mtbf = mtbf, mtbf_targets = float(mtbf_target) if mttr_target is not None else 0,
                                                                          a = total_A_percentage*100, a_targets = float(oee_targets[0][0]), a_previous = float(pre_total_A_percentage*100) if previous_oee_data else 0,
                                                                          p = total_P_percentage*100, p_targets = float(oee_targets[0][1]), p_previous = float(pre_total_P_percentage*100) if previous_oee_data else 0,
                                                                          q = total_Q_percentage*100, q_targets = float(oee_targets[0][2]), q_previous = float(pre_total_Q_percentage*100) if previous_oee_data else 0))
            self.spinner.stop()
            self.safe_connect(self.ui.OEE_comment_text_edit.textChanged, lambda: self.action_editing_finished(widget=self.ui.OEE_comment_text_edit,btn = self.ui.OEE_action_save_btn, data = {"has_OEE_comment": self.has_OEE_comment if hasattr(self, 'has_OEE_comment') else False,
                                                                                                                                                                                                "OEE_comment_id": self.OEE_comment_id,
                                                                                                                                                                                                "area_name": self.oee_df.loc[0, 'area_name'],
                                                                                                                                                                                                "select_category" : "line_name",
                                                                                                                                                                                                "line_name": self.oee_df.loc[0, 'line_name'],
                                                                                                                                                                                                "select_category_sql": "(SELECT line_id FROM production_lines WHERE line_name = :line_name)",
                                                                                                                                                                                                "for_month": self.ui.OEE_period_edit.date().month(),
                                                                                                                                                                                                "for_year": self.ui.OEE_period_edit.date().year(),
                                                                                                                                                                                                "action_for": "OEE"}))
        try:
            if hasattr(self, 'OEE_dashboard_worker') and self.OEE_dashboard_worker.isRunning():
                self.spinner.stop()
                return
            self.OEE_dashboard_worker = WorkerThread(fetch_data)
            self.OEE_dashboard_worker.finished.connect(lambda res: on_data_fetched(res))
            self.OEE_dashboard_worker.start()
        except Exception as e:
            self.spinner.stop()
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to refresh OEE page: {e}")
    
    @QtCore.pyqtSlot()
    def OEE_load_filter(self):
        try:
            line = self.ui.OEE_line_cbb.currentText()
            process = self.ui.OEE_process_cbb.currentText()
            model = self.ui.OEE_model_cbb.currentText()
            area = self.ui.OEE_area_cbb.currentText()
            if self.ui.OEE_month_radio.isChecked():
                month = self.ui.OEE_period_edit.date().month()
                year = self.ui.OEE_period_edit.date().year()
            if self.ui.OEE_year_radio.isChecked():
                month = 0
                year = int(self.ui.OEE_year_edit.currentText())
                
            self.refesh_OEE_page(area, model, month, year, line, process)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load OEE filters: {e}")
    
    @QtCore.pyqtSlot()
    def OEE_filter_changing(self,change_object,area_cbb = None, model_cbb = None, line_cbb = None , process_cbb = None , period_cbb = None, clicked = True):
        try:
            widgets = [
                area_cbb,
                line_cbb,
                model_cbb,
                process_cbb,
                period_cbb,
                self.ui.OEE_year_edit,
                self.ui.OEE_month_radio,
                self.ui.OEE_dashboard_view_stacked_widget,
            ]
            blockers = [QtCore.QSignalBlocker(w) for w in widgets if w is not None]
            if self.ui.OEE_stacked_widget.currentWidget() == self.ui.OEE_Dashboard_page:
                if change_object == "overview_mode":
                    self.ui.frame_180.setHidden(True)
                    self.ui.OEE_load_filter_btn.setHidden(True)
                    self.ui.OEE_dashboard_view_stacked_widget.setCurrentWidget(self.ui.OEE_over_view_page)
                    area_name = area_cbb.currentText()
                    if self.ui.OEE_month_radio.isChecked():
                        month = period_cbb.date().month()
                        year = period_cbb.date().year()
                    else:
                        month = 0
                        year = period_cbb.date().year()
                    self.OEE_overview_dashboard(area_name, month, year)
                    return
                elif change_object == "detail_mode":
                    self.ui.frame_180.setHidden(False)
                    self.ui.OEE_load_filter_btn.setHidden(False)
                    self.ui.OEE_dashboard_view_stacked_widget.setCurrentWidget(self.ui.OEE_detail_view_page)
                    if clicked != True:
                        return
                if self.ui.OEE_dashboard_view_stacked_widget.currentWidget() == self.ui.OEE_detail_view_page:
                    if self.ui.OEE_month_radio.isChecked():
                        self.ui.OEE_period_edit.setDisplayFormat("MMM-yyyy")
                        self.ui.OEE_year_edit.hide()
                        month = period_cbb.date().month()
                        year = period_cbb.date().year()
                        self.ui.OEE_period_edit.setHidden(False)
                        filter_script_area = f" AND MONTH(po.production_date) = {month} AND YEAR(po.production_date) = {year}"
                    else:
                        self.ui.OEE_year_edit.setHidden(False)
                        year = self.ui.OEE_year_edit.currentText()
                        self.ui.OEE_period_edit.hide()
                        filter_script_area = f" AND YEAR(po.production_date) = {year}"
                else:
                    if self.ui.OEE_month_radio.isChecked():
                        self.ui.OEE_year_edit.hide()
                        month = period_cbb.date().month()
                        year = period_cbb.date().year()
                        self.ui.OEE_period_edit.setHidden(False)
                        self.OEE_overview_dashboard(area_cbb.currentText(), month, year)
                    else:
                        self.ui.OEE_year_edit.setHidden(False)
                        year = self.ui.OEE_year_edit.currentText()
                        self.ui.OEE_period_edit.hide()
                        self.OEE_overview_dashboard(area_cbb.currentText(), 0, year)
                    return
            else:
                month = period_cbb.date().month()
                year = period_cbb.date().year()
                filter_script_area = f" AND MONTH(po.production_date) = {month} AND YEAR(po.production_date) = {year}"
            if change_object == "area" or change_object == "period":
                line_cbb.clear()
                model_cbb.clear()
                process_cbb.clear()
                area_name = area_cbb.currentText()
                lines = self.database_process.query(sql=f'''SELECT DISTINCT pl.line_name
                                                                FROM `production_lines` as pl
                                                                JOIN `production_output` as po ON pl.line_id = po.line_id
                                                                JOIN `product_models_oee` as pmo ON po.model_name = pmo.model_name
                                                                JOIN `downtime_areas` as da ON pmo.department_id = da.department_id
                                                                WHERE da.downtime_area_name = :area_name AND po.OK_qty > 100  {filter_script_area}
                                                                ORDER BY pl.line_name ASC;''', params={"area_name": area_name})
                OEE_model = self.database_process.query(sql=f'''SELECT DISTINCT pmo.model_name 
                                                                FROM `product_models_oee` as pmo
                                                                JOIN `production_output` as po ON pmo.model_name = po.model_name
                                                                JOIN `production_lines` as pl ON po.line_id = pl.line_id
                                                                WHERE pl.line_name = :line_name  {filter_script_area};''', params={"line_name": lines[0][0]})
                process = self.database_process.query(sql=f'''SELECT DISTINCT mor.process 
                                                                FROM `machine_OEE_register` as mor
                                                                JOIN `machines` as m ON mor.machine_id = m.machine_id
                                                                JOIN `production_lines` as pl ON m.line_id = pl.line_id
                                                                JOIN `product_models_oee` as pmo ON mor.model_id = pmo.model_id
                                                                WHERE pl.line_name = :line_name AND pmo.model_name = :model_name;''', params={"line_name": lines[0][0], "model_name": OEE_model[0][0]})
                line_cbb.clear()
                line_cbb.addItems([line[0] for line in lines])
                model_cbb.clear()
                model_cbb.addItems([model[0] for model in OEE_model])
                process_cbb.clear()
                process_cbb.addItems([p[0] for p in process])
            elif change_object == "line":
                model_cbb.clear()
                process_cbb.clear()
                line_name = line_cbb.currentText()
                OEE_model = self.database_process.query(sql=f'''SELECT DISTINCT pmo.model_name 
                                                                                FROM `product_models_oee` as pmo
                                                                                JOIN `production_output` as po ON pmo.model_name = po.model_name
                                                                                JOIN `production_lines` as pl ON po.line_id = pl.line_id
                                                                                WHERE pl.line_name = :line_name  {filter_script_area};''', params={"line_name": line_name})
                if not OEE_model:
                    return
                process = self.database_process.query(sql='''SELECT DISTINCT mor.process 
                                                                FROM `machine_OEE_register` as mor
                                                                JOIN `machines` as m ON mor.machine_id = m.machine_id
                                                                JOIN `production_lines` as pl ON m.line_id = pl.line_id
                                                                JOIN `product_models_oee` as pmo ON mor.model_id = pmo.model_id
                                                                WHERE pl.line_name = :line_name AND pmo.model_name = :model_name;''', params={"line_name": line_name, "model_name": OEE_model[0][0]})
                model_cbb.addItems([model[0] for model in OEE_model])
                process_cbb.addItems([p[0] for p in process])
            elif change_object == "model":
                line_name = line_cbb.currentText()
                model_name = model_cbb.currentText()
                process = self.database_process.query(sql='''SELECT DISTINCT mor.process 
                                                                FROM `machine_OEE_register` as mor
                                                                JOIN `machines` as m ON mor.machine_id = m.machine_id
                                                                JOIN `production_lines` as pl ON m.line_id = pl.line_id
                                                                JOIN `product_models_oee` as pmo ON mor.model_id = pmo.model_id
                                                                WHERE pl.line_name = :line_name AND pmo.model_name = :model_name;''', params={"line_name": line_name, "model_name": model_name})
                process_cbb.clear()
                if process:
                    process_cbb.addItems([p[0] for p in process])
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to apply OEE filters: {e}")
    
    def extract_content(self, text_edit):
        result = []
        doc = text_edit.document()
        block = doc.begin()
        while block.isValid():
            it = block.begin()
            while not it.atEnd():
                fragment = it.fragment()
                if fragment.isValid():
                    fmt = fragment.charFormat()
                    text = fragment.text()
                    if fmt.isAnchor():                     
                        result.append({
                            'type': 'link',
                            'text': text,                
                            'href': fmt.anchorHref()      
                        })
                    else:                                
                        result.append({'type': 'text', 'text': text})
                it += 1
            block = block.next()
        return result
    
    @QtCore.pyqtSlot()
    def action_editing_finished(self,widget, btn, data = {}):
        if not btn.isHidden():
            return
        btn.setHidden(False)
        btn.setEnabled(True)
        self.safe_connect(btn.clicked, lambda: self.save_actions(widget, btn, data))
    
    @QtCore.pyqtSlot()
    def save_actions(self,widget, btn, data = {}):
        def sync_action_state(has_comment, comment_id):
            if data.get("action_for") == "OEE":
                self.has_OEE_comment = has_comment
                self.OEE_comment_id = comment_id
            elif data.get("action_for") == "DT":
                self.has_DT_comment = has_comment
                self.DT_comment_id = comment_id
        try:
            comment = widget.toPlainText()
            content = self.extract_content(widget)
            link_list = []
            for key, value in enumerate(content):
                if value['type'] == 'link':
                    link_list.append({
                        "file_name": value['text'],
                        "file_path": value['href']
                    }
                )
                    comment = comment.replace(value['text'], '')
            link_list = json.dumps(link_list, ensure_ascii=False)
            if content:
                if data.get("has_OEE_comment"):
                    self.database_process.query(
                        sql='''UPDATE downtime_actions 
                               SET action_content = :action_content, action_report_link = :action_report_link
                               WHERE action_id = :action_id''',
                        params={
                            "action_content": comment,
                            "action_report_link": link_list,
                            "action_id": data.get("OEE_comment_id")
                        }
                    )
                else:
                    select_category = data.get("select_category")
                    if select_category in ["line_name", "machine_code", "error_code"]:
                        result = self.database_process.query(f''' INSERT INTO downtime_actions (action_content, action_report_link, line_id, downtime_area_id, for_month, for_year, action_for)
                                                    VALUES (:action_content, :action_report_link, {data.get("select_category_sql")},
                                                            (SELECT downtime_area_id FROM downtime_areas WHERE downtime_area_name = :area_name),
                                                                :for_month, :for_year, "{data.get("action_for")}");''',
                                                    params={
                                                        "action_content": comment,
                                                        "action_report_link": link_list,
                                                        select_category : data.get(select_category),
                                                        "area_name": data.get("area_name"),
                                                        "for_month": data.get("for_month"),
                                                        "for_year": data.get("for_year")
                                                    }
                        ) 
                    else:
                        result = self.database_process.query(f''' INSERT INTO downtime_actions (action_content, action_report_link, downtime_area_id, for_month, for_year, action_for)
                                                    VALUES (:action_content, :action_report_link,
                                                            (SELECT downtime_area_id FROM downtime_areas WHERE downtime_area_name = :area_name),
                                                                :for_month, :for_year, "{data.get("action_for")}");''',
                                                    params={
                                                        "action_content": comment,
                                                        "action_report_link": link_list,
                                                        "area_name": data.get("area_name"),
                                                        "for_month": data.get("for_month"),
                                                        "for_year": data.get("for_year")
                                                    }
                        )
                    sync_action_state(
                        has_comment=True,
                        comment_id=result.lastrowid
                    )
                        
            else:
                if data.get("has_OEE_comment"):
                    self.database_process.query(
                        sql='''DELETE FROM downtime_actions WHERE action_id = :action_id''',
                        params={
                            "action_id": data.get("OEE_comment_id")
                        }
                    )
                    sync_action_state(
                        has_comment=False,
                        comment_id=None
                    )
            btn.setHidden(True)
            btn.setEnabled(False)
            
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self,
                "Error",
                f"An error occurred while saving the OEE comment: {str(e)}"
            )
    
    @QtCore.pyqtSlot()
    def action_text_drop_event(self, widget, event):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0].toLocalFile()
            extracted_text = url.split("/")[-1]
            file_url = QtCore.QUrl.fromLocalFile(url).toString()
            cursor = widget.textCursor()
            cursor.movePosition(QtGui.QTextCursor.End)
            widget.setTextCursor(cursor)
            widget.append(
                f'<a href="{file_url}" style="color:#007acc; font-size:10px; text-decoration:underline;">'
                f'{extracted_text}</a> '
            )
            event.acceptProposedAction()

    @QtCore.pyqtSlot()
    def action_text_mouse_press_event(self, widget, event):
        anchor = widget.anchorAt(event.pos())
        if anchor:
            url = QtCore.QUrl(anchor)
            path = url.toLocalFile() if url.isLocalFile() else anchor
            path = os.path.normpath(path)  
            if os.path.exists(path):
                try:
                    os.startfile(path)
                except OSError as e:
                    QtWidgets.QMessageBox.warning(
                        widget, "Lỗi", f"Không mở được file:\n{e}")
            else:
                QtWidgets.QMessageBox.warning(
                    widget, "Không tìm thấy file",
                    f"File không tồn tại hoặc không truy cập được:\n{path}")
            return
        QtWidgets.QTextEdit.mousePressEvent(widget, event)
        
    @QtCore.pyqtSlot()
    def action_text_mouse_move_event(self, widget, event):
        if widget.anchorAt(event.pos()):
            widget.viewport().setCursor(QtCore.Qt.PointingHandCursor)
        else:
            widget.viewport().setCursor(QtCore.Qt.IBeamCursor)
        QtWidgets.QTextEdit.mouseMoveEvent(widget, event)

    @QtCore.pyqtSlot()
    def OEE_detail_page(self):
        self.ui.OEE_stacked_widget.setCurrentWidget(self.ui.OEE_Detail_page)
        self.style_button_with_shadow((self.ui.OEE_data_btn,self.ui.OEE_dashboard_btn))
        if self.ui.OEE_Data_Area_cbb.count() > 0:
            return
        try:
            if self.ui.OEE_Data_Area_cbb.count() > 0:
                return
            areas = self.database_process.query(
                sql="SELECT downtime_area_name FROM downtime_areas;")
            self.ui.OEE_Data_Area_cbb.addItems([area[0] for area in areas])
            self.ui.OEE_Data_period_edit.setDate(QtCore.QDate.currentDate().addMonths(-1))
            self.OEE_filter_changing(change_object="period", area_cbb=self.ui.OEE_Data_Area_cbb, model_cbb=self.ui.OEE_Data_Model_cbb, line_cbb=self.ui.OEE_Data_Line_cbb, process_cbb=self.ui.OEE_Data_Process_cbb, period_cbb=self.ui.OEE_Data_period_edit)
            headers = ["Area", "Line", "Model", "Process", "Production\nDate", "Working Shift\n(hours)", "Break Time\n(mins)", "Setup Time\n(mins)", "Planed Time\n(mins)" , "Total Loss\n(mins)", "Available\nTime\n(mins)", "FGs Output\n(pcs)", "Defect\n(pcs)", "Availability\n(%)", "Performance\n(%)", "Quality\n(%)", "OEE\n(%)"]
            self.OEE_Data_model = QtGui.QStandardItemModel(0, len(headers))
            self.OEE_Data_model.setHorizontalHeaderLabels(headers)
            self.ui.OEE_Data_table.setModel(self.OEE_Data_model)
            self.ui.OEE_Data_table.verticalHeader().setVisible(False)
            self.ui.OEE_Data_table.setColumnHidden(0, True)
            self.ui.OEE_Data_table.setColumnHidden(1, True)
            self.ui.OEE_Data_table.setColumnHidden(2, True)
            self.ui.OEE_Data_table.setColumnHidden(3, True)
            self.ui.OEE_Data_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)  
            self.ui.OEE_Data_table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
            vetical_header = ["Cycle Time" , "Total Availability (%)", "Total Performance (%)", "Total Quality (%)", "Total OEE (%)"]
            self.OEE_Data_summary_model = QtGui.QStandardItemModel(len(vetical_header), 2)
            for i in range(len(vetical_header)):
                item = QtGui.QStandardItem(vetical_header[i])
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.OEE_Data_summary_model.setItem(i, 0, item)
            self.ui.OEE_Data_sumary_table.setModel(self.OEE_Data_summary_model)
            self.ui.OEE_Data_sumary_table.setColumnWidth(0, 150)
            self.ui.OEE_Data_sumary_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Stretch)
            self.safe_connect(self.ui.OEE_Data_table.customContextMenuRequested, lambda pos: self.table_context_menu(pos = pos,table = self.ui.OEE_Data_table, actions= ["edit", "separator", "delete"],
                                                                                                                     functions_dict={"edit": self.edit_OEE_data, "delete": self.delete_OEE_data}))
            self.safe_connect(self.ui.OEE_Data_calendar_widget.currentPageChanged, lambda year,
                                month: self.update_date_from_calendar(year, month, self.ui.OEE_Data_period_edit))
            self.safe_connect(self.ui.OEE_Data_Load_btn.clicked, self.OEE_detail_data)
            self.safe_connect(self.ui.OEE_Data_Other_btn.clicked, lambda: self.Load_OEE_Data_Other())
            self.safe_connect(self.ui.OEE_Data_period_edit.dateChanged, lambda: self.OEE_filter_changing(change_object="period", area_cbb=self.ui.OEE_Data_Area_cbb, model_cbb=self.ui.OEE_Data_Model_cbb, line_cbb=self.ui.OEE_Data_Line_cbb, process_cbb=self.ui.OEE_Data_Process_cbb , period_cbb=self.ui.OEE_Data_period_edit))
            self.safe_connect(self.ui.OEE_Data_Area_cbb.currentIndexChanged, lambda: self.OEE_filter_changing(change_object="area", area_cbb=self.ui.OEE_Data_Area_cbb, model_cbb=self.ui.OEE_Data_Model_cbb, line_cbb=self.ui.OEE_Data_Line_cbb, process_cbb=self.ui.OEE_Data_Process_cbb, period_cbb=self.ui.OEE_Data_period_edit))
            self.safe_connect(self.ui.OEE_Data_Model_cbb.currentIndexChanged, lambda: self.OEE_filter_changing(change_object="model", area_cbb=self.ui.OEE_Data_Area_cbb, model_cbb=self.ui.OEE_Data_Model_cbb, line_cbb=self.ui.OEE_Data_Line_cbb, process_cbb=self.ui.OEE_Data_Process_cbb, period_cbb=self.ui.OEE_Data_period_edit))
            self.safe_connect(self.ui.OEE_Data_Line_cbb.currentIndexChanged, lambda: self.OEE_filter_changing(change_object="line", area_cbb=self.ui.OEE_Data_Area_cbb, model_cbb=self.ui.OEE_Data_Model_cbb, line_cbb=self.ui.OEE_Data_Line_cbb, process_cbb=self.ui.OEE_Data_Process_cbb, period_cbb=self.ui.OEE_Data_period_edit))
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load OEE detail page: {e}")
    
    @QtCore.pyqtSlot()
    def OEE_detail_data(self):
        try:
            area_name = self.ui.OEE_Data_Area_cbb.currentText()
            model_name = self.ui.OEE_Data_Model_cbb.currentText()
            line = self.ui.OEE_Data_Line_cbb.currentText()
            process = self.ui.OEE_Data_Process_cbb.currentText()
            month = self.ui.OEE_Data_period_edit.date().month()
            year = self.ui.OEE_Data_period_edit.date().year()
            oee_data = self.database_process.query(sql=f'''SELECT * FROM `oee_report`
                                                    WHERE area_name = :area_name 
                                                    AND model_name = :model_name 
                                                    AND MONTH(production_date) = :month
                                                    AND YEAR(production_date) = :year
                                                    AND line_name = :line AND process = :process;''', params={"area_name": area_name, "model_name": model_name, "month": month, "year": year, "line": line, "process": process})
            oee_targets = self.database_process.query(sql=f'''SELECT availability_target, performance_target, quality_target, oee_target 
                                                                FROM `oee_targets` as ot
                                                                JOIN `product_models_oee` as pmo ON ot.model_id = pmo.model_id
                                                                JOIN `production_lines` as pl ON ot.line_id = pl.line_id
                                                                WHERE pmo.model_name = :model_name
                                                                AND pl.line_name = :line AND ot.process = :process
                                                                AND ot.date_created <= :date
                                                                ORDER BY ot.date_created DESC
                                                                LIMIT 1;''', params={ "model_name": model_name, "line": line, "process": process, "date": f"{year}-{month:02d}-{calendar.monthrange(year, month)[1]}"})
            cycle_time = self.database_process.query(sql='''SELECT mct.create_at, mct.cycle_time_seconds , m.machine_code
                                                                FROM machine_cycle_times AS mct
                                                                JOIN product_models_oee AS pmo ON mct.model_id = pmo.model_id
                                                                JOIN machine_oee_register AS mor ON pmo.model_id = mor.model_id
                                                                JOIN production_lines AS pl ON mor.line_id = pl.line_id
                                                                JOIN machines AS m ON mor.machine_id = m.machine_id
                                                                WHERE mor.process = :process AND pmo.model_name = :model_name AND pl.line_name = :line_name AND mct.create_at <= :date
                                                                ORDER BY mct.create_at DESC LIMIT 1;
                                                            ''', params={"process": process, "model_name": model_name, "line_name": line, "date": f"{year}-{month:02d}-{calendar.monthrange(year, month)[1]}"})
            self.OEE_data_frame = pd.DataFrame(oee_data, columns=["area_name", "line_name", "model_name", "process", "production_date", "working_shift_hours","break_time", "setup_time", "planed_time", "fgs_output_pcs", "defect_pcs", "total_loss_mins", "repair_time_mins", "available_time_mins", "availability_percentage", "performance_percentage", "quality_percentage", "oee_percentage"])
            self.OEE_data_frame.drop(columns=["repair_time_mins"], inplace=True)
            col = self.OEE_data_frame.pop("total_loss_mins")
            self.OEE_data_frame.insert(9, "total_loss_mins", col)
            col = self.OEE_data_frame.pop("available_time_mins")
            self.OEE_data_frame.insert(10, "available_time_mins", col)
            self.OEE_data_frame["availability_percentage"] = self.OEE_data_frame["availability_percentage"].apply(lambda x: round(x*100, 2))
            self.OEE_data_frame["performance_percentage"] = self.OEE_data_frame["performance_percentage"].apply(lambda x: round(x*100, 2))
            self.OEE_data_frame["quality_percentage"] = self.OEE_data_frame["quality_percentage"].apply(lambda x: round(x*100, 2))
            self.OEE_data_frame["oee_percentage"] = self.OEE_data_frame["oee_percentage"].apply(lambda x: round(x*100, 2))
            self.add_data_to_model(self.OEE_data_frame.values.tolist(), self.ui.OEE_Data_table, self.OEE_Data_model, tooltip_Enable=False, target_for_item = {16: oee_targets[0][3]}, highlight_color= (239, 141, 205,100))
            total_OK_qty = float(self.OEE_data_frame['fgs_output_pcs'].sum())
            total_NG_qty = float(self.OEE_data_frame['defect_pcs'].sum())
            total_operation_hours = float(self.OEE_data_frame['planed_time'].sum())
            self.cycle_time_value = float(cycle_time[0][1]) if cycle_time else 0
            self.OEE_detail_machine_code = cycle_time[0][2] if cycle_time else None
            total_runtime = float(self.OEE_data_frame['available_time_mins'].sum())
            total_A_percentage = total_runtime / (total_operation_hours)
            total_P_percentage = (
                self.cycle_time_value * (total_OK_qty + total_NG_qty)) / (total_runtime*60)
            total_Q_percentage = total_OK_qty / (total_OK_qty + total_NG_qty)
            total_OEE = total_A_percentage * total_P_percentage * total_Q_percentage
            item0 = QtGui.QStandardItem(self.change_time_format(self.cycle_time_value, 's')['s'] + " secs")
            item0.setTextAlignment(QtCore.Qt.AlignCenter)
            self.OEE_Data_summary_model.setItem(0, 1, item0)
            item1 = QtGui.QStandardItem(f"{total_A_percentage*100:.2f} %")
            item1.setTextAlignment(QtCore.Qt.AlignCenter)
            if total_A_percentage*100 < oee_targets[0][0]:
                item1.setBackground(QtGui.QColor(239, 141, 205,100))
            self.OEE_Data_summary_model.setItem(1, 1, item1)
            item2 = QtGui.QStandardItem(f"{total_P_percentage*100:.2f} %")
            item2.setTextAlignment(QtCore.Qt.AlignCenter)
            if total_P_percentage*100 < oee_targets[0][1]:
                item2.setBackground(QtGui.QColor(239, 141, 205,100))
            self.OEE_Data_summary_model.setItem(2, 1, item2)
            item3 = QtGui.QStandardItem(f"{total_Q_percentage*100:.2f} %")
            item3.setTextAlignment(QtCore.Qt.AlignCenter)
            if total_Q_percentage*100 < oee_targets[0][2]:
                item3.setBackground(QtGui.QColor(239, 141, 205,100))
            self.OEE_Data_summary_model.setItem(3, 1, item3)
            item4 = QtGui.QStandardItem(f"{total_OEE*100:.2f} %")
            item4.setTextAlignment(QtCore.Qt.AlignCenter)
            if total_OEE*100 < oee_targets[0][3]:
                item4.setForeground(QtGui.QColor(255, 0, 0))
            self.OEE_Data_summary_model.setItem(4, 1, item4)
            # mttr = total_loss / downtime_count[0][0] if downtime_count[0][0] > 0 else 0
            # mtbf = total_runtime / downtime_count[0][0] if downtime_count[0][0] > 0 else 0

        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load OEE detail data: {e}")

    @QtCore.pyqtSlot()
    def edit_OEE_data(self, index):
        row = index.row()
        temp_data_frame = self.OEE_data_frame.iloc[row].copy()
        self.edit_dialog = OEE_Edit_Data(database = self.database_process ,data = temp_data_frame, cycle_time = self.cycle_time_value, machine_code = self.OEE_detail_machine_code)        
        self.edit_dialog.accepted.connect(lambda: self.on_oee_edit_accepted(row,temp_data_frame,self.edit_dialog.downtime_records))
        self.edit_dialog.show()    
        
    @QtCore.pyqtSlot()
    def on_oee_edit_accepted(self, row, data, downtime_records):
        downtime_changes = []
        downtime_news = []
        for idx in range(len(downtime_records)):
            if downtime_records[idx].id is None:
                downtime_news.append({ 
                    "date": downtime_records[idx].date,
                    "start_time": downtime_records[idx].start_time,
                    "repair_time": downtime_records[idx].repair_time,
                    "end_time": downtime_records[idx].end_time,
                    "staff_name": downtime_records[idx].staff_name,
                    "error_code": downtime_records[idx].error_code,
                    "machine_code": downtime_records[idx].machine_code,
                    "line_name": downtime_records[idx].line_name })
                continue
            downtime_changes.append({
                "id": downtime_records[idx].id,
                "date": downtime_records[idx].date,
                "start_time": downtime_records[idx].start_time,
                "repair_time": downtime_records[idx].repair_time,
                "end_time": downtime_records[idx].end_time,
                "staff_name": downtime_records[idx].staff_name,
                "error_code": downtime_records[idx].error_code,
                "machine_code": downtime_records[idx].machine_code,
                "line_name": downtime_records[idx].line_name })
        try:
            self.database_process.query(sql='''UPDATE `line_operation_times`
                                                SET operation_hours = :working_shift_hours
                                                WHERE line_id = (SELECT line_id FROM production_lines WHERE line_name = :line_name) and model_running = :model_name
                                                AND operation_date = :production_date;
                                            ''', params={"working_shift_hours": data["working_shift_hours"], "break_time": data["break_time"], "setup_time": data["setup_time"], "line_name": data["line_name"], "model_name": data["model_name"], "production_date": data["production_date"]})
            self.database_process.query(sql='''UPDATE `production_output`
                                                SET OK_qty = :fgs_output_pcs, NG_qty = :defect_pcs
                                                WHERE line_id = (SELECT line_id FROM production_lines WHERE line_name = :line_name)
                                                AND model_name = :model_name
                                                AND production_date = :production_date;
                                            ''', params={"fgs_output_pcs": data["fgs_output_pcs"], "defect_pcs": data["defect_pcs"], "line_name": data["line_name"], "model_name": data["model_name"], "production_date": data["production_date"]})
            if downtime_changes:
                self.database_process.executemany(sql='''UPDATE `downtime_records`
                                                        SET downtime_date = :date, downtime_start_time = :start_time, downtime_start_repair_time = :repair_time
                                                        , downtime_end_time = :end_time, staff_name = :staff_name, error_code = :error_code
                                                        WHERE downtime_record_id = :id;''', params_list=downtime_changes)
            if downtime_news:
                self.database_process.executemany(sql='''INSERT INTO `downtime_records`
                                                        (downtime_date, downtime_start_time, downtime_start_repair_time, downtime_end_time, staff_name, error_code, machine_id, line_id)
                                                        VALUES (:date, :start_time, :repair_time, :end_time, :staff_name, :error_code, 
                                                  (SELECT machine_id FROM machines WHERE machine_code = :machine_code), 
                                                  (SELECT line_id FROM production_lines WHERE line_name = :line_name));''', params_list=downtime_news)
            self.OEE_detail_data()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to update OEE data: {e}")
    
    @QtCore.pyqtSlot()
    def Load_OEE_Data_Other(self):
        self.OEE_Other_data_dialog = OEE_Other_Data(parent = self, database=self.database_process)
        self.OEE_Other_data_dialog.exec()

    @QtCore.pyqtSlot()
    def delete_OEE_data(self, index):
        print("Delete OEE data at row:", index.row(), "Column:", index.column())

# ==========================Function of OEE page ====================================================================================END
# ==========================Function of OEE page ====================================================================================END
# ==========================Function of OEE page ====================================================================================END

# ==========================Function of Maintenance page =============================================================================BEGIN
# ==========================Function of Maintenance page =============================================================================BEGIN
# ==========================Function of Maintenance page =============================================================================BEGIN


    @QtCore.pyqtSlot()
    def Maintenance_page(self):
        self.spinner.start()
        self.ui.main_stacked.setCurrentWidget(self.ui.Maintenance_page)
        self.scan_QRcode = Scan_record_process()
        self.set_stylesheet_change_page((self.ui.Maintenance_btn, self.ui.OEE_btn,
                                        self.ui.Home_btn, 
                                        self.ui.KPI_btn,
                                        self.ui.Stock_btn, self.ui.Downtime_btn))
        if not self.is_expanded:
            self.is_expanded = True
            self.expand_windown_animation(self.is_expanded)
        self.Mainten_Home_page()
        QtCore.QTimer.singleShot(1000, self.spinner.stop)

    @QtCore.pyqtSlot()
    def Mainten_Home_page(self):
        self.style_button_with_shadow((self.ui.Main_Home_btn, self.ui.Main_detail_plan_btn,
                                      self.ui.Main_Input_record_btn, self.ui.Main_Print_record_btn))
        self.ui.Maintenance_stacked.setCurrentWidget(self.ui.Home_page_M)

        def job():
            result = self.database_process.query(sql='''
                SELECT machine_code,machine_name,department_name,line_name, working_week,status 
                FROM maintenance_with_status
                ORDER BY working_week ASC
            ''')
            kpi_status = self.database_process.query(sql='''SELECT (1 - COUNT(CASE WHEN mp.status = "Overdue" THEN 1 END)/COUNT(CASE WHEN mp.status IN ('Ontime','Overdue') THEN 1 END)) AS kpi_maintenance
                                                            FROM Maintenance_plan mp
                                                            JOIN `Months_Years` as my ON my.month_year_id = mp.month_year_id
                                                            WHERE my.year = :year;''', params={"year": dt.datetime.now().year})
            return (result, kpi_status[0][0] if kpi_status else 1)

        self.worker = WorkerThread(job)
        self.worker.finished.connect(
            lambda result: self.on_home_page_data_ready(result[0], result[1]))
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.start()

    @QtCore.pyqtSlot()
    def on_home_page_data_ready(self, result, kpi):
        try:
            headers = ["Code", "Name", "Group", "Line",
                       "Working\nWeek", "Status", "Action"]
            self.data_model = QtGui.QStandardItemModel()
            self.data_model.setHorizontalHeaderLabels(headers)
            self.add_data_to_model(
                result, self.ui.Maintenance_table, self.data_model, callback=self.count_equipment)
            self.ui.Maintenance_actual_num.setText(f"{kpi*100:.0f}%")
            if kpi < 1.0:
                self.ui.Maintenance_actual_num.setStyleSheet("color: red; font-weight: bold;")
                self.ui.Maintenance_status.setText(f'<span style="color: red; font-weight: bold; font-size: 14px;"> ● Missed target</span>')
            else:
                self.ui.Maintenance_actual_num.setStyleSheet("color: green; font-weight: bold;")
                self.ui.Maintenance_status.setText(f'<span style="color: green; font-weight: bold; font-size: 14px;"> ● Reached target</span>')
            # self.ui.
            delegate = StatusColorDelegate(self.ui.Maintenance_table)
            self.ui.Maintenance_table.setItemDelegate(delegate)
            delegate_btn = ButtonDelegate(buttons=("Detail", "Update"))
            self.ui.Maintenance_table.setItemDelegateForColumn(6, delegate_btn)
            self.safe_connect(delegate_btn.ButtonClicked, lambda name,
                              idx: self.on_delegate_btn_clicked(name, idx))
            self.ui.Maintenance_table.setMouseTracking(True)
            self.ui.Maintenance_table.viewport().setMouseTracking(True)
            self.ui.Maintenance_table.setSortingEnabled(True)
            self.ui.Maintenance_table.setColumnWidth(1, 230)
            for i in range(2, 7):
                self.ui.Maintenance_table.setColumnWidth(i, 80)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")
        self.monitor_week_page()

    def add_data_to_model(self, data, target, model, callback=None, column_range=None, tooltip_Enable=False , target_for_item = None , highlight_color = None):
        model.removeRows(0, model.rowCount())
        for row in data:
            items = []
            display_row = row
            if isinstance(column_range, tuple) and len(column_range) >= 2:
                display_row = row[column_range[0]:column_range[1]]
            if tooltip_Enable:
                row_tooltip = f'''<b>More Information:</b>
                <br/>Issue Description: {row[10] if row[10] else "N/A"}
                <br/>Corrective Action: {row[11] if row[11] else "N/A"}
                '''
            for idx, col in enumerate(display_row):
                item = QtGui.QStandardItem(str(col) if col is not None else "")
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                if tooltip_Enable:
                    item.setToolTip(row_tooltip)
                items.append(item)
                if target == self.ui.OEE_Data_table and target_for_item is not None:
                    if idx in target_for_item and not ( col >= target_for_item[idx] and col <= 100):
                        item.setBackground(QtGui.QColor(*highlight_color))
                        item.setForeground(QtGui.QColor(255, 0, 0))

            model.appendRow(items)
        if callback is not None and callable(callback):
            callback()
        target.setModel(model)

    @QtCore.pyqtSlot()
    def on_delegate_btn_clicked(self, name, index):
        model = index.model()
        row = index.row()
        code = model.data(model.index(row, 0))
        dep = model.data(model.index(row, 2))
        if name == "Detail":
            self.detail_machine_information = Machine_information(
                database=self.database_process, code=code)
            self.detail_machine_information.show()
        else:
            if self.login_info["role_level"] in ["Manager", "Admin"]:
                pass
            elif (self.login_info["department"] == dep) and (self.login_info["role_level"] == "Supervisor"):
                pass
            else:
                QtWidgets.QMessageBox.information(
                    self, "Permission denied", "Your don't have permission to update this machine info")
                return
            self.update_info_dialog = Update_machine_info(
                parent=self, code=code)
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
            self.ui.status_cbb.addItems(
                ["", "Upcoming", "Near due", "Overdue", "No schedule"])
            self.safe_connect(self.ui.apply_btn.clicked, self.filter_process)
            self.safe_connect(self.ui.cancel_btn.clicked, self.hide_filter)
            self.safe_connect(self.ui.code_lnedit.textChanged, lambda: self.filter_suggestion(
                self.ui.code_lnedit, "machine_code", "maintenance_with_status"))
            self.safe_connect(self.ui.name_lnedit.textChanged, lambda: self.filter_suggestion(
                self.ui.name_lnedit, "machine_name", "maintenance_with_status"))
            self.safe_connect(self.ui.group_cbb.currentTextChanged,
                              self.group_cbb_Home_Maintenance_change)
        self.ui.filter_mainten_frame.show()

    @QtCore.pyqtSlot()
    def hide_filter(self):
        self.ui.filter_mainten_frame.hide()

    @QtCore.pyqtSlot()
    def filter_suggestion(self, target, text, table, where=None):
        if len(target.text()) < 2:
            return
        suggestions = []
        machine_code = []
        script = f'''SELECT {text} FROM {table}'''
        if where != None:
            script = script + where + "LIMIT 10;"
        else:
            script = script + \
                f" WHERE {text} LIKE '%{target.text()}%' " + "LIMIT 10;"
        try:
            machine_code = self.database_process.query(sql=script)
            suggestions = [str(name[0]) if len(
                name) == 1 else f"{name[0]} : {name[1]}" for name in machine_code]
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to fetch machine names: {e}")
            suggestions = []
        if suggestions:
            self.completer = QtWidgets.QCompleter(suggestions, self)
            self.completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
            popup = self.completer.popup()
            popup.setAlternatingRowColors(True)
            fm = popup.fontMetrics()
            max_width = max(fm.horizontalAdvance(s) for s in suggestions) + 30
            popup.setMinimumWidth(max_width)
            target.setCompleter(self.completer)

    @QtCore.pyqtSlot()
    def filter_process(self):
        try:
            query = []
            if self.ui.code_lnedit.text() != "":
                query.append(
                    f'machine_code LIKE "%{self.ui.code_lnedit.text()}%"')
            if self.ui.name_lnedit.text() != "":
                query.append(
                    f'machine_name LIKE "%{self.ui.name_lnedit.text()}%"')
            if self.ui.group_cbb.currentText() != "":
                query.append(
                    f'department_name = "{self.ui.group_cbb.currentText()}"')
            if self.ui.line_cbb.currentText() != "":
                query.append(f'line_name = "{self.ui.line_cbb.currentText()}"')
            if self.ui.status_cbb.currentText() != "":
                query.append(
                    f'status COLLATE utf8mb4_unicode_ci = "{self.ui.status_cbb.currentText()}"')
            query = " AND ".join(query)
            if query == "":
                result = self.database_process.query(sql='''SELECT machine_code,machine_name,department_name,line_name,working_week,status 
                                                            FROM maintenance_with_status
                                                            ORDER BY next_due_date ASC''')
                self.add_data_to_model(
                    result, self.ui.Maintenance_table, self.data_model, callback=self.count_equipment)
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
        self.add_data_to_model(result, self.ui.Maintenance_table,
                               self.data_model, callback=self.count_equipment)
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
                                                            WHERE department_name = :dep''', params={'dep': dep})
            self.ui.line_cbb.addItems([""] + [line[0] for line in line_list])
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    def count_equipment(self):
        try:
            upcoming = self.database_process.query(sql='''SELECT COUNT(status)
                                                        FROM maintenance_with_status
                                                        WHERE status COLLATE utf8mb4_unicode_ci = "Upcoming";''')
            # overdue = self.database_process.query(sql='''SELECT COUNT(status)
            #                                             FROM maintenance_with_status
            #                                             WHERE status COLLATE utf8mb4_unicode_ci = "Overdue";''')
            near_due = self.database_process.query(sql='''SELECT COUNT(status)
                                                        FROM maintenance_with_status
                                                        WHERE status COLLATE utf8mb4_unicode_ci = "Near due";''')

            def set_fontsize(target,num, color = (0, 255, 0)):
                if num > 999:
                    self.draw_circle(target, 90, color, num, font_size = 14)
                    return
                elif num > 99:
                    self.draw_circle(target, 90, color, num, font_size = 20)
                    return
                else:
                    return
            set_fontsize(self.ui.upcoming_label, upcoming[0][0], color = (0, 255, 0))
            set_fontsize(self.ui.neardue_label, near_due[0][0], color = (196, 41, 0))
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to count data: {e}")

    @QtCore.pyqtSlot()
    def monitor_week_page(self):
        self.ui.monitor_stacked.setCurrentWidget(self.ui.monitor_week_page)
        self.style_button_with_shadow(
            (self.ui.weekly_btn, self.ui.monthly_btn, self.ui.inyear_btn))

        def job():
            Total = self.database_process.query(
                ''' SELECT COUNT(DISTINCT mp.line_id) AS total
                    FROM `Maintenance_plan` mp
                    JOIN `Production_Lines` p ON p.line_id = mp.line_id
                    JOIN `Departments` d ON p.department_id = d.department_id
                    JOIN `Months_Years` as my ON mp.month_year_id = my.month_year_id
                    WHERE week = :week AND d.department_id < 7 AND my.year = :year AND (mp.status IN ('Ontime','Overdue') OR mp.status IS NULL);''',
                params={'week': self.week_num, 'year': self.year_num}
            )
            sql = '''SELECT p.line_name, COUNT(p.line_name) AS plan_count
                    FROM Maintenance_plan mp
                    JOIN Production_Lines p ON p.line_id = mp.line_id
                    JOIN Departments d ON p.department_id = d.department_id
                    JOIN `Months_Years` as my ON mp.month_year_id = my.month_year_id
                    WHERE d.department_name = :dept AND mp.week = :week AND my.year = :year AND (mp.status IN ('Ontime','Overdue') OR mp.status IS NULL)
                    GROUP BY p.line_name;'''
            PE1 = self.database_process.query(
                sql=sql, params={'week': self.week_num, 'dept': "PE1", 'year': self.year_num})
            PE2 = self.database_process.query(
                sql=sql, params={'week': self.week_num, 'dept': "PE2", 'year': self.year_num})
            PE3 = self.database_process.query(
                sql=sql, params={'week': self.week_num, 'dept': "PE3", 'year': self.year_num})
            PE4 = self.database_process.query(
                sql=sql, params={'week': self.week_num, 'dept': "PE4", 'year': self.year_num})
            PE5 = self.database_process.query(
                sql=sql, params={'week': self.week_num, 'dept': "PE5", 'year': self.year_num})
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
            result = self.database_process.query(
                sql=sql2, params={'week': self.week_num, 'year': self.year_num})

            return {"Total": Total, "PE1": PE1, "PE2": PE2, "PE3": PE3,
                    "PE4": PE4, "PE5": PE5, "Result": result}
        self.worker = WorkerThread(job)
        self.worker.finished.connect(
            lambda data: self.on_monitor_data_ready(data=data))
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.start()

    @QtCore.pyqtSlot()
    def on_monitor_data_ready(self, data):
        try:
            Total, PE1, PE2, PE3, PE4, PE5, result = (
                data["Total"], data["PE1"], data["PE2"],
                data["PE3"], data["PE4"], data["PE5"], data["Result"]
            )

            headers = ["Item", "Content"]
            monitor_model = QtGui.QStandardItemModel()
            monitor_model.setHorizontalHeaderLabels(headers)

            PE1_line = []
            PE2_line = []
            PE3_line = []
            PE4_line = []
            PE5_line = []
            machine_qty = [0, 0, 0, 0, 0]

            def insert_into_item(data: list, target: list, group: str):
                if data and data[0][0] is not None:
                    target.extend([(row[0],) for row in data])
                    self.insert_item(target, group, monitor_model)
                else:
                    self.insert_item("", group, monitor_model)

            def total_machine(data: list, target: list, index: int):
                machine_qty[index] = sum([data[i][1]
                                         for i in range(len(data))])
            self.insert_item([(str(Total[0][0]),)], "Total", monitor_model)
            insert_into_item(PE1, PE1_line, "PE1")
            insert_into_item(PE2, PE2_line, "PE2")
            insert_into_item(PE3, PE3_line, "PE3")
            insert_into_item(PE4, PE4_line, "PE4")
            insert_into_item(PE5, PE4_line, "PE5")
            total_machine(PE1, machine_qty, 0)
            total_machine(PE2, machine_qty, 1)
            total_machine(PE3, machine_qty, 2)
            total_machine(PE4, machine_qty, 3)
            total_machine(PE5, machine_qty, 4)

            self.ui.week_plan_table_line.setModel(monitor_model)
            self.ui.week_plan_table_line.setColumnWidth(0, 10)
            self.ui.week_plan_table_line.resizeRowsToContents()

            self.draw_monitor_chart(target=self.ui.week_plan_chart_line, lines=["PE1", "PE2", "PE3", "PE4", "PE5"],
                                    plan=[len(PE1), len(PE2), len(
                                        PE3), len(PE4), len(PE5)],
                                    result=[result[0][0], result[1][0], result[2][0], result[3][0], result[4][0]])
            self.draw_monitor_chart(target=self.ui.week_plan_chart_mc, lines=["PE1", "PE2", "PE3", "PE4", "PE5"],
                                    plan=machine_qty,
                                    result=[result[0][1], result[1][1], result[2][1], result[3][1], result[4][1]], set_title=True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    def insert_item(self, data, name, model):
        data = set(data)
        str_line = []
        for item in data:
            str_line.append(item[0])
        str_line = ", ".join(str_line)
        row = [QtGui.QStandardItem(name), QtGui.QStandardItem(str_line)]
        model.appendRow(row)

    def draw_monitor_chart(self, target, lines, plan, result, set_title=False):
        layout = target.layout()
        if layout is None:
            layout = QtWidgets.QVBoxLayout(target)
            target.setLayout(layout)
        if not hasattr(target, "canvas"):
            fig, ax = plt.subplots(figsize=(5, 3))
            target.canvas = FigureCanvas(fig)
            target.canvas.setFixedSize(target.width()-5, target.height()-10)
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
        plan_col = ax.bar([i - 0.2 for i in x], plan, width=0.4,
                          label="Plan", color=(18/255, 184/255, 234/255, 1))
        result_col = ax.bar([i + 0.2 for i in x], result, width=0.4,
                            label="Result", color=(63/255, 218/255, 155/255, 1))
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
        self.style_button_with_shadow(
            (self.ui.monthly_btn, self.ui.weekly_btn, self.ui.inyear_btn))
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
                                                params={'month': self.month_num, 'year': self.year_num})
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
            PE1 = self.database_process.query(
                sql=sql, params={'dept': "PE1", 'month': self.month_num, 'year': self.year_num})
            PE2 = self.database_process.query(
                sql=sql, params={'dept': "PE2", 'month': self.month_num, 'year': self.year_num})
            PE3 = self.database_process.query(
                sql=sql, params={'dept': "PE3", 'month': self.month_num, 'year': self.year_num})
            PE4 = self.database_process.query(
                sql=sql, params={'dept': "PE4", 'month': self.month_num, 'year': self.year_num})
            PE5 = self.database_process.query(
                sql=sql, params={'dept': "PE5", 'month': self.month_num, 'year': self.year_num})
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
            result = self.database_process.query(
                sql=sql, params={'month': self.month_num, 'year': self.year_num})
            headers = ["Item", "Content"]
            monitor_model = QtGui.QStandardItemModel()
            monitor_model.setHorizontalHeaderLabels(headers)
            monitor_model.removeRows(0, monitor_model.rowCount())
            PE1_line = []
            PE2_line = []
            PE3_line = []
            PE4_line = []
            PE5_line = []
            machine_qty = [0, 0, 0, 0, 0]

            def insert_into_item(data: list, target: list, group: str):
                if data and data[0][0] is not None:
                    target.extend([(row[0],) for row in data])
                    self.insert_item(target, group, monitor_model)
                else:
                    self.insert_item("", group, monitor_model)

            def total_machine(data: list, target: list, index: int):
                machine_qty[index] = sum([data[i][1]
                                         for i in range(len(data))])

            self.insert_item([(str(Total[0][0]),)], "Total", monitor_model)
            insert_into_item(PE1, PE1_line, "PE1")
            insert_into_item(PE2, PE2_line, "PE2")
            insert_into_item(PE3, PE3_line, "PE3")
            insert_into_item(PE4, PE4_line, "PE4")
            insert_into_item(PE5, PE5_line, "PE5")
            total_machine(PE1, machine_qty, 0)
            total_machine(PE2, machine_qty, 1)
            total_machine(PE3, machine_qty, 2)
            total_machine(PE4, machine_qty, 3)
            total_machine(PE5, machine_qty, 4)
            self.ui.month_plan_table_line.setModel(monitor_model)
            self.ui.month_plan_table_line.setColumnWidth(0, 10)
            self.ui.month_plan_table_line.resizeRowsToContents()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")
        self.draw_monitor_chart(target=self.ui.month_plan_chart_line, lines=["PE1", "PE2", "PE3", "PE4", "PE5"],
                                plan=[len(PE1), len(PE2), len(
                                    PE3), len(PE4), len(PE5)],
                                result=[result[0][0], result[1][0], result[2][0], result[3][0], result[4][0]])
        self.draw_monitor_chart(target=self.ui.month_plan_chart_mc, lines=["PE1", "PE2", "PE3", "PE4", "PE5"],
                                plan=machine_qty,
                                result=[result[0][1], result[1][1], result[2][1], result[3][1], result[4][1]], set_title=True)

    @QtCore.pyqtSlot()
    def monitor_inyear_page(self):
        def safe_divide(a, b):
            if b == 0:
                return 0
            return a / b
        self.ui.monitor_stacked.setCurrentWidget(self.ui.monitor_year_page)
        self.style_button_with_shadow(
            (self.ui.inyear_btn, self.ui.weekly_btn, self.ui.monthly_btn))
        headers = ["", "Total", "Overdue"]
        model = QtGui.QStandardItemModel()
        model.setHorizontalHeaderLabels(headers)
        self.ui.KPI_table.setModel(model)
        self.ui.KPI_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Stretch)
        total_sql = '''SELECT
                        (SELECT "Total" ) as department_name,
                        COUNT(CASE WHEN mp.status IN ('Ontime','Overdue') THEN 1 END) AS total,
                        COUNT(CASE WHEN mp.status = "Overdue" THEN 1 END) AS overdue
                        FROM Maintenance_plan mp
                        JOIN `Months_Years` as my ON my.month_year_id = mp.month_year_id
                        WHERE my.year = :year;'''
        dep_sql = '''SELECT
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
            result = self.database_process.query(
                sql=total_sql, params={'year': self.year_num})
            result += self.database_process.query(
                sql=dep_sql, params={'year': self.year_num})
            for row in result:
                items = []
                for col in row:
                    item = QtGui.QStandardItem(
                        str(col) if col is not None else "")
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    items.append(item)
                model.appendRow(items)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")
        self.ui.KPI_table.setAlternatingRowColors(True)
        self.create_pie_chart(
            self.ui.total_kpi, result[0][1]-result[0][2], result[0][2])
        self.ui.PE1_KPI.setValue(
            int((1-safe_divide(result[1][2], result[1][1]))*100))
        self.ui.PE2_KPI.setValue(
            int((1-safe_divide(result[2][2], result[2][1]))*100))
        self.ui.PE3_KPI.setValue(
            int((1-safe_divide(result[3][2], result[3][1]))*100))
        self.ui.PE4_KPI.setValue(
            int((1-safe_divide(result[4][2], result[4][1]))*100))
        self.ui.PE5_KPI.setValue(
            int((1-safe_divide(result[5][2], result[5][1]))*100))

    def create_pie_chart(self, target, plan, result, fontdict={'fontsize': 14,
                                                               'fontweight': 'bold',
                                                               'color': '#008b8b',
                                                               'fontname': "Comic Sans MS"
                                                               }, figsize=(3,1.5), x_legend = 0.25):
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
            fig, ax = plt.subplots(figsize=figsize)
            fig.patch.set_alpha(0.0)
            ax.set_facecolor("none")
            target.canvas = FigureCanvas(fig)
            # target.canvas.setFixedSize(target.width(), target.height())
            target.canvas.setSizePolicy(
                    QtWidgets.QSizePolicy.Policy.Expanding,
                    QtWidgets.QSizePolicy.Policy.Expanding
                )
            target.ax = ax
            target.fig = fig
            layout.addWidget(target.canvas)
        else:
            target.ax.clear()

        ax = target.ax
        fig = target.fig

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

        ax.set_title("Maintenance KPI chart", fontdict=fontdict)
        ax.legend(loc="upper right", bbox_to_anchor=(x_legend, 1),
                  fontsize=8, labels=labels)
        fig.tight_layout(pad=0.5)
        target.canvas.draw()

    @QtCore.pyqtSlot()
    def next_monitor_page(self):
        self.ui.monitor_next_btn.setEnabled(False)
        QtCore.QTimer.singleShot(
            300, lambda: self.ui.monitor_next_btn.setEnabled(True))
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
        QtCore.QTimer.singleShot(
            300, lambda: self.ui.monitor_back_btn.setEnabled(True))
        if (self.ui.monitor_stacked.currentWidget() is self.ui.monitor_week_page):
            self.week_num = self.week_num - 1
            if (self.week_num < 1):
                self.week_num = self.ui.company_week_number(
                    dt.date(self.year_num, 12, 31))
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
        self.style_button_with_shadow(
            (self.ui.Main_Print_record_btn, self.ui.Main_detail_plan_btn, self.ui.Main_Home_btn, self.ui.Main_Input_record_btn))
        self.ui.Maintenance_stacked.setCurrentWidget(self.ui.Print_page_M)
        try:
            if self.ui.Group_cbb_PF.count() == 0:
                self.result_print_record = pd.DataFrame(columns=["machine_code", "machine_name", "attached_equipment", "department_name", "line_name", "last_maintenance_date", "week", "issued_maintenance_date", "technician", "form_name", "form_link"])
                self.columns_print_record_dict = {idx: col for idx, col in enumerate(self.result_print_record.columns)}
                self.ui.Group_cbb_PF.addItems(
                    [group[0] for group in self.group])
                self.ui.Group_cbb_PF.setCurrentText(
                    self.login_info['department'])
                self.department_print_record = self.ui.Group_cbb_PF.currentText()
                if self.login_info['role_level'] == 'Admin':
                    self.ui.Group_cbb_PF.setEnabled(True)
                else:
                    self.ui.Group_cbb_PF.setEnabled(False)
                current_week = self.ui.company_week_number(self.ui.today)
                week_lst = [str(week_num)
                            for week_num in range(1, self.qty_week+1)]
                self.ui.FromWeek_cbb_PF.addItems(week_lst)
                self.ui.FromWeek_cbb_PF.setCurrentIndex(current_week - 1)
                self.ui.ToWeek_cbb_PF.addItems(week_lst)
                self.ui.ToWeek_cbb_PF.setCurrentIndex(current_week - 1)
                header = ["Machine code", "Machine Name", "Attached\nequipment", "Group", "Line",
                          "Last maintenance\n date", "Week plan", "Issued Maintenance\n date", "Technical", "Form type"]
                self.ui.print_record_table.setColumnCount(len(header))
                self.ui.print_record_table.setHorizontalHeaderLabels(header)
                self.ui.print_record_table.setColumnWidth(0, 100)
                self.ui.print_record_table.setColumnWidth(1, 300)
                self.ui.print_record_table.setColumnWidth(2, 100)
                self.ui.print_record_table.setColumnWidth(3, 100)
                self.ui.print_record_table.setColumnWidth(4, 100)
                self.ui.print_record_table.setColumnWidth(5, 100)
                self.ui.print_record_table.setColumnWidth(6, 100)
                self.ui.print_record_table.setColumnWidth(7, 120)
                self.ui.print_record_table.setColumnWidth(8, 100)
                self.ui.print_record_table.setColumnWidth(9, 300)
                self.ui.print_record_table.setSortingEnabled(True)
                self.ui.print_record_table.setMouseTracking(True)
                self.ui.print_record_table.viewport().setMouseTracking(True)
                self.button_delegate = ButtonDelegate(buttons=["+"], target_indexes=set())
                self.ui.print_record_table.setItemDelegateForColumn(2, self.button_delegate)
                self.button_delegate.ButtonClicked.connect(lambda name,index: self.on_add_attached_equipment_clicked(name, index))
                self.add_item_line_PF()
                self.safe_connect(
                    self.ui.FromWeek_cbb_PF.currentIndexChanged, self.check_cbb)
                self.safe_connect(
                    self.ui.ToWeek_cbb_PF.currentIndexChanged, self.check_cbb)
                self.safe_connect(self.ui.Load_btn.clicked,
                                  lambda _: self.load_record_form())
                self.safe_connect(self.ui.insert_row_btn_PF.clicked, lambda _: self.insert_row(
                    target=self.ui.print_record_table))
                self.safe_connect(self.ui.del_row_btn_PF.clicked, lambda _: self.delete_row( form = "print_record",
                    target=self.ui.print_record_table, save_list=self.result_print_record))
                self.safe_connect(self.ui.print_btn_PF.clicked,
                                  lambda _: self.print_record())
                self.safe_connect(self.ui.Update_form_btn.clicked,
                                  lambda _: self.update_register_form())
                self.safe_connect(self.ui.Register_form_btn.clicked,
                                  lambda _: self.register_new_form())
                self.safe_connect(self.ui.Clear_btn.clicked,
                                  lambda _: self.clear_print_data())
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Fail to load data: {e}")

    @QtCore.pyqtSlot()
    def add_item_line_PF(self):
        try:
            lines = self.database_process.query(sql='''SELECT p.line_name
                                                        FROM `Production_Lines` as p 
                                                        JOIN `Departments` as d
                                                        ON p.department_id = d.department_id
                                                        WHERE d.department_name = :dep
                                                        ORDER BY p.line_name ASC''', params={'dep': self.ui.Group_cbb_PF.currentText()})
            items = ["All"] + [line[0] for line in lines]
            self.ui.Line_cbb_PF.clear()
            self.ui.Line_cbb_PF.addItems(items)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot()
    def check_cbb(self):
        if self.ui.ToWeek_cbb_PF.currentIndex() >= self.ui.FromWeek_cbb_PF.currentIndex():
            return
        QtWidgets.QMessageBox.warning(
            self, "Wrong select", f"""You have selected the wrong format."From week" must be smaller than "To Week", please select again.""")
        self.ui.ToWeek_cbb_PF.setCurrentIndex(
            self.ui.FromWeek_cbb_PF.currentIndex())

    def on_add_attached_equipment_clicked(self, name, index):
        if name == "+":
            row, col = index.row(), index.column()
            dialog = QtWidgets.QDialog(self)
            dialog.setWindowTitle("Add Attached Equipment")
            layout = QtWidgets.QVBoxLayout(dialog)
            list_widget = QtWidgets.QListWidget()
            list_widget.setSpacing(8) 
            line_edits = []

            @QtCore.pyqtSlot(str)
            def add_equipment_row(text=""):
                insert_pos = max(0, list_widget.count() - 1)
                item = QtWidgets.QListWidgetItem()
                line_edit = QtWidgets.QLineEdit()
                line_edit.setPlaceholderText(f"Equipment attach")
                line_edit.setText(text)
                line_edit.setMinimumHeight(30)
                line_edit.setStyleSheet('background-color: rgba(0, 0, 0, 0.05); border-radius: 5px; padding: 5px;')
                line_edit.setFrame(False)
                line_edit.textChanged.connect(lambda text: self.filter_suggestion(line_edit, "DISTINCT (machine_code)", "maintenance_form_info", f'''
                                                                                                        WHERE department_name = "{self.login_info['department']}"
                                                                                                        AND machine_code LIKE "%{text}%"'''))
                line_edit.editingFinished.connect(lambda: on_equipment_editing_finished(line_edit))
                list_widget.insertItem(insert_pos, item)
                list_widget.setItemWidget(item, line_edit)
                item.setSizeHint(line_edit.sizeHint())
                line_edits.append(line_edit)
            @QtCore.pyqtSlot()
            def on_equipment_editing_finished(_line_edit):
                local_line_edits = _line_edit
                text = local_line_edits.text().strip()
                if text:
                    current_week = self.ui.company_week_number(self.ui.today)
                    result = self.database_process.query(sql=f''' SELECT machine_code FROM maintenance_form_info
                                                        WHERE machine_code = :machine_code
                                                        AND department_name = :department_name AND `week` BETWEEN :week_from AND :week_to;''',
                                                        params={'machine_code': text, 'department_name': self.login_info['department'], 'week_from': current_week - 4, 'week_to': current_week + 4})
                    if not result:
                        local_line_edits.clear()
                        QtWidgets.QMessageBox.warning(dialog, "Warning", "No maintenance record found for the selected equipment in the current week range.")

            add_item = QtWidgets.QListWidgetItem()
            add_btn = QtWidgets.QPushButton("Add equipment")
            add_btn.clicked.connect(lambda: add_equipment_row())
            list_widget.addItem(add_item)
            list_widget.setItemWidget(add_item, add_btn)
            add_item.setSizeHint(add_btn.sizeHint())

            button_layout = QtWidgets.QHBoxLayout()
            ok_btn = QtWidgets.QPushButton("OK")
            cancel_btn = QtWidgets.QPushButton("Cancel")
            ok_btn.clicked.connect(dialog.accept)
            cancel_btn.clicked.connect(dialog.reject)
            button_layout.addWidget(ok_btn)
            button_layout.addWidget(cancel_btn)

            layout.addWidget(list_widget)
            layout.addLayout(button_layout)

            if dialog.exec_() == QtWidgets.QDialog.Accepted:
                attached_equipment = [
                    le.text().strip() for le in line_edits if le.text().strip()
                ]
                editor_main = self.ui.print_record_table.cellWidget(row, 0)
                if editor_main is None:
                    QtWidgets.QMessageBox.critical(self, "Error", f"Row {row+1}: missing machine code editor")
                    return
                main_machine = editor_main.text().strip()
                if main_machine == "":
                    QtWidgets.QMessageBox.critical(self, "Error", f"Row {row+1}: machine code is empty")
                    return
                self.result_print_record.loc[self.result_print_record["machine_code"] == main_machine, "attached_equipment"] = attached_equipment
                self.add_item_into_print_record(self.result_print_record, c=2, r=row, is_readOnly=True, widget = "attached_equipment")

    def add_item_into_print_record(self, data: list, c=None, r=None, is_readOnly=None, widget=None):
        if c == 2:
            value = str(data.loc[r, self.columns_print_record_dict[c]]).strip()
            item = QtWidgets.QTableWidgetItem()
            item.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            if value != "" and value != "None":
                self.button_delegate._target_indexes.discard((r, 2))
                item.setText(value)
            else:
                self.button_delegate._target_indexes.add((r, 2))
                item.setText("")

            self.ui.print_record_table.setItem(r, 2, item)
            self.ui.print_record_table.viewport().update()
            return
        if widget == None:
            editor = QtWidgets.QLineEdit()
            if len(data) == 1:
                r_data = 0
            else:
                r_data = r
            editor.setText(str(data.loc[r_data,self.columns_print_record_dict[c]]))
            editor.setFrame(False)
            editor.setAlignment(QtCore.Qt.AlignCenter)
            if not is_readOnly:
                editor.setStyleSheet("background-color: rgba(0,0,0,0.05);")
            else:
                self.safe_connect(editor.textChanged, self.handle_text_changed)
                self.safe_connect(editor.returnPressed,self.handle_editing_finished)
            editor.setEnabled(is_readOnly)
            self.ui.print_record_table.setCellWidget(r, c, editor)
        else:
            widget.setText(str(data.loc[0,self.columns_print_record_dict[c]]))

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
        filter_script = [
            f" department_name = '{self.ui.Group_cbb_PF.currentText()}'"]
        if self.ui.Line_cbb_PF.currentText() != "All":
            filter_script.append(
                f"line_name = '{self.ui.Line_cbb_PF.currentText()}'")
        filter_script = filter_script + \
            [f"week >= {self.ui.FromWeek_cbb_PF.currentText()}",
             f"week <= {self.ui.ToWeek_cbb_PF.currentText()}"]
        filter_script = " AND ".join(filter_script)
        try:
            result_print_record = self.database_process.query(sql=f'''SELECT machine_code,machine_name,NULL,department_name,line_name,last_maintenance_date, week,(SELECT CURRENT_DATE),technician,form_name,form_link FROM  maintenance_form_info
                                                            WHERE {filter_script}''')
            if len(result_print_record) == 0:
                raise ValueError("Don't see machine in maintenance plan")
            self.ui.print_record_table.setRowCount(len(result_print_record))
            self.result_print_record.drop(self.result_print_record.index, inplace=True)
            result_print_record = pd.DataFrame(result_print_record, columns=self.result_print_record.columns)
            self.result_print_record = pd.concat(
                [self.result_print_record, result_print_record],
                ignore_index=True
            )
            for row in range(len(self.result_print_record)):
                self.add_item_into_print_record(
                    self.result_print_record, c= 0, r=row, is_readOnly=True)
                self.add_item_into_print_record(
                    self.result_print_record, c= 1, r=row, is_readOnly=False)
                self.add_item_into_print_record(
                    self.result_print_record, c= 2, r=row, is_readOnly=True)
                self.add_item_into_print_record(
                    self.result_print_record, c= 3, r=row, is_readOnly=False)
                self.add_item_into_print_record(
                    self.result_print_record, c= 4, r=row, is_readOnly=True)
                self.add_item_into_print_record(
                    self.result_print_record, c= 5, r=row, is_readOnly=False)
                self.add_item_into_print_record(
                    self.result_print_record, c= 6, r=row, is_readOnly=False)
                self.add_item_into_print_record(
                    self.result_print_record, c= 7, r=row, is_readOnly=True)
                self.add_item_into_print_record(
                    self.result_print_record, c= 8, r=row, is_readOnly=True)
                self.add_item_into_print_record(
                    self.result_print_record, c= 9, r=row, is_readOnly=False)
            # delegate = DynamicSuggestion(
            #     database=self.database_process, dep=self.ui.Group_cbb_PF.currentText(), year=self.year_num)
            # self.ui.print_record_table.setItemDelegateForColumn(2, delegate)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot()
    def on_text_changed(self, text, c, r, target, select_col, table, where):
        if c == 0:
            self.filter_suggestion(target.cellWidget(
                r, c), select_col, table, where)
        elif c == 4:
            editor = target.cellWidget(r, 3)
            value = editor.text()
            self.filter_suggestion(target.cellWidget(r, c), "p.line_name", "`Production_Lines` as p ", f'''JOIN `Departments` as d
                                                                                                        ON p.department_id = d.department_id
                                                                                                        WHERE d.department_name = "{value}"
                                                                                                        AND p.line_name LIKE "%{text}%"''')

    @QtCore.pyqtSlot()
    def reload_data(self, r, c):
        editor = self.ui.print_record_table.cellWidget(r, c)
        if editor is None:
            return
        if c != 0:
            value = editor.text()
            if c == 4:
                dep_editor = self.ui.print_record_table.cellWidget(r, 3)
                dep = dep_editor.text()
                iscorrectDep = self.database_process.query(sql=''' SELECT 1 FROM `Production_Lines` as p
                                                                    JOIN Departments as d
                                                                    ON p.department_id = d.department_id
                                                                    WHERE p.line_name = :line AND d.department_name = :dep ''', params={'line': value, 'dep': dep})
                if not iscorrectDep:
                    editor.clear()
                    return
            edit_row = list(self.result_print_record.loc[r])
            edit_row[c] = value.upper()
            self.result_print_record.loc[r] = tuple(edit_row)
            return
        editor = self.ui.print_record_table.cellWidget(r, 0)
        value = editor.text()
        try:
            result = self.database_process.query(sql=f'''SELECT machine_code,machine_name,NULL,department_name,line_name,last_maintenance_date, week,(SELECT CURRENT_DATE),technician,form_name,form_link 
                                                            FROM  maintenance_form_info
                                                            WHERE machine_code = :code AND department_name = :dep
                                                            ORDER BY week ASC LIMIT 1;''', params={'code': value, 'dep': self.ui.Group_cbb_PF.currentText()})
            if len(result) == 0:
                for col in range(self.ui.print_record_table.columnCount()):
                    editor = self.ui.print_record_table.cellWidget(r, col)
                    if editor is None: continue
                    editor.clear()
                self.result_print_record.loc[r] = [""] * len(self.result_print_record.columns)
                raise ValueError("Don't see machine in maintenance plan")
                
            current_week = self.ui.company_week_number(self.ui.today)
            if (result[0][6] > (current_week + 4)) or (result[0][6] < (current_week - 4)): #LỖI Ở ĐÂY, NHẤN 2 LẦN LÀ SẼ CHO IN MÁY QUÁ HẠN
                for col in range(self.ui.print_record_table.columnCount()):
                    editor = self.ui.print_record_table.cellWidget(r, col)
                    if editor is None: continue
                    editor.clear()
                self.result_print_record.loc[r] = [""] * len(self.result_print_record.columns)
                raise ValueError(
                    "This machine is not in the maintenance plan of time")
            result = pd.DataFrame(result, columns=self.result_print_record.columns)
            self.result_print_record.loc[r] = result.loc[0]
            self.add_item_into_print_record(
                data=result, c=0, is_readOnly=True, widget=self.ui.print_record_table.cellWidget(r, 0))
            self.add_item_into_print_record(
                data=result, c=1, is_readOnly=False, widget=self.ui.print_record_table.cellWidget(r, 1))
            self.add_item_into_print_record(
                data=result, c=3, is_readOnly=False, widget=self.ui.print_record_table.cellWidget(r, 3))
            self.add_item_into_print_record(
                data=result, c=4, is_readOnly=True, widget=self.ui.print_record_table.cellWidget(r, 4))
            self.add_item_into_print_record(
                data=result, c=5, is_readOnly=False, widget=self.ui.print_record_table.cellWidget(r, 5))
            self.add_item_into_print_record(
                data=result, c=6, is_readOnly=False, widget=self.ui.print_record_table.cellWidget(r, 6))
            self.add_item_into_print_record(
                data=result, c=7, is_readOnly=True,  widget=self.ui.print_record_table.cellWidget(r, 7))
            self.add_item_into_print_record(
                data=result, c=8, is_readOnly=True,  widget=self.ui.print_record_table.cellWidget(r, 8))
            self.add_item_into_print_record(
                data=result, c=9, is_readOnly=False, widget=self.ui.print_record_table.cellWidget(r, 9))
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Fail to load data: {e}")

    @QtCore.pyqtSlot()
    def insert_row(self, form=None, target=None):
        current_row = target.rowCount()
        target.insertRow(current_row)
        if form == None:
            new_data = { col: "" for col in self.columns_print_record_dict.values()}
            self.result_print_record.loc[current_row] = new_data
            self.add_item_into_print_record(
                data=self.result_print_record, c=0, r=current_row, is_readOnly=True)
            self.add_item_into_print_record(
                data=self.result_print_record, c=1, r=current_row, is_readOnly=False)
            self.add_item_into_print_record(
                data=self.result_print_record, c=2, r=current_row, is_readOnly=True)
            self.add_item_into_print_record(
                data=self.result_print_record, c=3, r=current_row, is_readOnly=False)
            self.add_item_into_print_record(
                data=self.result_print_record, c=4, r=current_row, is_readOnly=True)
            self.add_item_into_print_record(
                data=self.result_print_record, c=5, r=current_row, is_readOnly=False)
            self.add_item_into_print_record(
                data=self.result_print_record, c=6, r=current_row, is_readOnly=False)
            self.add_item_into_print_record(
                data=self.result_print_record, c=7, r=current_row, is_readOnly=True)
            self.add_item_into_print_record(
                data=self.result_print_record, c=8, r=current_row, is_readOnly=True)
            self.add_item_into_print_record(
                data=self.result_print_record, c=9, r=current_row, is_readOnly=False)
        return current_row

    @QtCore.pyqtSlot()
    def delete_row(self, form=None, target=None, save_list=None, oneFile=None):
        current_row = target.currentRow()
        if form == "print_record":
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
                save_list.drop(save_list.index[current_row], inplace=True)
                save_list.reset_index(drop=True, inplace=True) 
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Error: {e}")
            self.button_delegate._target_indexes = {
                    (rr - 1 if rr > current_row else rr, cc)
                    for (rr, cc) in self.button_delegate._target_indexes
                    if rr != current_row
                }
            self.button_delegate._target_indexes = {
                (r, 2)
                for r in range(len(save_list))
                if str(save_list.iloc[r][self.columns_print_record_dict[2]]).strip() in ("", "None", "nan")
            }
            self.button_delegate._buttons.clear()
            self.ui.print_record_table.viewport().update()
        else:
            try:
                code = target.item(current_row, 0).text()
            except:
                code = "text"
            try:
                if oneFile == False:
                    index = next(i for i, row in enumerate(
                        save_list) if row[1]["machine_code"] == code)
                    save_list.pop(index)  # đang có lỗi ở đây, cần kiểm tra lại
                else:
                    index = next(i for i, row in enumerate(
                        save_list) if row[1]["machine_code"] == code)
                    save_list[index][1] = "text"
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Error: {e}")
        target.removeRow(current_row)

    @QtCore.pyqtSlot()
    def print_record(self):  # đang có lỗi ở đây, cần kiểm tra lại sẽ xuất hiện tình trạng không insert vào db máy đính kèm nhưng vẫn in được form
        role, dep = self.login_info['role_level'], self.login_info['department']
        if role not in ["Admin", "Manager"] and dep != self.department_print_record:
            QtWidgets.QMessageBox.information(
                self, "Permission denied", "You don't have permission to print maintenance record")
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
                SELECT m.machine_code AS attach_machine, m2.machine_code AS main_machine
                FROM Record_pending AS rp
                JOIN Machines AS m ON rp.machine_id = m.machine_id
                JOIN Machines AS m2 ON rp.attached_equipment = m2.machine_id
            ''')
            self.record_pending = [r[0] for r in res_pending]
            exist_map = {r[0]: r[1] for r in res_attach_exist}
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Database load failed: {e}")
            return
        attached_machine= self.result_print_record.set_index('machine_code')['attached_equipment'].to_dict()
        print(attached_machine)
        return
        for r in range(rows):
            for c in range(cols):
                if c == 5:
                    continue
                # elif c == 2:
                #     editor = self.ui.print_record_table.cellWidget(r, 0)
                #     main_code = editor.text().strip()
                #     text = self.ui.print_record_table.model().index(r, c).data(QtCore.Qt.EditRole)
                #     if text is None or text == "":
                #         continue
                #     attached_machine[main_code] = [p.strip()
                #                                    for p in text.split(';') if p and p.strip()]
                elif c == 7:
                    try:
                        editor = self.ui.print_record_table.cellWidget(r, c)
                        text = editor.text().strip()
                        if not STRICT_DATE.match(text):
                            raise ValueError("Format must be YYYY-MM-DD")
                        dt.datetime.strptime(text, "%Y-%m-%d")
                    except ValueError:
                        QtWidgets.QMessageBox.critical(
                            self, "Error", f"Please enter the correct date format (YYYY-MM-DD) at row {r+1}, column {c+1}")
                        return
                else:
                    try:
                        editor = self.ui.print_record_table.cellWidget(r, c)
                        text = editor.text().strip()
                        if editor is None or text == "" or text == "None":
                            QtWidgets.QMessageBox.critical(
                                self, "Error", f"Please fill in all blanks or cells with 'None' content")
                            return
                    except Exception as e:
                        QtWidgets.QMessageBox.critical(
                            self, "Error", f"Error at row {r+1}, column {c+1}: {e}")
                        return
            if (main_code in self.record_pending) and (main_code not in exist_map.keys()) and (main_code not in exist_map.values()):
                reply = QtWidgets.QMessageBox.question(
                    self, "Record Exists",
                    f"Record for {main_code} already printed.\nDo you want to print again?",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                    QtWidgets.QMessageBox.No
                )
                if reply == QtWidgets.QMessageBox.No:
                    self.result_print_record = [
                        rec for rec in self.result_print_record if rec[0] != main_code]
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
                            'code': main_code}
                    )
                    self.duplicate.append(r)
                except Exception as e:
                    QtWidgets.QMessageBox.critical(
                        self, "Error", f"Update failed: {e}")
                    return
            elif (main_code in self.record_pending) and ((main_code in exist_map.keys()) or (main_code in exist_map.values())):
                partner = exist_map.get(main_code)
                if partner is None:
                    partner = next(
                        (k for k, v in exist_map.items() if v == main_code), None)

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
            attach_codes = [c for lst in attached_machine.values()
                            for c in lst]
            dup = [c for c in set(attach_codes) if attach_codes.count(c) > 1]
            if dup:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Duplicated attached equipment: {', '.join(dup)}")
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
                    AND m.machine_code IN ({code_list});
                '''
                res_valid = self.database_process.query(
                    query, params={'year': self.year_num, 'dep': dep})
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
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Validation failed: {e}")
            return
        self.print_selector = Print_selector(self, quantity=len(self.result_print_record), data=self.result_print_record,
                                             attached_machine=attached_machine, database=self.database_process, duplicate=self.duplicate)
        self.print_selector.show()

    @QtCore.pyqtSlot()
    def register_new_form(self):
        if (not hasattr(self, "form_modification")) or sip.isdeleted(self.form_modification):
            self.form_modification = Form_Modification(parent=self)
        self.form_modification.register_form_page()
        self.form_modification.show()
        self.form_modification.ui.stackedWidget.setCurrentWidget(
            self.form_modification.ui.register_form_page)

    @QtCore.pyqtSlot()
    def update_register_form(self):
        if (not hasattr(self, "form_modification")) or sip.isdeleted(self.form_modification):
            self.form_modification = Form_Modification(parent=self)
        self.form_modification.update_form_page()
        self.form_modification.show()
        self.form_modification.ui.stackedWidget.setCurrentWidget(
            self.form_modification.ui.update_form_page)

    @QtCore.pyqtSlot()
    def clear_print_data(self):
        self.ui.print_record_table.clearContents()
        self.ui.print_record_table.setRowCount(0)
        self.result_print_record.drop(self.result_print_record.index, inplace=True)

    @QtCore.pyqtSlot()
    def Mainten_Input_page(self):
        self.style_button_with_shadow(
            (self.ui.Main_Input_record_btn, self.ui.Main_detail_plan_btn, self.ui.Main_Home_btn, self.ui.Main_Print_record_btn))
        self.wrong_scan = []
        try:
            model = self.ui.Main_pending_table.model()
            model.removeRows(0, model.rowCount())
        except AttributeError:
            pass
        self.scan_result_list = None
        self.scan_QRcode_link = None
        self.ui.Maintenance_stacked.setCurrentWidget(self.ui.Input_page_M)
        self.ui.Main_save_link.setPlaceholderText(
            r"\\172.30.73.156\nd_ie2\Noise Device - IE Data\DANH MUC THIET BI BAO TRI\2025")
        self.ui.Main_scan_link.setPlaceholderText(
            "Just accept folder or pdf file")
        if self.ui.Main_update_group_cbb.count() <= 0:
            self.ui.Main_update_group_cbb.addItems(
                [""]+[group[0] for group in self.group])
            self.ui.Main_update_group_cbb.setCurrentIndex(0)
        self.data_model = QtGui.QStandardItemModel()
        headers = ["Machine code", "Machine Name", "Group", "Line",
                   "Technical", "Maintenance\ndate", "Attached of"]
        self.data_model.setHorizontalHeaderLabels(headers)
        headers = ["Machine code", "Machine Name", "Group", "Line", "Technical",
                   "Maintenance\ndate", "Next due\ndate", "Attached\nof", "Page\nNumber"]
        self.ui.Main_scan_result_table.setColumnCount(len(headers))
        self.ui.Main_scan_result_table.setHorizontalHeaderLabels(headers)
        self.ui.Main_pending_table.setEditTriggers(
            QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.Main_pending_table.setAlternatingRowColors(True)
        self.ui.Main_scan_result_table.setAlternatingRowColors(True)

        def job():
            try:
                PE1_pending = self.database_process.query(sql=''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name = "PE1"''')
                PE2_pending = self.database_process.query(sql=''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name = "PE2"''')
                PE3_pending = self.database_process.query(sql=''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name = "PE3"''')
                PE5_pending = self.database_process.query(sql=''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name = "PE5"''')
                PE4_pending = self.database_process.query(sql=''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name = "PE4"''')
                ELSE_pending = self.database_process.query(sql=''' SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                            FROM `View_Record_Pending` 
                                                            WHERE department_name NOT IN ("PE1","PE2","PE3","PE5","PE4") ''')
                result_pending = PE1_pending + PE2_pending + \
                    PE3_pending + PE4_pending + PE5_pending + ELSE_pending
                pending_record_dep = {"PE1": {(code[0], code[3], str(code[5])) for code in PE1_pending}, "PE2": {(code[0], code[3], str(code[5])) for code in PE2_pending}, "PE3": {(code[0], code[3], str(code[5])) for code in PE3_pending}, "PE4": {
                    (code[0], code[3], str(code[5])) for code in PE4_pending}, "PE5": {(code[0], code[3], str(code[5])) for code in PE5_pending}, "ELSE": {(code[0], code[3], str(code[5])) for code in ELSE_pending}}
                temp = [code[0] for code in result_pending]
                placeholders = ",".join(f"'{k}'" for k in temp)
                sql = f'''
                    SELECT machine_code,maintenance_frequency
                    FROM machines
                    WHERE machine_code IN ({placeholders})
                '''
                result = self.database_process.query(sql=sql)
                maintenance_frequency_dict = dict(result)
            except Exception as e:
                return {"result_pending": False, "error": e}
            return {"result_pending": result_pending, "pending_record_dep": pending_record_dep, "maintenance_frequency_dict": maintenance_frequency_dict}
        self.worker = WorkerThread(job)
        self.worker.finished.connect(
            lambda data: self.on_update_data_ready(data=data))
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.start()

    @QtCore.pyqtSlot()
    def on_update_data_ready(self, data):
        if not data["result_pending"]:
            error_message = str(data['error'])
            if "')' at line 3" in error_message:
                QtWidgets.QMessageBox.information(
                    self, "Information", "All records are up to date. No pending records found.")
                return
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {error_message}")
            return
        try:
            self.ui.Main_scan_result_table.clearContents()
            self.ui.Main_scan_result_table.setRowCount(0)
            result_pending, self.pending_record_dep, self.maintenance_frequency_dict = (
                data["result_pending"], data["pending_record_dep"], data["maintenance_frequency_dict"])
            self.add_data_to_model(
                data=result_pending, target=self.ui.Main_pending_table, model=self.data_model)
            self.ui.Main_pending_table.setColumnWidth(0, 100)
            self.ui.Main_pending_table.setColumnWidth(1, 200)
            self.ui.Main_pending_table.setColumnWidth(2, 50)
            self.ui.Main_pending_table.setColumnWidth(3, 50)
            self.ui.Main_pending_table.setColumnWidth(4, 60)
            self.ui.Main_pending_table.setColumnWidth(5, 90)
            self.ui.Main_pending_table.setColumnWidth(6, 100)
            self.ui.Main_scan_result_table.setColumnWidth(0, 100)
            self.ui.Main_scan_result_table.setColumnWidth(1, 180)
            self.ui.Main_scan_result_table.setColumnWidth(2, 50)
            self.ui.Main_scan_result_table.setColumnWidth(3, 50)
            self.ui.Main_scan_result_table.setColumnWidth(4, 60)
            self.ui.Main_scan_result_table.setColumnWidth(5, 90)
            self.ui.Main_scan_result_table.setColumnWidth(6, 90)
            self.ui.Main_scan_result_table.setColumnWidth(7, 90)
            self.ui.Main_scan_result_table.setColumnWidth(8, 50)
            self.safe_connect(self.ui.Main_scan_btn.clicked,
                              lambda _: self.scan_record())
            self.safe_connect(self.ui.Main_update_group_cbb.currentIndexChanged,
                              lambda _: self.change_text_pending(dep_changed=True))
            self.safe_connect(self.ui.Main_update_line_cbb.currentIndexChanged,
                              lambda _: self.change_text_pending())
            self.safe_connect(self.ui.Main_Scan_result_insert_btn.clicked,
                              lambda _: self.insert_scan_result_row())
            self.list_of_keys = ["machine_code", "machine_name", "group", "line",
                                 "technical", "maintenance_date", "attached_machine", "page_num"]
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot()
    def scan_record(self):
        self.scan_result_final = []
        try:
            self.ui.Main_scan_result_table.clearContents()
            self.ui.Main_scan_result_table.setRowCount(0)
            link = self.ui.Main_scan_link.text()
            link = link.strip().replace('"', '').replace("'", '')
            if not link:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Please enter the path to the record scan folder.")
                return
            self.scan_QRcode_link = self.scan_QRcode.paths(link)
            self.safe_connect(self.ui.Main_delete_row_btn.clicked, lambda _: self.delete_row(
                form="update_table", target=self.ui.Main_scan_result_table, save_list=self.scan_QRcode_link, oneFile=self.scan_QRcode.oneFile))
            self.safe_connect(self.ui.Main_update_btn.clicked,
                              lambda _: self.update_record())
            self.safe_connect(
                self.ui.Main_sync_missing_data_btn.clicked, lambda _: self.Sync_missing_data())
            self.scan_worker = Worker_Pool(
                self.scan_job, self.scan_QRcode_link)
            self.progress_window = Printer_progress(
                max=len(self.scan_QRcode_link), text="scanned")
            self.progress_window.ui.label.setText("Scanning...")
            self.scan_worker.signals.progress_changed.connect(
                lambda value: self.progress_window.update_progress(value=value))
            self.scan_worker.signals.finished.connect(
                self.progress_window.on_finished)
            self.scan_worker.signals.result_ready.connect(
                lambda row, scan_result: self.update_table_row(row, scan_result))
            self.scan_worker.signals.error.connect(lambda msg:
                                                   QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {msg}"))
            self.progress_window.show()
            QtCore.QThreadPool.globalInstance().start(self.scan_worker)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to scan data: {e}")
            self.progress_window.close()

    def scan_job(self, paths):
        if os.path.isdir(self.scan_QRcode.link):
            self.scan_result_list = []
            for i, path in enumerate(paths):
                try:
                    scan_result = self.scan_QRcode.scanning_dir(path)
                    try:
                        scan_result = json.loads(scan_result)
                        self.scan_QRcode_link[i] = [
                            self.scan_QRcode_link[i],] + [scan_result]
                        self.scan_result_list.append(scan_result)
                    except json.JSONDecodeError:
                        self.scan_QRcode_link[i] = [
                            self.scan_QRcode_link[i],] + ["text"]
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
                self.scan_result_list = self.scan_QRcode.scanning_oneFile(
                    paths[0])
                for row, item in enumerate(self.scan_result_list):
                    try:
                        scan_result = json.loads(item)
                        temp.append([self.scan_QRcode.link,] + [scan_result])
                        self.scan_worker.signals.result_ready.emit(
                            row, scan_result)
                        self.take_code_and_page.append(
                            scan_result["machine_code"])
                    except json.JSONDecodeError:
                        continue
                self.scan_QRcode_link = temp
            except Exception as e:
                self.scan_worker.signals.error.emit(str(e))

    @QtCore.pyqtSlot()
    def update_table_row(self, row, scan_result):
        table_row = self.insert_row(
            form="update_table", target=self.ui.Main_scan_result_table)
        for col in range(self.ui.Main_scan_result_table.columnCount()):
            try:
                self.add_item_to_scan_result(table_row, col, scan_result)
            except Exception as e:
                if col == 0:
                    current_row = self.ui.Main_scan_result_table.rowCount() - 1
                    editor = QtWidgets.QLineEdit()
                    editor.setStyleSheet(''' border: none;''')
                    self.safe_connect(editor.textChanged,
                                      lambda text, r=current_row, c=0: self.on_text_changed(text=text, c=c, r=r, target=self.ui.Main_scan_result_table,
                                                                                            select_col="machine_code", table="View_Record_Pending ", where=f"WHERE machine_code LIKE '%{text}%'")
                                      )
                    self.safe_connect(
                        editor.editingFinished, lambda r=current_row: self.load_pending_record(row=current_row))
                    self.ui.Main_scan_result_table.setCellWidget(
                        current_row, 0, editor)

    @QtCore.pyqtSlot()
    def change_text_pending(self, dep_changed=False):
        self.data_model.removeRows(0, self.data_model.rowCount())
        sql = '''SELECT machine_code, machine_name, department_name, line_name, technical, maintenance_date, attached_equipment_code
                                                                        FROM `View_Record_Pending`'''
        dep = self.ui.Main_update_group_cbb.currentText()
        line = self.ui.Main_update_line_cbb.currentText()
        if dep_changed:
            line_list = self.database_process.query(
                sql="SELECT DISTINCT line_name FROM `View_Record_Pending` WHERE department_name = :dep", params={'dep': dep})
            self.ui.Main_update_line_cbb.clear()
            self.ui.Main_update_line_cbb.addItem("")
            self.ui.Main_update_line_cbb.addItems(
                [line[0] for line in line_list])
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
            result = self.database_process.query(
                sql=sql, params={'dep': dep, 'line': line})
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")
        self.add_data_to_model(
            result, self.ui.Main_pending_table, self.data_model)

    @QtCore.pyqtSlot()
    def update_record(self):  # đang có lỗi ở đây, cần kiểm tra lại: xuất hiện tình trạng đã insert vào db và delete record pending nhưng ko có file đc lưu vào folder
        if self.login_info['role_level'] not in ["Secretary", "Manager", "Admin"]:
            QtWidgets.QMessageBox.information(
                self, "Permission denied", "Your don't have permission update maintenance record")
            return
        update_list = []
        machine_update_status = []
        save_link = self.ui.Main_save_link.text()
        jobs = []
        record_link_dict = {}
        if not os.path.exists(save_link):
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Save link not exists")
            return

        def save_pdf(path, save_link):
            for index in range(len(path)):
                if (path[index][1] != "text") and (path[index][1]["machine_code"] not in self.wrong_scan):
                    machine_info_list = [path[index][1]["machine_code"], path[index][1]["machine_name"], path[index][1]
                                         ["group"], path[index][1]["line"], path[index][1]["technical"], path[index][1]["maintenance_date"]]
                    file_list = [[path[index][1]["machine_code"], "_".join(
                        x.strip().replace(" ", "_") for x in machine_info_list) + ".pdf"]]
                    if path[index][1]["attached_machine"] != "" and path[index][1]["attached_machine"] != []:
                        attached_name = self.database_process.query(
                            f'''SELECT machine_name FROM `Machines` WHERE machine_code IN ({",".join([f"'{c}'" for c in path[index][1]["attached_machine"]])});''')
                        for index_1, item in enumerate(path[index][1]["attached_machine"]):
                            attached_file = [
                                item, attached_name[index_1][0]] + machine_info_list[2:]
                            file_list.append([item, "_".join(x.strip().replace(
                                " ", "_") for x in attached_file) + ".pdf"])
                    for file in file_list:
                        record_link_dict[file[0]
                                         ] = f"{save_link}/{file[1].replace('/', '-')}"
                        if os.path.isdir(self.scan_QRcode.link):
                            pass
                        else:
                            try:
                                result = self.database_process.query(
                                    '''SELECT machine_code, page_num FROM `maintenance_form_info`;''')
                                self.page_num_dict = dict(result)
                                pages = int(self.page_num_dict.get(
                                    path[index][1]["machine_code"], 1))
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
                                QtWidgets.QMessageBox.critical(
                                    self, "Error", f"Failed to load data scan: {e}")
                else:
                    continue
        save_pdf(path=self.scan_QRcode_link, save_link=save_link)
        try:
            for row in range(self.ui.Main_scan_result_table.rowCount()):
                if self.ui.Main_scan_result_table.item(row, 0).text() in self.wrong_scan:
                    continue
                else:
                    machine_code = self.ui.Main_scan_result_table.item(
                        row, 0).text()
                    line_name = self.ui.Main_scan_result_table.item(
                        row, 3).text()
                    technical = self.ui.Main_scan_result_table.item(
                        row, 4).text()
                    maintenance_date = self.ui.Main_scan_result_table.item(
                        row, 5).text()
                    next_date = self.ui.Main_scan_result_table.item(
                        row, 6).text()
                    record_link = record_link_dict[machine_code]
                    update_list.append({'code': machine_code,
                                        'maintenance_date': maintenance_date,
                                        'technical': technical,
                                        'next_date': next_date,
                                        'line': line_name,
                                        'link': record_link})
                    machine_update_status.append({'code': machine_code,
                                                  'line': line_name,
                                                  'status': 'GOOD'})
            success = self.database_process.executemany(sql=f'''INSERT INTO `maintenance_records` ( machine_id, maintenance_date, technician, `next_due_date`, line_id, record_link )
                                                SELECT m.machine_id,
                                                    :maintenance_date,
                                                    :technical,
                                                    :next_date,
                                                    p.line_id,
                                                    :link
                                                FROM `machines` m
                                                JOIN `production_Lines` p ON p.line_name = :line
                                                WHERE m.machine_code = :code;''', params_list=update_list).rowcount
            # success = 1
            if success > 0:
                QtWidgets.QMessageBox.information(
                    self, "Success", f"Đã cập nhật {success.rowcount} bản ghi.")
                self.database_process.executemany(sql='''UPDATE `machines` AS m
                                                            JOIN `production_lines` AS p ON p.line_name = :line
                                                            SET m.machine_status = :status, m.line_id = p.line_id
                                                            WHERE m.machine_code = :code;''', params_list=machine_update_status)
                for job in jobs:
                    job()
            else:
                QtWidgets.QMessageBox.warning(
                    self, "No Change", "Không có bản ghi nào được cập nhật.")
            self.ui.Main_scan_result_table.clearContents()
            self.ui.Main_scan_result_table.setRowCount(0)
            self.scan_QRcode.link = ""
            self.change_text_pending()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    def add_item_to_scan_result(self, row, col, scan_result):
        if col < 6:
            item = QtWidgets.QTableWidgetItem(
                str(scan_result[self.list_of_keys[col]]))
            if (col == 0) and ((scan_result[self.list_of_keys[0]], scan_result[self.list_of_keys[3]], scan_result[self.list_of_keys[5]]) not in self.pending_record_dep.get(f"{scan_result[self.list_of_keys[2]]}", self.pending_record_dep["ELSE"])):
                item.setBackground(QtGui.QColor(255, 80, 80))
                self.wrong_scan.append(
                    str(scan_result[self.list_of_keys[col]]))
                if scan_result[self.list_of_keys[-2]] != "" and scan_result[self.list_of_keys[-2]] != []:
                    self.wrong_scan + scan_result[self.list_of_keys[-2]]
            item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.ui.Main_scan_result_table.setItem(row, col, item)
        elif col == 6:
            date_object = dt.datetime.strptime(
                self.ui.Main_scan_result_table.item(row, col-1).text(), '%Y-%m-%d').date()
            next_date = date_object + relativedelta(months=int(
                self.maintenance_frequency_dict[self.ui.Main_scan_result_table.item(row, 0).text()]))
            item = QtWidgets.QTableWidgetItem(next_date.strftime("%Y-%m-%d"))
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.ui.Main_scan_result_table.setItem(row, col, item)
        elif col == 7:
            if scan_result[self.list_of_keys[-2]] == "" or scan_result[self.list_of_keys[-2]] == []:
                return
            else:
                for item in scan_result[self.list_of_keys[-2]]:
                    self.copy_row(row, item)
        else:
            item = QtWidgets.QTableWidgetItem(
                str(scan_result[self.list_of_keys[-1]]))
            item.setFlags(item.flags() | QtCore.Qt.ItemIsEditable)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.ui.Main_scan_result_table.setItem(row, col, item)

    def copy_row(self, row_index, attached_machine):
        row_count = self.ui.Main_scan_result_table.rowCount()
        self.ui.Main_scan_result_table.insertRow(row_count)
        for column_index in range(self.ui.Main_scan_result_table.columnCount()):
            src_item = self.ui.Main_scan_result_table.item(
                row_index, column_index)
            if src_item is not None:
                if column_index == 0:
                    new_item = QtWidgets.QTableWidgetItem(
                        str(attached_machine))
                else:
                    new_item = QtWidgets.QTableWidgetItem(src_item.text())
                new_item.setBackground(src_item.background())
                new_item.setForeground(src_item.foreground())
                new_item.setFont(src_item.font())
            else:
                if column_index == 7:
                    text0 = self.ui.Main_scan_result_table.item(row_index, 0)
                    new_item = QtWidgets.QTableWidgetItem(
                        text0.text() if text0 is not None else "")
                    if text0 is not None:
                        new_item.setBackground(text0.background())
                        new_item.setForeground(text0.foreground())
                        new_item.setFont(text0.font())
                else:
                    new_item = QtWidgets.QTableWidgetItem("")
            new_item.setFlags(new_item.flags() & ~QtCore.Qt.ItemIsEditable)
            new_item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.ui.Main_scan_result_table.setItem(
                row_count, column_index, new_item)

    @QtCore.pyqtSlot()
    def insert_scan_result_row(self):
        if self.scan_QRcode.link == "" or self.scan_QRcode.link is None:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Please scan QR code first.")
            return
        row_count = self.ui.Main_scan_result_table.rowCount()
        self.ui.Main_scan_result_table.insertRow(row_count)
        for column_index in range(self.ui.Main_scan_result_table.columnCount()):
            if column_index != 0:
                new_item = QtWidgets.QTableWidgetItem("")
                new_item.setFlags(new_item.flags() & ~QtCore.Qt.ItemIsEditable)
                new_item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.ui.Main_scan_result_table.setItem(
                    row_count, column_index, new_item)
            else:
                editor = QtWidgets.QLineEdit()
                editor.setStyleSheet(''' border: none;''')
                self.safe_connect(editor.textChanged,
                                  lambda text, r=row_count, c=0: self.on_text_changed(text=text, c=c, r=r, target=self.ui.Main_scan_result_table,
                                                                                      select_col="machine_code", table="View_Record_Pending ", where=f"WHERE machine_code LIKE '%{text}%'")
                                  )
                self.safe_connect(
                    editor.editingFinished, lambda r=row_count: self.load_pending_record(row=r))
                self.ui.Main_scan_result_table.setCellWidget(
                    row_count, 0, editor)

    @QtCore.pyqtSlot()
    def load_pending_record(self, row):
        code = self.ui.Main_scan_result_table.cellWidget(row, 0).text()
        for item in self.scan_QRcode_link:
            if item[1]["machine_code"] == code:
                return
        try:
            result = self.database_process.query(
                sql=''' SELECT * FROM  `View_Record_Pending` WHERE machine_code = :code''', params={'code': code})
            if not result:
                QtWidgets.QMessageBox.warning(
                    self, "Not found", f"Machine code '{code}' not found in pending records.")
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
            self.scan_QRcode_link.append(
                [self.scan_QRcode.path[0], scan_result_dict])
            for col in range(self.ui.Main_scan_result_table.columnCount()):
                self.add_item_to_scan_result(row, col, scan_result_dict)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot()
    def Sync_missing_data(self):
        if not self.scan_QRcode_link:
            QtWidgets.QMessageBox.information(
                self, "Info", "No scanned data to sync.")
            return
        existing_data = []
        line_name = self.scan_QRcode_link[0][1]["line"]
        for item in self.scan_QRcode_link:
            if item[1] == "text":
                continue
            if item[1]["line"] != line_name:
                QtWidgets.QMessageBox.critical(
                    self, "Error", "Scanned data contains multiple line names. Please scan data from the same line to sync missing records.")
                return
            existing_data.append(item[1]["machine_code"])
            if item[1]["attached_machine"] != "" and item[1]["attached_machine"] != []:
                existing_data += item[1]["attached_machine"]
        existing_data = list(set(existing_data))
        try:
            temp = self.database_process.query(sql=f''' SELECT * FROM `View_Record_Pending`
                                                            WHERE line_name = :line 
                                                                AND machine_code NOT IN ({",".join([f"'{code}'" for code in existing_data])});''',
                                               params={'line': line_name})
            self.sync_missing_list = {}
            for record in temp:
                self.sync_missing_list[record[0]] = {
                    "machine_code": record[0],
                    "machine_name": record[1],
                    "group": record[2],
                    "line": record[3],
                    "technical": record[4],
                    "maintenance_date": str(record[5]),
                    "attached_machine": record[6].split(",") if record[6] else "",
                    "page_num": None
                }
            self.sync_window = Sync_Missing_Data(parent=self, line_name=line_name, data_list=[
                                                 [data["machine_code"], data["page_num"]] for key, data in self.sync_missing_list.items()])
            self.sync_window.synced.connect(self.on_missing_data_synced)
            self.sync_window.show()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to sync data: {e}")

    @QtCore.pyqtSlot()
    def on_missing_data_synced(self):
        if not self.sync_missing_list:
            return
        for code, data in self.sync_missing_list.items():
            if data["page_num"] is not None:
                row_count = self.ui.Main_scan_result_table.rowCount()
                self.ui.Main_scan_result_table.insertRow(row_count)
                for col in range(self.ui.Main_scan_result_table.columnCount()):
                    self.add_item_to_scan_result(row_count, col, data)
                self.scan_QRcode_link.append([self.scan_QRcode.path[0], data])

    @QtCore.pyqtSlot()
    def Mainten_Detail_plan_page(self):
        self.style_button_with_shadow(
            (self.ui.Main_detail_plan_btn, self.ui.Main_Input_record_btn, self.ui.Main_Home_btn, self.ui.Main_Print_record_btn))
        self.ui.Maintenance_stacked.setCurrentWidget(
            self.ui.Detail_plan_page_M)
        self.itemChanged = {"update": set(), "insert": set()}
        self.department_maintenance_plan = None
        if self.ui.Group_cbb_DP.count() <= 0:
            self.ui.Year_plan_cbb_DP.clear()
            self.ui.Year_plan_cbb_DP.addItems(
                [str(y) for y in range(2025, 2035)])
            self.ui.Year_plan_cbb_DP.setCurrentText(str(self.year_num))
            headers = ["Machine\ncode", "Machine Name", "Group"]
            self.ui.Detail_plan_frze_table.setColumnCount(len(headers))
            self.ui.Detail_plan_frze_table.setHorizontalHeaderLabels(headers)
            self.ui.Detail_plan_frze_table.setColumnWidth(0, 80)
            self.ui.Detail_plan_frze_table.setColumnWidth(1, 240)
            self.ui.Detail_plan_frze_table.setColumnWidth(2, 50)
            self.ui.Detail_plan_frze_table.setVerticalScrollBarPolicy(
                QtCore.Qt.ScrollBarAlwaysOff)
            hearders = ["Week\n(M1)", "Status\n(M1)", "Line\n(M1)", "Record\n(M1)", "Week\n(M2)", "Status\n(M2)", "Line\n(M2)", "Record\n(M2)", "Week\n(M3)", "Status\n(M3)", "Line\n(M3)", "Record\n(M3)", "Week\n(M4)", "Status\n(M4)", "Line\n(M4)", "Record\n(M4)",
                        "Week\n(M5)", "Status\n(M5)", "Line\n(M5)", "Record\n(M5)", "Week\n(M6)", "Status\n(M6)", "Line\n(M6)", "Record\n(M6)", "Week\n(M7)", "Status\n(M7)", "Line\n(M7)", "Record\n(M7)", "Week\n(M8)", "Status\n(M8)", "Line\n(M8)", "Record\n(M8)",
                        "Week\n(M9)", "Status\n(M9)", "Line\n(M9)", "Record\n(M9)", "Week\n(M10)", "Status\n(M10)", "Line\n(M10)", "Record\n(M10)", "Week\n(M11)", "Status\n(M11)", "Line\n(M11)", "Record\n(M11)", "Week\n(M12)", "Status\n(M12)", "Line\n(M12)", "Record\n(M12)"]
            self.ui.Detail_plan_table.setColumnCount(len(hearders))
            self.ui.Detail_plan_table.setHorizontalHeaderLabels(hearders)
            self.ui.Detail_plan_table.verticalHeader().setVisible(False)
            for i in range(self.ui.Detail_plan_table.columnCount()):
                self.ui.Detail_plan_table.setColumnWidth(i, 59)
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
            self.group_cbb_DP_change()
            self.safe_connect(self.ui.Group_cbb_DP.currentIndexChanged,
                              lambda _: self.group_cbb_DP_change())
            self.safe_connect(self.ui.Code_lnedit_DP.textChanged, lambda text: self.filter_suggestion(target=self.ui.Code_lnedit_DP,
                                                                                                      text="DISTINCT ( m.machine_code )", table="`Maintenance_plan` as mp",
                                                                                                      where=f""" JOIN `Machines` as m
                                                                                                        ON mp.machine_id = m.machine_id
                                                                                                        JOIN `Production_Lines` as p
                                                                                                        ON mp.line_id = p.line_id
                                                                                                        JOIN `Months_Years` as my ON my.month_year_id = mp.month_year_id
                                                                                                        WHERE p.line_name = '{self.ui.Line_cbb_DP.currentText()}' AND m.machine_code LIKE '%{text}%' AND my.year = {self.year_num} """))
            self.safe_connect(self.ui.Load_btn_DP.clicked,
                              lambda _: self.Load_Maintenance_plan())
            self.safe_connect(self.ui.Update_btn_DP.clicked,
                              lambda _: self.Update_maintenance_plan())
            self.safe_connect(self.ui.Delete_btn_DP.clicked,
                              lambda _: self.Delete_plan())
            self.safe_connect(self.ui.Insert_btn_DP.clicked,
                              lambda _: self.Insert_plan())

    @QtCore.pyqtSlot()
    def group_cbb_DP_change(self):
        dep = self.ui.Group_cbb_DP.currentText()
        try:
            lines = self.database_process.query(sql='''SELECT DISTINCT( p.line_name )
                                                        FROM  `Maintenance_plan` as mp
                                                        JOIN `Production_Lines` as p
                                                        ON mp.line_id = p.line_id
                                                        JOIN `Departments` as d
                                                        ON p.department_id = d.department_id
                                                        JOIN `Months_Years` as my 
                                                        ON my.month_year_id = mp.month_year_id
                                                        WHERE d.department_name = :dep AND my.year = :year;''', params={'dep': dep, 'year': self.year_num})
            items = [" "] + [line[0] for line in lines]
            self.ui.Line_cbb_DP.clear()
            self.ui.Line_cbb_DP.addItems(items)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot()
    def Load_Maintenance_plan(self):
        self.pdf_windows = []
        self.itemChanged = {"update": set(), "insert": set()}
        self.ui.Detail_plan_table.clearContents()
        self.ui.Detail_plan_table.setRowCount(0)
        self.ui.Detail_plan_frze_table.clearContents()
        self.ui.Detail_plan_frze_table.setRowCount(0)
        self.department_maintenance_plan = self.ui.Group_cbb_DP.currentText()
        column_colors = {
            range(3, 7): (0, 51, 102, 50),
            range(7, 11): (255, 182, 193, 50),
            range(11, 15): (102, 205, 170, 50),
            range(15, 19): (255, 239, 153, 50),
            range(19, 23): (64, 224, 208, 50),
            range(23, 27): (255, 140, 0, 50),
            range(27, 31): (220, 20, 60, 50),
            range(31, 35): (255, 215, 0, 50),
            range(35, 39): (184, 115, 51, 50),
            range(39, 43): (204, 85, 0, 50),
            range(43, 47): (54, 69, 79, 50),
            range(47, 51): (128, 0, 32, 50),
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
        params = {'dep': self.ui.Group_cbb_DP.currentText(
        ), 'year': self.ui.Year_plan_cbb_DP.currentText()}
        if self.ui.Line_cbb_DP.currentText() != " " and self.ui.Line_cbb_DP.currentText() != "":
            script += " AND p.line_name = :line"
            params['line'] = self.ui.Line_cbb_DP.currentText()
        if self.ui.Code_lnedit_DP.text() != "":
            script += f" AND m.machine_code LIKE '%{self.ui.Code_lnedit_DP.text()}%'"
        try:
            result = self.database_process.query(sql=f''' 
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
                                        ORDER BY week_q1 ASC;''', params=params)
            self.ui.Detail_plan_table.setRowCount(len(result))
            self.ui.Detail_plan_frze_table.setRowCount(len(result))
            for row in range(len(result)):
                for col in range(len(result[row])):
                    if result[row][col] is None:
                        item = QtWidgets.QTableWidgetItem("")
                    else:
                        item = QtWidgets.QTableWidgetItem(
                            str(result[row][col]))
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    if col < 3:
                        item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                        self.ui.Detail_plan_frze_table.setItem(row, col, item)
                    else:
                        color = get_color_for_column(col)
                        item.setBackground(QtGui.QColor(
                            color[0], color[1], color[2], color[3]))
                        if result[row][col] == "Overdue":
                            item.setForeground(
                                QtGui.QBrush(QtGui.QColor(255, 0, 0)))
                        elif result[row][col] == "Ontime":
                            item.setForeground(
                                QtGui.QBrush(QtGui.QColor(0, 128, 0)))
                        elif str(result[row][col]).lower().endswith(".pdf"):
                            btn = QtWidgets.QPushButton("")
                            icon = QtGui.QIcon()
                            icon.addFile(resource_path(
                                u"Icons/hyperlink.ico"), QtCore.QSize(), QtGui.QIcon.Normal, QtGui.QIcon.Off)
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
                            self.safe_connect(
                                btn.clicked, lambda _, link=result[row][col].lower(): self.open_pdf(link=link))
                            self.ui.Detail_plan_table.setCellWidget(
                                row, col - 3, btn)
                            for index in range(1, 4):
                                item = self.ui.Detail_plan_table.item(
                                    row, col - 3 - index)
                                item.setFlags(item.flags() & ~
                                              QtCore.Qt.ItemIsEditable)
                            continue
                        self.ui.Detail_plan_table.setItem(row, col - 3, item)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")
        self.ui.Detail_plan_table.itemChanged.connect(
            lambda item: self.on_item_in_Detail_plan_table_change(item=item))

    @QtCore.pyqtSlot()
    def on_item_in_Detail_plan_table_change(self, item):
        row = item.row()
        col = item.column()
        if row not in self.itemChanged["insert"]:
            self.itemChanged["update"].add(row)
        item.setBackground(QtGui.QColor(255, 255, 150))

    @QtCore.pyqtSlot()
    def Update_maintenance_plan(self):
        if self.login_info["role_level"] in ["Manager", "Admin"]:
            pass
        elif (self.login_info["department"] == self.department_maintenance_plan) and (self.login_info["role_level"] in ["Supervisor"]):
            pass
        else:
            QtWidgets.QMessageBox.information(
                self, "Permission denied", "Your don't have permission to update this machine info")
            return
        update_list = []
        insert_list = []
        finish = 0
        year = int(self.ui.Year_plan_cbb_DP.currentText())

        def update_job(type: str, list: list):
            for r in self.itemChanged[type]:
                try:
                    code = self.ui.Detail_plan_frze_table.item(r, 0).text()
                except:
                    code = self.ui.Detail_plan_frze_table.cellWidget(
                        r, 0).text()
                for col_offset in range(0, 12):
                    if not self.ui.Detail_plan_table.item(r, 0+4*col_offset):
                        continue
                    else:
                        line = self.ui.Detail_plan_table.item(
                            r, 2+4*col_offset).text()
                        week = self.ui.Detail_plan_table.item(
                            r, 0+4*col_offset).text()
                        try:
                            status = self.ui.Detail_plan_table.item(
                                r, 1+4*col_offset).text()
                        except:
                            status = ""
                        if week == "":
                            new_month = ""
                            quarter = ""
                        else:
                            new_month = self.ui.company_week_month(
                                year, int(week))
                            quarter = (new_month - 1) // 3 + 1
                        old_month = col_offset + 1
                        list.append({'code': code, 'line': line, 'quarter': quarter, 'new_month': new_month,
                                    'year': year, 'week': week, 'old_month': old_month, 'status': status})
        if self.itemChanged["update"]:
            update_job(type="update", list=update_list)
            delete_month = []
            update_list_month = []
            for item in update_list:
                if item['week'] == "" and item['line'] == "":
                    delete_month.append(
                        {'code': item['code'], 'del_month': item['old_month'], 'year': item['year']})
                else:
                    update_list_month.append(item)
            try:
                finish += self.database_process.executemany(sql=''' DELETE mp
                                                                FROM `Maintenance_plan` AS mp
                                                                JOIN `Machines` AS m
                                                                ON mp.machine_id = m.machine_id
                                                                JOIN `Months_Years` AS my
                                                                ON mp.month_year_id = my.month_year_id
                                                                WHERE m.machine_code = :code AND my.month = :del_month AND my.year = :year; ''', params_list=delete_month).rowcount
                check_sql = '''
                            SELECT 1
                            FROM `Maintenance_plan` AS mp
                            JOIN `Machines` AS m ON mp.machine_id = m.machine_id
                            JOIN `Months_Years` AS my ON mp.month_year_id = my.month_year_id
                            WHERE m.machine_code = :code AND my.month = :old_month AND my.year = :year;
                            '''
                for data in update_list_month:
                    exists = self.database_process.query(sql=check_sql, params={
                                                         'code': data['code'], 'old_month': data['old_month'], 'year': data['year']})
                    if len(exists) > 0:
                        finish += self.database_process.query(sql=''' UPDATE Maintenance_plan AS mp 
                                                                            JOIN Machines AS m ON mp.machine_id = m.machine_id 
                                                                            JOIN Months_Years AS my ON mp.month_year_id = my.month_year_id 
                                                                            SET 
                                                                            mp.line_id = ( SELECT p.line_id FROM Production_Lines as p WHERE p.line_name = :line ),
                                                                            mp.month_year_id = ( SELECT my2.month_year_id FROM Months_Years as my2 WHERE my2.month = :new_month AND my2.year = :year ), 
                                                                            mp.week = :week,
                                                                            mp.status = :status
                                                                    WHERE m.machine_code = :code AND my.month = :old_month AND my.year = :year ;''', params=data).rowcount
                        self.database_process.query(sql='''UPDATE machines
                                                            SET machine_status = 'GOOD'
                                                            WHERE machine_code = :code;''', params={'code': data['code']})
                    else:
                        finish += self.database_process.query(sql=''' INSERT INTO `Maintenance_plan` (machine_id,line_id,month_year_id,quarter,week,original_week)
                                                                            SELECT m.machine_id,p.line_id,my.month_year_id,:quarter,:week,:week
                                                                            FROM `Machines` as m
                                                                            JOIN `Production_Lines` as p 
                                                                            ON p.line_name = :line
                                                                            JOIN `Months_Years` as my
                                                                            ON my.month = :new_month AND my.year = :year
                                                                            WHERE m.machine_code = :code; ''', params=data).rowcount
            except Exception as e:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Failed to update data: {e}")
        if self.itemChanged["insert"]:
            update_job(type="insert", list=insert_list)
            try:
                finish += self.database_process.executemany(sql=''' INSERT INTO `Maintenance_plan` (machine_id,line_id,month_year_id,quarter,week,original_week)
                                                                    SELECT m.machine_id,p.line_id,my.month_year_id,:quarter,:week,:week
                                                                    FROM `Machines` as m
                                                                    JOIN `Production_Lines` as p 
                                                                    ON p.line_name = :line
                                                                    JOIN `Months_Years` as my
                                                                    ON my.month = :new_month AND my.year = :year
                                                                    WHERE m.machine_code = :code; ''', params_list=insert_list).rowcount
                for data in insert_list:
                    self.database_process.query(sql='''UPDATE machines
                                                        SET machine_status = 'GOOD'
                                                        WHERE machine_code = :code;''', params={'code': data['code']})

            except Exception as e:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Failed to update data: {e}")
        try:
            self.database_process.query(
                sql="CALL update_maintenance_plan_status;")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to update data: {e}")
            return
        if finish > 0:
            QtWidgets.QMessageBox.information(
                self, "Success", f"Đã cập nhật {finish} bản ghi.")
            return

    @QtCore.pyqtSlot()
    def open_pdf(self, link):
        pdf = pdf_view(link)
        self.pdf_windows.append(pdf)
        pdf.show()

    @QtCore.pyqtSlot()
    def Delete_plan(self):
        if self.login_info["role_level"] in ["Manager", "Admin"]:
            pass
        elif (self.login_info["department"] == self.department_maintenance_plan) and (self.login_info["role_level"] in ["Supervisor"]):
            pass
        else:
            QtWidgets.QMessageBox.information(
                self, "Permission denied", "Your don't have permission to update this machine info")
            return
        current_row = self.ui.Detail_plan_frze_table.currentRow()
        code_item = self.ui.Detail_plan_frze_table.item(current_row, 0)
        if code_item is None:
            self.ui.Detail_plan_table.removeRow(current_row)
            self.ui.Detail_plan_frze_table.removeRow(current_row)
            return
        code = code_item.text()
        question = QtWidgets.QMessageBox.question(
            self, "Delete", f"Are you sure to delete the maintenance plan for the machine '{code}'?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
        if question == QtWidgets.QMessageBox.Yes:
            try:
                self.database_process.query(sql='''   DELETE mp
                                                        FROM `Maintenance_plan` as mp 
                                                        JOIN `Machines` AS m 
                                                        ON mp.machine_id = m.machine_id
                                                        JOIN `Months_Years` AS my
                                                        ON mp.month_year_id = my.month_year_id
                                                        WHERE m.machine_code = :code AND my.year = :year ; ''', params={'code': code, 'year': self.year_num})
                self.ui.Detail_plan_table.removeRow(current_row)
                self.ui.Detail_plan_frze_table.removeRow(current_row)
                self.itemChanged["update"] = {
                    r - 1 if r > current_row else r for r in self.itemChanged["update"] if r != current_row}
                self.itemChanged["insert"] = {
                    r - 1 if r > current_row else r for r in self.itemChanged["insert"] if r != current_row}
            except Exception as e:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot()
    def Insert_plan(self):
        if self.login_info["role_level"] in ["Manager", "Admin"]:
            pass
        elif (self.login_info["department"] == self.department_maintenance_plan) and (self.login_info["role_level"] in ["Supervisor"]):
            pass
        else:
            QtWidgets.QMessageBox.information(
                self, "Permission denied", "Your don't have permission to update this machine info")
            return
        row = self.ui.Detail_plan_frze_table.rowCount()
        self.ui.Detail_plan_frze_table.insertRow(row)
        self.ui.Detail_plan_table.insertRow(row)
        editor = QtWidgets.QLineEdit()
        editor.setAlignment(QtCore.Qt.AlignCenter)
        editor.setStyleSheet(''' border: none;''')
        self.safe_connect(editor.textChanged, self.handle_text_changed_DP)
        self.safe_connect(editor.editingFinished,
                          self.handle_editing_finished_DP)
        self.ui.Detail_plan_frze_table.setCellWidget(row, 0, editor)
        self.itemChanged["insert"].add(row)

    @QtCore.pyqtSlot(str)
    def handle_text_changed_DP(self, text):
        editor = self.sender()
        if editor is None:
            return
        table = self.ui.Detail_plan_frze_table
        for r in range(table.rowCount()):
            if table.cellWidget(r, 0) is editor:
                return self.on_text_changed(text=text, c=0, r=r, target=self.ui.Detail_plan_frze_table,
                                            select_col="machine_code", table="Machines ", where=f"WHERE machine_code LIKE '%{text}%'")

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
    def load_machine_detail_plan(self, r):
        item = self.ui.Detail_plan_frze_table.cellWidget(r, 0)
        code = item.text()
        dep = self.ui.Group_cbb_DP.currentText()
        try:
            isCorrectGroup = self.database_process.query(sql='''SELECT m.machine_name,d.department_name 
                                                                FROM `Machines` as m
                                                                JOIN `Production_Lines` as p
                                                                ON m.line_id = p.line_id
                                                                JOIN `Departments` as d
                                                                ON p.department_id = d.department_id
                                                                WHERE d.department_name = :dep AND m.machine_code = :code
                                                                GROUP BY m.machine_id ;''', params={'code': code, 'dep': dep})
            if not isCorrectGroup:
                raise Exception("The machine is not in your group")
            isCodeInPlan = self.database_process.query(sql='''SELECT 1  
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
                                                                GROUP BY m.machine_id ;''', params={'code': code, 'dep': dep, 'year': self.year_num})
            if isCodeInPlan:
                raise Exception("The machine already have plan")
            self.ui.Detail_plan_frze_table.setItem(
                r, 1, QtWidgets.QTableWidgetItem(isCorrectGroup[0][0]))
            self.ui.Detail_plan_frze_table.setItem(
                r, 2, QtWidgets.QTableWidgetItem(isCorrectGroup[0][1]))
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")
            item.clear()

    def send_notification(self, data: dict):
        try:
            clean_data = {k: v for k, v in data.items() if v is not None}

            if 'payload' in clean_data and isinstance(clean_data['payload'], (dict, list)):
                clean_data['payload'] = json.dumps(
                    clean_data['payload'], ensure_ascii=False)

            columns = ', '.join(clean_data.keys())
            placeholders = ','.join([f":{k}" for k in clean_data.keys()])
            sql = f"INSERT INTO `Notifications` ({columns}) VALUES ({placeholders})"

            self.database_process.query(sql=sql, params=clean_data)

        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to send notification: {e}")

# ==========================Function of Maintenance page ==================================================================================END
# ==========================Function of Maintenance page ==================================================================================END
# ==========================Function of Maintenance page ==================================================================================END

# ==========================Function of Part Order page ==================================================================================BEGIN
# ==========================Function of Part Order page ==================================================================================BEGIN
# ==========================Function of Part Order page ==================================================================================BEGIN

    # @QtCore.pyqtSlot()
    # def Part_order_page(self):
    #     self.ui.main_stacked.setCurrentWidget(self.ui.Part_order_page)
    #     self.set_stylesheet_change_page((self.ui.Order_btn, self.ui.OEE_btn, self.ui.Home_btn,
    #                                     self.ui.Maintenance_btn, self.ui.Stock_btn, self.ui.Downtime_btn))
    #     if not self.is_expanded:
    #         self.is_expanded = True
    #         self.expand_windown_animation(self.is_expanded)

# ==========================Function of Part Order page ==================================================================================END
# ==========================Function of Part Order page ==================================================================================END
# ==========================Function of Part Order page ==================================================================================END

# ==========================Function of Stock control page ==================================================================================BEGIN
# ==========================Function of Stock control page ==================================================================================BEGIN
# ==========================Function of Stock control page ==================================================================================BEGIN

    @QtCore.pyqtSlot()
    def Stock_control_page(self):
        if not self.is_expanded:
            self.is_expanded = True
            self.expand_windown_animation(self.is_expanded)
        self.safe_connect(self.ui.update_inventory_btn.clicked,
                          lambda _: self.run_inventory_update())
        self.ui.main_stacked.setCurrentWidget(self.ui.Stock_control_page)
        self.set_stylesheet_change_page((self.ui.Stock_btn, self.ui.OEE_btn, self.ui.Home_btn,
                                        self.ui.Maintenance_btn, 
                                        self.ui.KPI_btn,
                                        self.ui.Downtime_btn))
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
            header = ["Spare part", "Safety\nstock", "Current\nstock", "Stock up\nreminder",
                      "Unit\nprice", "Total\ncost", "Lead\ntime", "Life\ntime", "Last\nrequest", "Group", "Add PO"]
            self.stock_model = QtGui.QStandardItemModel(0, len(header))
            self.stock_model.setHorizontalHeaderLabels(header)
            self.ui.stock_table.setModel(self.stock_model)
            result = self.database_process.query('''SELECT * FROM `Spare_part_View`
                                                    ORDER BY stockup DESC;''')
            self.inventory_update_date = self.database_process.query(
                '''SELECT MAX(update_at) FROM `inventory`;''')[0][0]
            self.ui.inventory_update_date.setText(
                str(self.inventory_update_date))
            self.image_files = {
                item[0]: item[-1]
                for item in result
                if item[-1] is not None
            }
            ImageCache.init(self.ui.stock_table)
            delegate = StockItemDelegate(buttons=("+",))
            self.safe_connect(delegate.clicked, self.on_button_clicked_stock)
            self.ui.stock_table.setItemDelegate(delegate)
            self.add_data_to_stock_model(result)
            self.ui.stock_table.setUpdatesEnabled(True)
            self.ui.stock_table.setMouseTracking(True)
            self.ui.stock_table.viewport().update()
            self.ui.stock_table.viewport().setMouseTracking(True)
            header = self.ui.stock_table.horizontalHeader()
            self.ui.stock_table.setColumnWidth(0, 600)
            for col in range(1, 11):
                header.setSectionResizeMode(col, QtWidgets.QHeaderView.Stretch)
            self.ui.stock_table.horizontalHeader().setStyleSheet(
                "QHeaderView::section { qproperty-alignment: AlignCenter; }")
            self.ui.stock_table.setSortingEnabled(True)
            self.ui.stock_table.setAlternatingRowColors(True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot()
    def on_button_clicked_stock(self):
        # model = index.model()
        # row = index.row()
        QtWidgets.QMessageBox.information(
            self, "Info", "Function to add PO is still under development.")

    @QtCore.pyqtSlot()
    def show_filter_stock(self):
        self.ui.filter_stock_frame.show()
        self.safe_connect(self.ui.apply_stock_btn.clicked,
                          self.filter_process_stock)
        self.safe_connect(self.ui.cancel_stock_btn.clicked,
                          self.hide_filter_stock)
        self.safe_connect(self.ui.code_stock_lnedit.textChanged,
                          lambda text: self.filter_suggestion_stock(text, "code"))
        self.safe_connect(self.ui.name_stock_lnedit.textChanged,
                          lambda text: self.filter_suggestion_stock(text, "name"))

    @QtCore.pyqtSlot()
    def hide_filter_stock(self):
        self.ui.filter_stock_frame.hide()

    @QtCore.pyqtSlot()
    def filter_suggestion_stock(self, text, fill=""):
        if fill == "code":
            if len(text) < 3:
                return
            SCRIPT = '''SELECT part_code FROM spare_part_view WHERE part_code LIKE :text LIMIT 5;'''
            target = self.ui.code_stock_lnedit
        else:
            if len(text) < 3:
                return
            SCRIPT = '''SELECT part_name FROM spare_part_view WHERE part_name LIKE :text LIMIT 5;'''
            target = self.ui.name_stock_lnedit
        suggestions = []
        part_code = []
        try:
            part_code = self.database_process.query(
                SCRIPT, params={"text": f"%{text}%"})
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
                query.append(
                    f'part_code = "{self.ui.code_stock_lnedit.text()}"')
            if self.ui.name_stock_lnedit.text() != "":
                query.append(
                    f'part_name = "{self.ui.name_stock_lnedit.text()}"')
            if self.ui.group_stock_cbb.currentText() != "":
                query.append(
                    f'department_name = "{self.ui.group_stock_cbb.currentText()}"')
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

    def make_item(self, text, align=QtCore.Qt.AlignCenter):
        value = round(text, 0) if isinstance(
            text, float) and text == 0 else text
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
            data = {"image": image_path,
                    "name": f"{result[row][1]}", "code": f"{result[row][0]}"}
            spare_part = QtGui.QStandardItem()
            spare_part.setData(data, QtCore.Qt.UserRole)
            if (float(result[row][4]) > 0):
                code += 1
            value += float(result[row][4])
            cost += float(result[row][6]) / \
                float(self.exchange_rates[result[row][11]])
            row_items = [
                spare_part,
                self.make_item(float(result[row][2])),  # safety stock
                self.make_item(float(result[row][3])),  # current stock
                self.make_item(float(result[row][4])),  # stock up
                self.make_item(round(float(
                    # unit cost
                    result[row][5]) / float(self.exchange_rates[result[row][11]]), 2)),
                self.make_item(round(float(
                    # total cost
                    result[row][6]) / float(self.exchange_rates[result[row][11]]), 2)),
                self.make_item(result[row][7]),  # lead time
                self.make_item(result[row][8]),  # life time
                self.make_item(result[row][9]),  # last request
                self.make_item(result[row][10])  # department name
            ]
            self.stock_model.appendRow(row_items)
        self.ui.total_part_num.setText(f"{len(result)}")
        self.ui.code_need_order_num.setText(f"{int(code)}")
        self.ui.quantity_order_num.setText(f"{int(value)}")
        self.ui.total_cost_num.setText(f"${round(cost, 2)}")
        self.ui.stock_table.resizeRowsToContents()

    def call_inventory_update(self):
        url = os.getenv("API_UPDATE_INVENTORY")
        try:
            response = requests.get(url, timeout=180)
            return response.json()
        except Exception as e:
            return {"status": "error", "detail": str(e)}

    @QtCore.pyqtSlot()
    def run_inventory_update(self):
        QtWidgets.QMessageBox.information(
            self, "Info", "Inventory update on the client side is under development, and can be auto-updated from the server side for now.")
        return
        try:
            isLocked = self.database_process.query(
                '''SELECT GET_LOCK('inventory_update', 3);''')[0][0]
            if isLocked != 1:
                return QtWidgets.QMessageBox.information(self,
                                                         "In Progress",
                                                         "Another update session is running. Try again after 2 minutes."
                                                         )
            last_update = self.database_process.query(
                '''SELECT MAX(update_at) FROM `inventory`;''')
            last_update = last_update[0][0]
            now = dt.datetime.now()
            diff = now - last_update
            minutes = diff.total_seconds() / 60
            if last_update != self.inventory_update_date and minutes > 3:
                return self._start_inventory_update()

            if last_update == self.inventory_update_date and minutes > 3:
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
            self.database_process.query(
                '''SELECT RELEASE_LOCK('inventory_update');''')
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
                self, "Error", f"Error in update Inventory process: {str(data)}")
        self.database_process.query(
            '''SELECT RELEASE_LOCK('inventory_update');''')

    @QtCore.pyqtSlot()
    def on_inventory_update_error(self, err):
        self.spinner.stop()
        self.setEnabled(True)
        QtWidgets.QMessageBox.warning(
            self, "Error", f"Inventory update failed: {str(err)}")
        self.database_process.query(
            '''SELECT RELEASE_LOCK('inventory_update');''')
# ==========================Function of Stock control page ==================================================================================END
# ==========================Function of Stock control page ==================================================================================END
# ==========================Function of Stock control page ==================================================================================END

    def closeEvent(self, event):
        if hasattr(self, "database_process") and self.database_process:
            try:
                self.database_process.close()
            except Exception as e:
                print("Error closing database_process:", e)

        event.accept()
        QtWidgets.QApplication.quit()

    def draw_circle(self, widget, r, color=(), value=0, font_size=14):
        
        pixmap = QtGui.QPixmap(widget.size())
        pixmap.fill(QtCore.Qt.transparent)         

        painter = QtGui.QPainter(pixmap)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        x = widget.width() / 2 - r / 2
        y = widget.height() / 2 - r / 2
        gradient = QtGui.QConicalGradient(x + r/2, y + r/2, 127)

        gradient.setColorAt(0.0, QtGui.QColor(
            color[0], color[1], color[2], 20))
        gradient.setColorAt(0.25, QtGui.QColor(
            color[0], color[1], color[2], 127))
        gradient.setColorAt(0.5, QtGui.QColor(
            color[0], color[1], color[2], 255))
        gradient.setColorAt(0.75, QtGui.QColor(
            color[0], color[1], color[2], 127))
        gradient.setColorAt(1.0, QtGui.QColor(
            color[0], color[1], color[2], 20))

        pen = QtGui.QPen(QtGui.QBrush(gradient), 20)
        painter.setPen(pen)

        painter.drawEllipse(int(x), int(y), int(r), int(r))
        
        font = QtGui.QFont("Comic Sans MS", font_size, QtGui.QFont.Weight.Bold)
        painter.setFont(font)
        painter.setPen(QtGui.QColor(0, 0, 0))
        painter.drawText(int(x + r/2-25), int(y + r/2 + 5), str(value))
        painter.end()

        widget.setPixmap(pixmap)

    def style_button_with_shadow(self, button: tuple):
        button[0].setStyleSheet('''
                                    QPushButton {
                                                background-color: rgba(0, 0, 0, 0.08);
                                                border: none;
                                                border-radius: 0px;
                                                border-bottom: 2px solid rgba(0, 0, 255, 1);
                                                padding: 5px 15px;
                                                font-weight: bold;
                                    }
        ''')
        for i in range(1, len(button)):
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

# ==========================================================================================================================


# ==========================================================================================================================


# ==========================================================================================================================

    @QtCore.pyqtSlot()
    def Downtime_page(self):
        self.ui.main_stacked.setCurrentWidget(self.ui.Downtime_page)
        self.set_stylesheet_change_page((self.ui.Downtime_btn, self.ui.OEE_btn, self.ui.Home_btn,
                                        self.ui.Maintenance_btn, self.ui.Stock_btn,
                                        self.ui.KPI_btn
                                        ))
        if not self.is_expanded:
            self.is_expanded = True
            self.expand_windown_animation(self.is_expanded)
        self.Dashboard_Downtime_page()
        self.safe_connect(self.ui.DT_dashboard_btn.clicked,
                          self.Dashboard_Downtime_page)
        self.safe_connect(self.ui.DT_data_btn.clicked, self.Data_Downtime_page)
        self.safe_connect(self.ui.DT_import_data_btn.clicked,
                          self.Import_data_Downtime_page)
        self.safe_connect(self.ui.DT_problem_report_btn.clicked,
                          self.Problem_report_Downtime_page)
        self.DT_detail_report_list = []

    @QtCore.pyqtSlot()
    def Dashboard_Downtime_page(self):
        self.style_button_with_shadow(
            (self.ui.DT_dashboard_btn, self.ui.DT_data_btn, self.ui.DT_import_data_btn, self.ui.DT_problem_report_btn))
        self.ui.DT_stacked_widget.setCurrentWidget(self.ui.DT_Dashboard_widget)
        if self.ui.DT_area_cbb.count() > 0:
            return
        else:
            self.safe_connect(self.ui.DT_date_edit_2.dateChanged,
                          lambda: self.DT_filtering(changed_object="date_range"))
        self.DT_calendar_widget = self.ui.DT_date_edit_2.calendarWidget()
        self.table_view_of_DTcalendar = self.DT_calendar_widget.findChild(
            QtWidgets.QTableView)
        try:
            if not hasattr(self, "areas") or not self.areas:
                self.areas = [area[0] for area in self.database_process.query(sql='''SELECT downtime_area_name
                                                                                    FROM `downtime_areas`;''')]
            self.ui.DT_area_cbb.clear()
            self.ui.DT_area_cbb.addItems(self.areas)
            area_name = self.ui.DT_area_cbb.currentText()
            nearest_date = self.database_process.query(
                sql='''SELECT MAX(downtime_date) FROM `downtime_records`;''')[0][0]
            year_exist = self.database_process.query(
                sql='''SELECT DISTINCT YEAR(downtime_date) FROM `downtime_records` ORDER BY YEAR(downtime_date) DESC;''')
            years = [str(year[0]) for year in year_exist]
            self.ui.DT_year_cbb.clear()
            self.ui.DT_year_cbb.addItems(years)
            self.ui.DT_date_edit_2.setDate(
                QtCore.QDate.fromString(str(nearest_date), "yyyy-MM-dd"))
            year = pd.to_datetime(nearest_date).year
            self.Dashboard_Downtime_page_refresh(
                area_name=area_name, target=nearest_date, year=year, view_by="day")
            self.DT_filtered_dict = {
                "area": area_name,
                "view_by": "day",
                "time_range": nearest_date
            }
            self.safe_connect(self.ui.DT_area_cbb.currentTextChanged,
                              lambda: self.DT_Viewby(changed_object="area"))
            self.safe_connect(self.ui.DT_day_radiobtn.clicked,
                              lambda: self.DT_Viewby(changed_object="day"))
            self.safe_connect(self.ui.DT_month_radiobtn.clicked,
                              lambda: self.DT_Viewby(changed_object="month"))
            self.safe_connect(self.ui.DT_year_radiobtn.clicked,
                              lambda: self.DT_Viewby(changed_object="year"))
            self.ui.DT_previous_result.mousePressEvent = lambda event: self.action_text_mouse_press_event(widget=self.ui.DT_previous_result, event=event)
            self.ui.DT_previous_result.setMouseTracking(True)
            self.ui.DT_previous_result.mouseMoveEvent = lambda event: self.action_text_mouse_move_event(widget=self.ui.DT_previous_result, event=event)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")
            return

    def Dashboard_Downtime_page_refresh(self, area_name, target, year, view_by="day"):
        self.style_button_with_shadow((self.ui.DT_detail_chart_line_btn, self.ui.DT_detail_chart_machine_btn,
                                      self.ui.DT_detail_chart_error_btn, self.ui.DT_detail_chart_time_btn))
        self.spinner.start()
        def assign_shift(t):
            hours = t.seconds // 3600
            if 6 <= hours < 14:
                return "Shift 1"
            elif 14 <= hours < 22:
                return "Shift 2"
            else:
                return "Shift 3"
        filter_scripts_report = f""
        if view_by == "day":
            prev_target = (pd.to_datetime(target) - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
            prev_year = prev_target.split("-")[0]
            filter_scripts = "Date = :object"
            filter_scripts_wt = f"lot.operation_date = :object AND YEAR(lot.operation_date) = :year"
            filter_scripts_tg = f"DATE(created_at) <= :object"
            self.ui.frame_94.hide()
        elif view_by == "year":
            prev_year = year - 1
            prev_target = prev_year
            filter_scripts = f"YEAR(Date) = :year"
            filter_scripts_wt = f''' YEAR(lot.operation_date) = :year'''
            filter_scripts_tg = f"YEAR(created_at) = :year"
            self.ui.frame_94.hide()
        elif view_by == "month":
            prev_target = target - 1 if target > 1 else 12
            prev_year = year - 1 if target == 1 else year
            filter_scripts = f"MONTH(Date) = :object AND YEAR(Date) = :year"
            filter_scripts_wt = f''' MONTH(lot.operation_date) = :object AND YEAR(lot.operation_date) = :year'''
            filter_scripts_report = f"AND dact.for_month = :month AND dact.for_year = :year"
            filter_scripts_tg = f"MONTH(created_at) <= :object AND YEAR(created_at) = :year"
            self.ui.frame_94.setHidden(False)
        def fetch_data():
            data = self.database_process.query(sql=f'''SELECT Date ,Start_Time, Start_Repair_Time, End_Time, 
                                                            Total_Loss, Repair_Time, Staff_Name, Error_Code, Machine_Code, Line_Name
                                                        FROM `downtime_report`
                                                        WHERE Downtime_Area = :area_name AND {filter_scripts}
                                                        ORDER BY Date,Start_Time ;''', params={"area_name": area_name, "object": target, "year": year})
            prev_data = self.database_process.query(sql=f'''SELECT sum(Total_Loss) AS Total_Loss, sum(Repair_Time) AS Repair_Time, count(*) AS Downtime_Count
                                                        FROM `downtime_report`
                                                        WHERE Downtime_Area = :area_name AND {filter_scripts}
                                                        ORDER BY Date,Start_Time ;''', params={"area_name": area_name, "object": prev_target, "year": prev_year})
            prev_working_time = self.database_process.query(sql=f'''SELECT pl.line_name, lot.operation_hours, lot.setup_time, lot.break_time
                                                                FROM `line_operation_times` as lot
                                                                JOIN downtime_areas_production_lines as dapl ON lot.line_id = dapl.line_id
                                                                JOIN downtime_areas as da ON dapl.downtime_area_id = da.downtime_area_id
                                                                JOIN production_lines as pl ON lot.line_id = pl.line_id
                                                                WHERE da.downtime_area_name = :area_name AND {filter_scripts_wt};''', params={"area_name": area_name, "object": prev_target, "year": prev_year})
            working_time = self.database_process.query(sql=f'''SELECT pl.line_name, lot.operation_hours , lot.setup_time, lot.break_time
                                                                FROM `line_operation_times` as lot
                                                                JOIN downtime_areas_production_lines as dapl ON lot.line_id = dapl.line_id
                                                                JOIN downtime_areas as da ON dapl.downtime_area_id = da.downtime_area_id
                                                                JOIN production_lines as pl ON lot.line_id = pl.line_id
                                                                WHERE da.downtime_area_name = :area_name AND {filter_scripts_wt};''', params={"area_name": area_name, "object": target, "year": year})
            action_report = self.database_process.query(sql=f'''SELECT dact.action_id, dact.action_content, dact.action_report_link, pl.line_name, m.machine_code, ecl.error_code, dact.for_month, dact.for_year
                                                                    FROM `downtime_actions` as dact
                                                                    JOIN `downtime_areas` as da ON dact.downtime_area_id = da.downtime_area_id
                                                                    LEFT JOIN `production_lines` as pl ON dact.line_id = pl.line_id
                                                                    LEFT JOIN `machines` as m ON dact.machine_id = m.machine_id
                                                                    LEFT JOIN `error_codes_list` as ecl ON dact.error_code = ecl.error_code
                                                                    WHERE da.downtime_area_name = :area_name {filter_scripts_report} AND dact.line_id IS NULL AND dact.machine_id IS NULL AND dact.error_code IS NULL;
                                                                    ''', params={"area_name": area_name, "month": target, "year": year}) if view_by == "month" else []
            prev_action_report = self.database_process.query(sql=f'''SELECT dact.action_id, dact.action_content, dact.action_report_link, pl.line_name, m.machine_code, ecl.error_code, dact.for_month, dact.for_year
                                                                    FROM `downtime_actions` as dact
                                                                    JOIN `downtime_areas` as da ON dact.downtime_area_id = da.downtime_area_id
                                                                    LEFT JOIN `production_lines` as pl ON dact.line_id = pl.line_id
                                                                    LEFT JOIN `machines` as m ON dact.machine_id = m.machine_id
                                                                    LEFT JOIN `error_codes_list` as ecl ON dact.error_code = ecl.error_code
                                                                    WHERE da.downtime_area_name = :area_name {filter_scripts_report};
                                                                    ''', params={"area_name": area_name, "month": prev_target, "year": prev_year}) if view_by == "month" else []
            KPI_data = self.database_process.query(sql= f''' SELECT pl.line_name, m.machine_code, dtg.mttr_target_value, dtg.mtbf_target_value
                                                                FROM `mttr_mtbf_targets` AS dtg
                                                                JOIN `downtime_areas` AS da ON dtg.downtime_area_id = da.downtime_area_id
                                                                LEFT JOIN `production_lines` AS pl ON dtg.line_id = pl.line_id
                                                                LEFT JOIN (SELECT line_id , MAX(created_at) AS max_created_at
                                                                    FROM mttr_mtbf_targets
                                                                    WHERE downtime_area_id = (SELECT downtime_area_id FROM downtime_areas WHERE downtime_area_name = :area_name) and {filter_scripts_tg}
                                                                    GROUP BY line_id
                                                                ) AS latest ON dtg.line_id = latest.line_id AND dtg.created_at = latest.max_created_at
                                                                LEFT JOIN `machines` AS m ON dtg.machine_id = m.machine_id
                                                                LEFT JOIN (SELECT machine_id , MAX(created_at) AS max_created_at
                                                                    FROM mttr_mtbf_targets
                                                                    WHERE downtime_area_id = (SELECT downtime_area_id FROM downtime_areas WHERE downtime_area_name = :area_name) and {filter_scripts_tg}
                                                                    GROUP BY machine_id
                                                                ) AS latest_machine ON dtg.machine_id = latest_machine.machine_id AND dtg.created_at = latest_machine.max_created_at
                                                                WHERE da.downtime_area_name = :area_name;''', params={"area_name": area_name, "object": target, "year": year})
            return {"data": data, "working_time": working_time , "action": action_report, "prev_action" : prev_action_report, "prev_data": prev_data, "prev_working_time": prev_working_time, "KPI_data": KPI_data}
        
        def on_data_fetched_dashboard(res, area_name, target, year, view_by):
            data = res["data"]
            working_time = res["working_time"]
            action_report = res["action"]
            prev_action_report = res["prev_action"]
            prev_data = res["prev_data"]
            prev_working_time = res["prev_working_time"]
            KPI_data = res["KPI_data"]
            self.ui.DT_previous_result.clear()
            action_editor_signals = QtCore.QSignalBlocker(
                self.ui.DT_current_actions)
            self.ui.DT_current_actions.clear()
            if not data and not working_time:
                QtWidgets.QMessageBox.information(
                    self, "No data", "No downtime records found for the selected area and date.")
                self.ui.DTime_value.clear()
                self.ui.DEvent_value.clear()
                self.ui.MTTR_value.clear()
                self.ui.MTBF_value.clear()
                self.ui.DT_table.model().removeRows(0, self.ui.DT_table.model().rowCount())
                self.ui.DT_chart.layout().takeAt(0).widget().deleteLater()
                self.ui.MTTR_chart.layout().takeAt(0).widget().deleteLater()
                self.ui.MTBF_chart.layout().takeAt(0).widget().deleteLater()
                self.spinner.stop()
                return
            self.data = pd.DataFrame(data, columns=["Date", "Downtime Start Time", "Downtime Start Repair Time", "Downtime End Time",
                                     "Total Loss Time", "Repair Time", "Staff Name", "Error Code", "Machine Code", "Line Name"])
            action_report_df = pd.DataFrame(action_report, columns=["Action ID", "Action Content", "Action Report Link", "Line Name", "Machine Code", "Error Code", "For Month", "For Year"])
            action_report_df.fillna("", inplace=True)
            if action_report:
                self.has_DT_comment = True
                self.DT_comment_id = action_report_df["Action ID"].iloc[0]
            else:
                self.has_DT_comment = False
                self.DT_comment_id = None
            prev_action_report = pd.DataFrame(prev_action_report, columns=["Action ID", "Action Content", "Action Report Link", "Line Name", "Machine Code", "Error Code", "For Month", "For Year"])
            prev_action_report.fillna("", inplace=True)
            prev_working_time = pd.DataFrame(prev_working_time, columns=["Line Name", "Working Shift", "Setup Time", "Break Time"])
            prev_working_time["Working Shift"] = prev_working_time["Working Shift"].astype(float)
            prev_working_time["Working Time"] = prev_working_time["Working Shift"]*60 - prev_working_time["Setup Time"] - prev_working_time["Break Time"]
            self.downtime_target = pd.DataFrame(KPI_data, columns=["Line Name","Machine Code" , "MTTR Target", "MTBF Target"])
            mask = self.downtime_target["Line Name"].isna() & self.downtime_target["Machine Code"].isna()
            self.downtime_target.loc[mask, "Line Name"] = "SC-A"
            self.data["Shift"] = self.data["Downtime Start Time"].apply(assign_shift)
            error_code = self.data["Error Code"].unique().tolist()
            error_code = ",".join([f'"{code}"' for code in error_code if code])
            self.error_code_dict = self.database_process.query(sql=f'''SELECT error_code, error_description FROM `error_codes_list`
                                                                    WHERE error_code IN ({error_code})''')
            self.error_code_dict = {code[0]: code[1] for code in self.error_code_dict}
            machine_code = self.data["Machine Code"].unique().tolist()
            machine_code = ",".join([f'"{code}"' for code in machine_code if code])
            self.machine_name_dict = self.database_process.query(sql=f'''SELECT machine_code, machine_name FROM `machines`
                                                                    WHERE machine_code IN ({machine_code})''')
            self.machine_name_dict = {code[0]: code[1] for code in self.machine_name_dict}
            self.working_time = pd.DataFrame(
                working_time, columns=["Line Name", "Working Shift", "Setup Time", "Break Time"])
            self.working_time["Working Shift"] = self.working_time["Working Shift"].astype(float)
            self.working_time["Working Time"] = self.working_time["Working Shift"]*60 - self.working_time["Setup Time"] - self.working_time["Break Time"]
            total_loss = self.data["Total Loss Time"].sum()
            downtime_count = len(self.data)
            mttr_value = self.data["Repair Time"].mean() if downtime_count > 0 else 0
            mttr = self.change_time_format(mttr_value,"m")
            mtbf_value = (self.working_time["Working Time"].sum()-total_loss) / \
                downtime_count if downtime_count > 0 else self.working_time["Working Time"].sum(
            )
            mtbf = self.change_time_format(mtbf_value,"m")
            delta = dt.timedelta(minutes=int(total_loss))
            seconds = int(delta.total_seconds())
            hours = seconds // 3600
            minutes = (seconds % 3600) // 60
            seconds = seconds % 60
            total_loss = f"{hours:02}:{minutes:02}:{seconds:02}"
            self.ui.DTime_value.setText(total_loss)
            self.ui.DEvent_value.setText(str(downtime_count))
            self.ui.MTTR_value.setText(f"{mttr['h']}:{mttr['m']}:{mttr['s']}")
            self.ui.MTBF_value.setText(f"{mtbf['h']}:{mtbf['m']}:{mtbf['s']}")
            prev_total_loss = float(prev_data[0][0]) if prev_data[0][0] is not None else 0
            prev_repair_time = float(prev_data[0][1]) if prev_data[0][1] is not None else 0
            prev_downtime_count = float(prev_data[0][2]) if prev_data[0][2] is not None else 0
            mttr_prev_value = float( (prev_repair_time / prev_downtime_count) if prev_downtime_count > 0 else 0)
            mtbf_prev_value = float(((prev_working_time["Working Time"].sum() - prev_total_loss) / prev_downtime_count) if prev_downtime_count > 0 else prev_working_time["Working Time"].sum())
            mttr_target = self.downtime_target.loc[self.downtime_target["Line Name"] == area_name, "MTTR Target"].iloc[0]
            mtbf_target = self.downtime_target.loc[self.downtime_target["Line Name"] == area_name, "MTBF Target"].iloc[0]
            self.DT_chart_current_group = None
            self.DT_detail_chart_drawing(
                area_name = area_name, group_col="Line Name", value_col="Repair Time", data=self.data, title="Downtime By Line", target=target, year=year, view_by=view_by, mttr_value=mttr_value, mtbf_value=mtbf_value, mttr_target=mttr_target, mtbf_target=mtbf_target , target_df=self.downtime_target)
            self.DT_KPI_chart(widget=self.ui.MTTR_chart, value=mttr_value, target_value=mttr_target, previous_value=mttr_prev_value, label="MTTR")
            self.DT_KPI_chart(widget=self.ui.MTBF_chart, value=float(mtbf_value), target_value=mtbf_target, previous_value=mtbf_prev_value, label="MTBF")
            self.ui.DT_current_actions.setPlainText(action_report_df["Action Content"].iloc[0] if not action_report_df.empty else "")
            self.DT_action_report_show(prev_action_report, widget = self.ui.DT_previous_result)
            self.safe_connect(self.ui.DT_detail_chart_line_btn.clicked, lambda: self.DT_detail_chart_drawing(
                area_name = area_name, group_col="Line Name", value_col="Repair Time", data=self.data, title="Downtime By Line", target=target, year=year, view_by=view_by, mttr_value=mttr_value, mtbf_value=mtbf_value, mttr_target=mttr_target, mtbf_target=mtbf_target , target_df=self.downtime_target))
            self.safe_connect(self.ui.DT_detail_chart_machine_btn.clicked, lambda: self.DT_detail_chart_drawing(
                area_name = area_name, group_col="Machine Code", value_col="Repair Time", data=self.data, title="Downtime By Machine", target=target, year=year, view_by=view_by, mttr_value=mttr_value, mtbf_value=mtbf_value, mttr_target=mttr_target, mtbf_target=mtbf_target, target_df=self.downtime_target))
            self.safe_connect(self.ui.DT_detail_chart_error_btn.clicked, lambda: self.DT_detail_chart_drawing(
                area_name = area_name, group_col="Error Code", value_col="Repair Time", data=self.data, title="Downtime By Error Code", target=target, year=year, view_by=view_by, mttr_value=mttr_value, mtbf_value=mtbf_value, mttr_target=mttr_target, mtbf_target=mtbf_target, target_df=self.downtime_target))
            self.safe_connect(self.ui.DT_detail_chart_time_btn.clicked, lambda: self.DT_detail_chart_drawing(
                area_name = area_name, group_col="Downtime Start Time", value_col="Repair Time", data=self.data, title="Downtime By Time", target=target, year=year, view_by=view_by))
            self.spinner.stop()
            self.safe_connect(self.ui.DT_current_actions.textChanged, lambda: self.action_editing_finished(widget=self.ui.DT_current_actions,btn = self.ui.DT_action_save_btn, data = {"has_OEE_comment": self.has_DT_comment if hasattr(self, 'has_DT_comment') else False, 
                                                                                                                                                                                        "OEE_comment_id": self.DT_comment_id,
                                                                                                                                                                                        "area_name": area_name,
                                                                                                                                                                                        "select_category" : "",
                                                                                                                                                                                        "for_month": self.ui.DT_date_edit_2.date().month(),
                                                                                                                                                                                        "for_year": self.ui.DT_date_edit_2.date().year(),
                                                                                                                                                                                        "action_for": "DT"}))
        try:
            if hasattr(self, 'DT_dashboard_worker') and self.DT_dashboard_worker.isRunning():
                self.spinner.stop()
                return
            self.DT_dashboard_worker = WorkerThread(fetch_data)
            self.DT_dashboard_worker.finished.connect(lambda res: on_data_fetched_dashboard(res, area_name, target, year, view_by))
            self.DT_dashboard_worker.start()
        except Exception as e:
            self.spinner.stop()
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    # @QtCore.pyqtSlot()
    # def show_sorting_filter_Downtime(self):
    #     if self.ui.DT_show_sorting_btn.isChecked():
    #         self.ui.frame_106.setEnabled(True)
    #         self.DT_silde_bar_animation.setStartValue(0)
    #         self.DT_silde_bar_animation.setEndValue(535)
    #     else:
    #         self.ui.frame_106.setEnabled(False)
    #         self.DT_silde_bar_animation.setStartValue(535)
    #         self.DT_silde_bar_animation.setEndValue(0)
    #     self.DT_silde_bar_animation.start()

    # def Sparkline_chart(self, widget, data, color, title=""):
    #     old_layout = widget.layout()
    #     if old_layout is not None:
    #         while old_layout.count():
    #             item = old_layout.takeAt(0)
    #             if item.widget():
    #                 item.widget().deleteLater()
    #     else:
    #         new_layout = QtWidgets.QVBoxLayout()
    #         new_layout.setContentsMargins(0, 0, 0, 30)
    #         new_layout.setSpacing(0)
    #         widget.setLayout(new_layout)
    #     plot = pg.PlotWidget()
    #     plot.setAntialiasing(True)
    #     plot.setFixedSize(140, 100)
    #     plot.setBackground(None)
    #     plot.hideAxis('left')
    #     plot.hideAxis('bottom')
    #     plot.setTitle(
    #         f'<span style="color: grey; font-size: 8pt">{title}</span>')
    #     plot.setMouseEnabled(x=False, y=False)
    #     plot.setMenuEnabled(False)
    #     x = np.arange(len(data))
    #     y = np.array(data, dtype=float)

    #     if len(y) > 100:
    #         y_smooth = gaussian_filter1d(y, sigma=len(y) / 90)
    #         y_smooth = np.clip(y_smooth, 0, None)
    #     else:
    #         y_smooth = y

    #     curve = pg.PlotCurveItem(x, y_smooth, pen=pg.mkPen(color, width=1.5))
    #     fill = pg.FillBetweenItem(
    #         curve,
    #         pg.PlotCurveItem(x, np.zeros_like(y_smooth)),
    #         brush=pg.mkBrush(116, 185, 232, 80)
    #     )
    #     plot.addItem(curve)
    #     plot.addItem(fill)
    #     widget.setMaximumSize(135, 150)
    #     widget.layout().addWidget(plot)

    def DT_KPI_chart(self, widget,value, target_value, previous_value,html_doc = True, label="MTTR"):
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
        
        bullet_status_mttr = Bullet_Status_Bar(value=value, target_value=target_value , previous_value= previous_value, html_doc = html_doc, format_time=self.change_time_format, label=label)
        layout = widget.layout()
        layout.addWidget(bullet_status_mttr)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

    def DT_table_show(self, data, working_time, group_col="Line Name", target_df = None ,area_name=None, target=None, year=None, view_by="day"):
        if data is None or data.empty or working_time is None or working_time.empty:
            QtWidgets.QMessageBox.information(
                self, "No data", "No downtime records found for the selected area and date.")
            return
        try:
            if hasattr(self, 'DT_dashboard_summary_model'):
                self.DT_dashboard_summary_model.removeRows(0, self.DT_dashboard_summary_model.rowCount())
            gb = data.groupby(group_col)
            if group_col == "Machine Code":
                line_of_machine = gb["Line Name"].first().to_dict()
            if group_col != "Error Code":
                summary_df = pd.DataFrame({
                    group_col: gb["Total Loss Time"].sum().index,
                    "Total Downtime": gb["Total Loss Time"].sum().values,
                    "Failure Event": gb.size().values,
                    "MTTR": gb["Repair Time"].mean().values,
                    "MTBF": gb["Total Loss Time"].apply(
                        lambda x: (
                            float(working_time.loc[ (working_time["Line Name"] == line_of_machine[x.name] if group_col == "Machine Code" else working_time["Line Name"] == x.name), "Working Time"].sum(
                            ))
                         - x.sum()) / len(x)
                        if len(x) > 0
                        else float(working_time.loc[(working_time["Line Name"] == line_of_machine[x.name] if group_col == "Machine Code" else working_time["Line Name"] == x.name), "Working Time"].sum() -  x.sum(0))
                    ).values
                })
                print("summary_df", working_time.loc[(working_time["Line Name"] == "A13"), "Working Time"].sum())
                summary_df = summary_df.merge(
                    target_df[[group_col, "MTTR Target"]],
                    on=group_col,
                    how="left"
                )
                summary_df["MTTR Evaluate"] = (
                    summary_df["MTTR"] <= summary_df["MTTR Target"]
                ).astype(int)

                summary_df = summary_df.merge(
                    target_df[[group_col, "MTBF Target"]],
                    on=group_col,
                    how="left"
                )
                summary_df["MTBF Evaluate"] = (
                    summary_df["MTBF"] >= summary_df["MTBF Target"]
                ).astype(int)
                summary_df.loc[
                    summary_df["MTTR Target"].isna(),
                    "MTTR Evaluate"
                ] = 1

                summary_df.loc[
                    summary_df["MTBF Target"].isna(),
                    "MTBF Evaluate"
                ] = 1
                summary_df["MTTR"] = summary_df["MTTR"].apply(
                    lambda x: str(self.change_time_format(x, "m", output_unit=True)) if x > 0 else "N/A")
                summary_df["MTBF"] = summary_df["MTBF"].apply(
                    lambda x: str(self.change_time_format(x, "m", output_unit=True)) if x > 0 else "N/A")
                summary_df = summary_df.sort_values(
                    by=group_col, ascending=False).reset_index(drop=True)
                evaluate_df = summary_df[["MTTR Evaluate", "MTBF Evaluate"]] if group_col != "Error Code" else pd.DataFrame()
                summary_df = summary_df.drop(columns=evaluate_df.columns)
                summary_df = summary_df.drop(columns=["MTTR Target", "MTBF Target"])
                header = ["\n".join(f"{group_col}".split()), "Total\nDowntime",
                      "Down\nEvents", "MTTR", "MTBF"]
            else:
                summary_df = pd.DataFrame({
                    group_col: gb["Total Loss Time"].sum().index,
                    "Total Downtime": gb["Total Loss Time"].sum().values,
                    "Failure Event": gb.size().values})
                header = ["\n".join(f"{group_col}".split()), "Total\nDowntime", "Down\nEvents"]
            summary_df["Total Downtime"] = summary_df["Total Downtime"].apply(
                    lambda x: str(self.change_time_format(x, "m", output_unit=True)))
            self.ui.DT_table.setUpdatesEnabled(False)
            self.ui.DT_table.setSortingEnabled(False)
            if hasattr(self, 'DT_dashboard_summary_model'):
                self.DT_dashboard_summary_model.removeRows(
                    0, self.DT_dashboard_summary_model.rowCount())
            self.DT_dashboard_summary_model = QtGui.QStandardItemModel(
                len(summary_df), len(summary_df.columns))
            self.DT_dashboard_summary_model.setHorizontalHeaderLabels(header)
            for r in range(len(summary_df)):
                for c in range(len(summary_df.columns)):
                    if c == 0:
                        item = QtGui.QStandardItem(str(summary_df.iat[r, c]))
                        item.setTextAlignment(QtCore.Qt.AlignCenter)
                    # elif group_col != "Error Code" and c in [3, 4]:
                    elif group_col not in ["Error Code", "Machine Code"] and c in [3, 4]:
                        if evaluate_df.iat[r, c - 3] == 1:
                            item = QtGui.QStandardItem(str(summary_df.iat[r, c]))
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                        else:
                            item = QtGui.QStandardItem(str(summary_df.iat[r, c]))
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                            item.setBackground(QtGui.QColor(255, 0, 0, 127))
                    else:
                        value = round(summary_df.iat[r, c], 0) if isinstance(
                            summary_df.iat[r, c], float) and summary_df.iat[r, c] == 0 else summary_df.iat[r, c]
                        item = QtGui.QStandardItem(str(value))
                        item.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.DT_dashboard_summary_model.setItem(r, c, item)
            self.ui.DT_table.setModel(self.DT_dashboard_summary_model)
            self.ui.DT_table.horizontalScrollBar().setVisible(False)
            self.ui.DT_table.setEditTriggers(
                QtWidgets.QAbstractItemView.NoEditTriggers)
            self.ui.DT_table.setSelectionBehavior(
                QtWidgets.QAbstractItemView.SelectRows)
            self.ui.DT_table.setSelectionMode(
                QtWidgets.QAbstractItemView.ExtendedSelection)
            self.ui.DT_table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
            for c in range(1, len(summary_df.columns)):
                self.ui.DT_table.horizontalHeader().setSectionResizeMode(c, QtWidgets.QHeaderView.Stretch)
            self.ui.DT_table.setSortingEnabled(True)
            self.ui.DT_table.setUpdatesEnabled(True)
            @QtCore.pyqtSlot(QtCore.QModelIndex)
            def on_double_click(index):
                selected_category = self.DT_dashboard_summary_model.item(index.row(), 0).text()
                total = summary_df.loc[summary_df[group_col] == selected_category, "Total Downtime"].values[0]
                self.Detail_dashboard(selected_category=   selected_category, 
                                        group_col=group_col, area_name=area_name,
                                        target=target, year=year, view_by=view_by, 
                                        total = total,
                                        cum = total,
                                          data =  data[data[group_col] == selected_category],
                                          summary_df = summary_df,
                                          name = self.machine_name_dict.get(selected_category, '') if group_col == 'Machine Code' else self.error_code_dict.get(selected_category, '') if group_col == 'Error Code' else None
                                        )
                
            self.safe_connect(self.ui.DT_table.doubleClicked, lambda index: on_double_click(index))
            return summary_df
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load summary data: {e}")
            return None

    def DT_action_report_show(self, data, widget=None):
        widget.clear()
        if data is None or len(data) == 0:
            return
        line_action = data[data["Line Name"] != ""]
        machine_action = data[data["Machine Code"] != ""]
        error_action = data[data["Error Code"] != ""]
        html_contents = ""
        if not line_action.empty:
            html_contents += "<h3>Line Actions</h3>"
            for _, row in line_action.iterrows():
                line_name = row["Line Name"]
                comment = row["Action Content"]
                report_link = row["Action Report Link"]
                html_contents += f'''<p style="margin:0;">
                                        • Line {line_name}
                                    </p>'''
                if report_link:
                    report_link = json.loads(report_link) if report_link else []
                    for link in report_link:
                        html_contents += f"""
                                           <div style="margin-left:20px;">
                                            Comment: {comment}<br>
                                            Report:
                                            <a href="{link['file_path']}"
                                            style="color:#007acc; text-decoration:underline;">
                                                {link['file_name']}
                                            </a>
                                        </div>
                                        """
        if not machine_action.empty:
            html_contents += "<h3>Machine Actions</h3>"
            for _, row in machine_action.iterrows():
                machine_code = row["Machine Code"]
                comment = row["Action Content"]
                report_link = row["Action Report Link"]
                html_contents += f'''<p style="margin:0;">
                                        • Machine {machine_code}
                                    </p>'''
                if report_link:
                    report_link = json.loads(report_link) if report_link else []
                    for link in report_link:
                        html_contents += f"""
                                           <div style="margin-left:20px;">
                                            Comment: {comment}<br>
                                            Report:
                                            <a href="{link['file_path']}"
                                            style="color:#007acc; text-decoration:underline;">
                                                {link['file_name']}
                                            </a>
                                        </div>
                                        """
        if not error_action.empty:
            html_contents += "<h3>Error Code Actions</h3>"
            for _, row in error_action.iterrows():
                error_code = row["Error Code"]
                comment = row["Action Content"]
                report_link = row["Action Report Link"]
                html_contents += f'''<p style="margin:0;">
                                        • Error Code {error_code}
                                    </p>'''
                if report_link:
                    report_link = json.loads(report_link) if report_link else []
                    for link in report_link:
                        html_contents += f"""
                                             <div style="margin-left:20px;">
                                            Comment: {comment}<br>
                                            Report:
                                            <a href="{link['file_path']}"
                                            style="color:#007acc; text-decoration:underline;">
                                                {link['file_name']}
                                            </a>
                                        </div>
                                        """
        widget.setHtml(html_contents)

    @QtCore.pyqtSlot(QtWidgets.QListWidgetItem)
    def DT_problem_report_open(self, item):
        report = item.data(QtCore.Qt.UserRole)
        if report is None:
            return
        try:
            file_path = report[6]
            if file_path and os.path.exists(os.path.normpath(file_path)):
                os.startfile(os.path.normpath(file_path))
            else:
                QtWidgets.QMessageBox.warning(
                    self, "Warning", "File path not found or not available.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to open report: {e}")

    @QtCore.pyqtSlot()
    def DT_detail_chart_drawing(self, area_name, group_col, value_col, data, title="", target=None, year=None, view_by="day" , mttr_value=None, mtbf_value=None, mttr_target=None, mtbf_target=None, target_df=None):
        if self.DT_chart_current_group == group_col:
            return
        if group_col == "Line Name":
            self.DT_chart_current_group = "Line Name"
            self.style_button_with_shadow((self.ui.DT_detail_chart_line_btn, self.ui.DT_detail_chart_machine_btn,
                                          self.ui.DT_detail_chart_error_btn, self.ui.DT_detail_chart_time_btn))
        elif group_col == "Machine Code":
            self.DT_chart_current_group = "Machine Code"
            self.style_button_with_shadow((self.ui.DT_detail_chart_machine_btn, self.ui.DT_detail_chart_line_btn,
                                          self.ui.DT_detail_chart_error_btn, self.ui.DT_detail_chart_time_btn))
        elif group_col == "Error Code":
            self.DT_chart_current_group = "Error Code"
            self.style_button_with_shadow((self.ui.DT_detail_chart_error_btn, self.ui.DT_detail_chart_line_btn,
                                          self.ui.DT_detail_chart_machine_btn, self.ui.DT_detail_chart_time_btn))
        else:
            self.DT_chart_current_group = "Downtime Start Time"
            self.style_button_with_shadow((self.ui.DT_detail_chart_time_btn, self.ui.DT_detail_chart_line_btn,
                                          self.ui.DT_detail_chart_machine_btn, self.ui.DT_detail_chart_error_btn))
            self.ui.DT_chart_legend.hide()
            self.DT_time_density_chart(data, value_col)
            return
        try:
            summary_df = self.DT_table_show(data, working_time = self.working_time, group_col = group_col, target_df=target_df, target=target, year=year, view_by=view_by, area_name=area_name)
            self.ui.DT_chart_legend.show()
            old_layout = self.ui.DT_chart.layout()
            if old_layout is not None:
                while old_layout.count():
                    item = old_layout.takeAt(0)
                    if item.widget():
                        item.widget().deleteLater()
            else:
                new_layout = QtWidgets.QVBoxLayout()
                new_layout.setContentsMargins(0, 0, 0, 0)
                new_layout.setSpacing(0)
                self.ui.DT_chart.setLayout(new_layout)
            self.bar_item = None
            self.line_item = None
            pivot = (data.groupby([group_col, "Shift"])[value_col]
                     .sum()
                     .unstack(fill_value=0)
                     .reindex(columns=["Shift 1", "Shift 2", "Shift 3"], fill_value=0))
            pivot["total"] = pivot.sum(axis=1)
            pivot = pivot.sort_values(
                "total", ascending=False).drop(columns="total")
            if len(pivot) > 25:
                pivot = pivot.iloc[:25]
            categories = pivot.index.tolist()
            s1 = pivot["Shift 1"].values.astype(float)
            s2 = pivot["Shift 2"].values.astype(float)
            s3 = pivot["Shift 3"].values.astype(float)
            total = s1 + s2 + s3
            x = np.arange(len(categories))
            if group_col == "Machine Code":
                _angle = -45
                dx = -15
                dy = 5
            elif group_col == "Error Code" and len(categories) > 10:
                _angle = -40
                dx = 0
                dy = 0
            else:
                _angle = 0
                dx = 0
                dy = 0
            x_axis = RotatedAxisItem(
                angle=_angle, dx=dx, dy=dy, orientation='bottom')
            plot = pg.PlotWidget(axisItems={'bottom': x_axis})
            chart_font = QtGui.QFont("Comic Sans MS", 9)
            chart_font.setStyleStrategy(QtGui.QFont.PreferAntialias)
            chart_font.setHintingPreference(QtGui.QFont.PreferFullHinting)
            chart_font.setBold(True)
            plot.setBackground(None)
            plot.setTitle(
                f'<span style="color: grey; font-size: 10pt; font-weight: bold">{title}</span>')
            plot.showGrid(x=False, y=False, alpha=0.3)
            plot.hideButtons()
            y0_axis = plot.getAxis('left')
            y1_axis = plot.getAxis('right')
            x_axis = plot.getAxis('bottom')
            y0_axis.setLabel("Time (minutes)")
            y0_axis.setPen(None)
            y0_axis.setTextPen('gray')
            x_axis.setLabel(group_col)
            if _angle == -45:
                x_axis.setStyle(tickTextHeight=10)
                x_axis.setHeight(70)
            plot.showAxis('right')
            y1_axis.setLabel("Cumulative %")
            y1_axis.setPen(None)
            y1_axis.setTicks(
                [[(0, '0%'), (20, '20%'), (40, '40%'), (60, '60%'), (80, '80%'), (100, '100%')]])
            y1_axis.setTextPen('gray')
            x_axis.setTickFont(chart_font)
            y0_axis.setTickFont(chart_font)
            y1_axis.setTickFont(chart_font)
            colors = {
                "Shift 1": (187, 78, 139, 150),
                "Shift 2": (142, 71, 130, 150),
                "Shift 3": (60, 209, 197, 150),
            }
            bar1 = pg.BarGraphItem(x=x, y0=np.zeros(len(x)), y1=s1,
                                   width=0.6, brush=pg.mkBrush(*colors["Shift 1"]),
                                   pen=pg.mkPen((248, 12, 145, 255), width=0.5))
            bar2 = pg.BarGraphItem(x=x, y0=s1, y1=s1+s2,
                                   width=0.6, brush=pg.mkBrush(*colors["Shift 2"]),
                                   pen=pg.mkPen((198, 53, 170, 255), width=0.5))
            bar3 = pg.BarGraphItem(x=x, y0=s1+s2, y1=total,
                                   width=0.6, brush=pg.mkBrush(*colors["Shift 3"]),
                                   pen=pg.mkPen((0, 255, 235, 255), width=0.5))
            plot.addItem(bar1)
            plot.addItem(bar2)
            plot.addItem(bar3)
            cum = np.cumsum(total) / total.sum() * 100
            vb2 = pg.ViewBox()
            plot.scene().addItem(vb2)
            vb2.setZValue(plot.getViewBox().zValue() + 10)
            transparent = pg.mkPen((0, 0, 0, 0))
            x_axis.setPen(transparent)
            x_axis.setTickPen(transparent)
            x_axis.setTicks([list(zip(x, categories))])
            x_axis.setStyle(tickTextOffset=20,  tickLength=0)
            y1_axis.unlinkFromView = lambda: None 
            y1_axis.linkToView(vb2)
            vb2.setXLink(plot)

            def update_views():
                vb2.setGeometry(plot.getViewBox().sceneBoundingRect())
                vb2.linkedViewChanged(plot.getViewBox(), vb2.XAxis)

            region = pg.LinearRegionItem(
                movable=False,
                brush=(255,255,255,40),
                pen=pg.mkPen((146, 207, 171,150), width=0)
            )
            plot.addItem(region)
            region.hide()
            self._hover_index = -1

            @QtCore.pyqtSlot(object)
            def on_hover(pos):
                if not plot.sceneBoundingRect().contains(pos):
                    QtWidgets.QToolTip.hideText()
                    if self._hover_index != -1:
                        region.hide()
                        self._hover_index = -1
                    return
                
                mouse_point = plot.getViewBox().mapSceneToView(pos)
                index = round(mouse_point.x())
                if not (0 <= index < len(categories)):
                    QtWidgets.QToolTip.hideText()
                    if self._hover_index != -1:
                        region.hide()
                        self._hover_index = -1
                    return
                if index != self._hover_index:
                    self._hover_index = index
                    region.setRegion((index - 0.3, index + 0.3))
                    region.show()
                QtWidgets.QToolTip.showText(
                    QtGui.QCursor.pos(),
                    categories[index]
                    + (
                        f"\n{self.error_code_dict.get(categories[index], '')}"
                        if group_col == 'Error Code'
                        else ''
                    )
                    + (
                        f"\n{self.machine_name_dict.get(categories[index], '')}"
                        if group_col == 'Machine Code'
                        else ''
                    )
                    + f"\nShift 1: {s1[index]:.2f} min"
                    + f"\nShift 2: {s2[index]:.2f} min"
                    + f"\nShift 3: {s3[index]:.2f} min"
                    + f"\nTotal: {total[index]:.2f} min"
                    + f"\nCumulative: {cum[index] - cum[index-1] if index > 0 else cum[index]:.2f} %"
                )
            @QtCore.pyqtSlot()
            def on_dbclicked(evt, categories):
                if evt.button() == QtCore.Qt.LeftButton:
                    mouse_point = plot.getViewBox().mapSceneToView(evt.scenePos())
                    index = round(mouse_point.x())
                    if not 0 <= index < len(categories):
                        return
                    selected_category = categories[index]
                    self.Detail_dashboard(selected_category=selected_category, 
                                            group_col=group_col, area_name=area_name,
                                            target=target, year=year, view_by=view_by,
                                            total = data[data[group_col] == selected_category]["Total Loss Time"].sum(),
                                            cum = cum[index] - cum[index-1] if index > 0 else cum[index],
                                            data =  data[data[group_col] == selected_category],
                                            summary_df = summary_df,
                                            name = self.machine_name_dict.get(selected_category, '') if group_col == 'Machine Code' else self.error_code_dict.get(selected_category, '') if group_col == 'Error Code' else None
                                            )
            
            @QtCore.pyqtSlot()
            def leave_plot(event):
                QtWidgets.QToolTip.hideText()
                region.hide()
                self._hover_index = -1

            plot.leaveEvent = leave_plot
            
            plot.getViewBox().sigResized.connect(update_views)
            if len(x) >= 3:
                x_smooth = np.linspace(x.min(), x.max(), 300)
                pchip = PchipInterpolator(x, cum)
                cum_smooth = np.clip(pchip(x_smooth), 0, 100)
            else:
                x_smooth = x
                cum_smooth = cum

            self.line_item = pg.PlotCurveItem(
                x=x_smooth, y=cum_smooth,
                pen=pg.mkPen((231, 76, 60), width=3),
                antialias=True
            )
            dot_item = pg.ScatterPlotItem(
                x=x, y=cum,
                size=8,
                pen=pg.mkPen((231, 76, 60), width=1.5),
                brush=pg.mkBrush(240, 240, 240),
                symbol='o',
                antialias=True
            )
            vb2.addItem(self.line_item)
            vb2.addItem(dot_item)
            vb2.setYRange(0, 100)
            update_views()
            y_min = 0
            y_max = max(total) * 1.2 if len(total) > 0 else 1
            tick_step = max(5, int(y_max / 8 / 5) * 5)
            major_ticks = [(v, str(int(v)))
                           for v in np.arange(0, y_max + tick_step, tick_step)]
            y0_axis.setTicks([major_ticks])
            plot.setYRange(y_min, y_max, padding=0)
            self.proxy = pg.SignalProxy(
                plot.scene().sigMouseMoved,
                rateLimit=10,
                slot=lambda evt: on_hover(evt[0])
            )
            plot.scene().sigMouseClicked.connect(lambda evt: on_dbclicked(evt, categories) if evt.double() else None)
            
            self.ui.DT_chart.layout().addWidget(plot)
            main_vb = plot.getViewBox()
            main_vb.mouseDragEvent = lambda ev, axis=None: ev.ignore()
            main_vb.wheelEvent = lambda ev: ev.ignore()
            vb2.mouseDragEvent = lambda ev, axis=None: ev.ignore()
            vb2.wheelEvent = lambda ev: ev.ignore()
            if self.ui.DT_chart_legend.layout() is not None:
                return
            legend_layout = QtWidgets.QHBoxLayout()
            legend_layout.setContentsMargins(0, 0, 0, 0)
            legend_layout.setSpacing(5)
            self.ui.DT_chart_legend.setLayout(legend_layout)
            horizontal_spacer = QtWidgets.QSpacerItem(
                20, 10, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
            legend_layout.addItem(horizontal_spacer)
            for label, color in colors.items():
                legend_item = QtWidgets.QWidget()
                legend_item.setFixedSize(100, 30)
                legend_layout_item = QtWidgets.QHBoxLayout(legend_item)
                legend_layout_item.setContentsMargins(0, 0, 0, 0)
                legend_layout_item.setSpacing(5)
                color_box = QtWidgets.QLabel()
                color_box.setFixedSize(12, 12)
                color_box.setStyleSheet(
                    f"background-color: rgba({color[0]}, {color[1]}, {color[2]}, {color[3]}); border: 1px solid rgba({color[0]}, {color[1]}, {color[2]}, 255); border-radius: 0px;")
                legend_layout_item.addWidget(color_box)
                legend_label = QtWidgets.QLabel()
                legend_label.setText(label)
                legend_label.setStyleSheet("color: gray; font-size: 8pt;")
                legend_layout_item.addWidget(legend_label)
                legend_layout.addWidget(legend_item)
            legend_layout.addItem(horizontal_spacer)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to draw chart: {e}")
    
    def Detail_dashboard(self, selected_category,area_name, target, year, view_by, group_col, total, cum, data, summary_df, mttr_value=None, mtbf_value=None, name = ""):
        try:
            if type(cum) == str:
                h, m, s = map(int, cum.split(':'))
                cum = h * 60 + m + s / 60
                total_downtime = summary_df["Total Downtime"].apply(lambda x: int(x.split(':')[0]) * 60 + int(x.split(':')[1]) + int(x.split(':')[2]) / 60)
                cum = (cum/total_downtime.sum()) * 100
            if type(total) != str:
                total = self.change_time_format(time_value=total, input_unit="m", output_unit=True)
            data_dict = { 
                        "selected_category": selected_category,
                        "area_name": area_name,
                        "target": target,
                        "year": year,
                        "view_by": view_by,
                        "group_col": group_col,
                        "total_downtime": total,
                        "percentage": cum,
                        "data_of_category": data,
                        "average_percent": 100/len(summary_df),
                        "name": name
                    }
            if group_col in ("Machine Code","Line Name"):
                data_dict["mttr"] = str(summary_df.loc[summary_df[group_col] == selected_category, "MTTR"].values[0])
                data_dict["mtbf"] = str(summary_df.loc[summary_df[group_col] == selected_category, "MTBF"].values[0])
                if group_col == "Line Name":
                    data_dict["mttr_target"] = self.downtime_target.loc[self.downtime_target["Line Name"] == selected_category, "MTTR Target"].iloc[0]
                    data_dict["mtbf_target"] = self.downtime_target.loc[self.downtime_target["Line Name"] == selected_category, "MTBF Target"].iloc[0]
                else:
                    # data_dict["mttr_target"] = self.downtime_target.loc[self.downtime_target["Machine Code"] == selected_category, "MTTR Target"].iloc[0]
                    # data_dict["mtbf_target"] = self.downtime_target.loc[self.downtime_target["Machine Code"] == selected_category, "MTBF Target"].iloc[0]
                    data_dict["mttr_target"] = 0
                    data_dict["mtbf_target"] = 0
            else:
                data_dict["events_count"] = str(summary_df.loc[summary_df[group_col] == selected_category, "Failure Event"].values[0])
            dialog = Downtime_Detail_dashboard(
                database=self.database_process,
                catagory=selected_category,
                data_dict = data_dict,
                button_style_call_back = self.style_button_with_shadow,
                change_time_format_call_back = self.change_time_format,
                KPI_chart_call_back = self.DT_KPI_chart,
                icon_from_path_call_back = self.icon_from_path,
                extract_content_call_back = self.extract_content
            )
            dialog.closed.connect(
                lambda dlg: self.DT_detail_report_list.remove(dlg) if dlg in self.DT_detail_report_list else None
            )
            self.DT_detail_report_list.append(dialog)
            dialog.show()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to open detail: {e}")

    @QtCore.pyqtSlot()
    def DT_Viewby(self, changed_object):

        if changed_object == self.DT_filtered_dict.get(changed_object):
            return
        try:
            area_name = self.ui.DT_area_cbb.currentText()
            year = self.ui.DT_date_edit_2.date().year()
            if self.ui.DT_day_radiobtn.isChecked():
                self.ui.frame_91.setEnabled(False)
                self.ui.frame_91.setMaximumWidth(0)
                self.ui.frame_87.setEnabled(True)
                self.ui.frame_87.setMaximumWidth(200)
                if self.table_view_of_DTcalendar:
                    self.table_view_of_DTcalendar.show()
                self.DT_calendar_widget.setFixedSize(270, 183)
                self.ui.DT_date_edit_2.setDisplayFormat("dd-MMM-yyyy")
                self.safe_connect(self.ui.DT_date_edit_2.dateChanged, lambda: self.DT_filtering(
                    changed_object="date_range"))
                try:
                    self.ui.DT_year_cbb.currentTextChanged.disconnect()
                except TypeError:
                    pass
                target = self.ui.DT_date_edit_2.date().toString("yyyy-MM-dd")
            elif self.ui.DT_year_radiobtn.isChecked():
                self.ui.frame_87.setEnabled(False)
                self.ui.frame_87.setMaximumWidth(0)
                self.ui.frame_91.setEnabled(True)
                self.ui.frame_91.setMaximumWidth(200)
                # if self.ui.DT_year_cbb.count() == 0:
                #     self.ui.DT_year_cbb.addItems(
                #         [str(year) for year in range(2026, 2026+1)])
                self.ui.DT_year_cbb.setCurrentText(
                    str(self.ui.today.year))
                self.safe_connect(self.ui.DT_year_cbb.currentTextChanged,
                                lambda: self.DT_filtering(changed_object="year_range"))
                try:
                    self.ui.DT_date_edit_2.dateChanged.disconnect()
                except TypeError:
                    pass
                target = int(self.ui.DT_year_cbb.currentText())
            else:
                if not self.ui.frame_87.isEnabled():
                    self.ui.frame_91.setEnabled(False)
                    self.ui.frame_91.setMaximumWidth(0)
                    self.ui.frame_87.setEnabled(True)
                    self.ui.frame_87.setMaximumWidth(200)
                self.ui.DT_date_edit_2.setDisplayFormat("MMM-yyyy")
                if self.table_view_of_DTcalendar:
                    self.table_view_of_DTcalendar.hide()
                self.DT_calendar_widget.setFixedSize(250, 30)
                self.safe_connect(self.DT_calendar_widget.currentPageChanged, lambda year,
                                month: self.update_date_from_calendar(year, month, self.ui.DT_date_edit_2))
                self.safe_connect(self.ui.DT_date_edit_2.dateChanged, lambda: self.DT_filtering(
                    changed_object="month_range"))
                target = self.ui.DT_date_edit_2.date().month()

            self.Dashboard_Downtime_page_refresh(
                area_name=area_name, target=target, year=year, view_by=changed_object)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to filter data 1: {e}")

    @QtCore.pyqtSlot(int, int, object)
    def update_date_from_calendar(self, year, month, object=None):
        new_date = QtCore.QDate(year, month, 1)
        if object:
            object.setDate(new_date)

    @QtCore.pyqtSlot()
    def DT_filtering(self, changed_object=None):
        if changed_object == None:
            return
        try:
            year = self.ui.DT_date_edit_2.date().year()
            if changed_object == "date_range":
                area_name = self.ui.DT_area_cbb.currentText()
                date = self.ui.DT_date_edit_2.date().toString("yyyy-MM-dd")
                self.Dashboard_Downtime_page_refresh(
                    area_name, date, year, view_by="day")
            elif changed_object == "year_range":
                area_name = self.ui.DT_area_cbb.currentText()
                year = int(self.ui.DT_year_cbb.currentText())
                self.Dashboard_Downtime_page_refresh(
                    area_name, target=year, year=year, view_by="year")
            elif changed_object == "month_range":
                area_name = self.ui.DT_area_cbb.currentText()
                month_num = self.ui.DT_date_edit_2.date().month()
                self.Dashboard_Downtime_page_refresh(
                    area_name, month_num, year, view_by="month")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to filter data 2: {e}")

    def DT_time_density_chart(self, data, value_col):
        def process_small_data_to_smooth_data(raw_points):
            x_old = np.linspace(0, 1, len(raw_points))
            x_new = np.linspace(0, 1, len(raw_points) * 10)
            f = interp1d(x_old, raw_points, kind='cubic')
            data_resampled = f(x_new)
            data_smooth = gaussian_filter1d(data_resampled, sigma=10)

            range_val = data_smooth.max() - data_smooth.min()
            if range_val == 0:
                # Trả về toàn 0 nếu không có data
                return np.zeros((len(x_new), 1))

            final_data = (data_smooth - data_smooth.min()) / range_val
            return final_data.reshape(-1, 1)
        try:
            old_layout = self.ui.DT_chart.layout()
            if old_layout is not None:
                while old_layout.count():
                    item = old_layout.takeAt(0)
                    if item.widget():
                        item.widget().deleteLater()
            else:
                new_layout = QtWidgets.QVBoxLayout()
                new_layout.setContentsMargins(0, 0, 0, 0)
                new_layout.setSpacing(0)
                self.ui.DT_chart.setLayout(new_layout)
            plot = pg.PlotWidget()
            plot.setBackground(None)
            plot.setTitle(
                f'<span style="color: grey; font-size: 10pt; font-weight: bold">Downtime Density Over Time</span>')
            plot.showGrid(x=True, y=True, alpha=0.3)
            x_axis = plot.getAxis('bottom')
            y_axis = plot.getAxis('left')
            x_axis.setLabel("Time")
            y_axis.setLabel("")
            plot.setMouseEnabled(False, False)
            full_10min = pd.timedelta_range(
                start="0 days", periods=144, freq="10min")
            pivot = (
                data.assign(
                    Time_10min=data["Downtime Start Time"].dt.floor(
                        "10min")  # timedelta trực tiếp
                )
                .groupby(["Time_10min", "Shift"])
                .size()
                .unstack(fill_value=0)
                .reindex(index=full_10min, columns=["Shift 1", "Shift 2", "Shift 3"], fill_value=0)
            )
            Shift_1 = pivot["Shift 1"].loc["0 days 06:00:00":"0 days 14:00:00"]
            Shift_2 = pivot["Shift 2"].loc["0 days 14:00:00":"0 days 22:00:00"]
            Shift_3 = pd.concat([pivot["Shift 3"].loc["0 days 22:00:00":"0 days 23:59:59"],
                                pivot["Shift 3"].loc["0 days 00:00:00":"0 days 06:00:00"]])
            Shift_1_density = process_small_data_to_smooth_data(Shift_1.values)
            Shift_2_density = process_small_data_to_smooth_data(Shift_2.values)
            Shift_3_density = process_small_data_to_smooth_data(Shift_3.values)
            s1 = Shift_1_density.flatten()
            s2 = Shift_2_density.flatten()
            s3 = Shift_3_density.flatten()
            min_len = min(len(s1), len(s2), len(s3))
            gap_col = np.full(min_len, -1.0)
            image_data = np.column_stack([
                s3[:min_len],
                gap_col,
                s2[:min_len],
                gap_col,
                s1[:min_len]
            ])
            lut = np.zeros((256, 3), dtype=np.uint8)
            lut[0] = [255, 255, 255]
            lut[127:, 0] = np.linspace(252, 203,   129).astype(np.uint8)
            lut[127:, 1] = np.linspace(220, 23,    129).astype(np.uint8)
            lut[127:, 2] = np.linspace(221, 17,    129).astype(np.uint8)
            img = pg.ImageItem()
            img.setImage(image_data)
            img.setLookupTable(lut)
            img.setLevels([-1, 1])
            font = QtGui.QFont()
            font.setFamily("Comic Sans MS")
            font.setPointSize(10)
            font.setBold(True)
            plot.showGrid(x=False, y=False)
            x_axis.setTicks([[(i * min_len / 8, f'{i}h') for i in range(9)]])
            x_axis.setTextPen(pg.mkPen(color=(100, 100, 100)))
            x_axis.setPen(None)
            x_axis.setTickFont(font)
            y_axis.setTicks(
                [[(0.5, 'Shift 3'), (2.5, 'Shift 2'), (4.5, 'Shift 1')]])
            y_axis.setTextPen(pg.mkPen(color=(100, 100, 100)))
            y_axis.setPen(None)
            y_axis.setTickFont(font)
            plot.addItem(img)
            plot.setXRange(0, min_len, padding=0)
            plot.setYRange(0, 3, padding=0)
            plot.setFixedSize(800, 300)
            plot.enableAutoRange(axis=pg.ViewBox.XYAxes, enable=True)
            shift_labels = [("22h", "6h"), ("14h", "22h"), ("6h", "14h")]
            for i, (t_start, t_end) in enumerate(shift_labels):
                y_pos = i*2 + 0.88
                lbl_s = pg.TextItem(t_start, color=(
                    100, 100, 100), anchor=(0, 1))
                lbl_e = pg.TextItem(t_end,   color=(
                    100, 100, 100), anchor=(1, 1))
                lbl_s.setFont(font)
                lbl_e.setFont(font)
                lbl_s.setPos(0,        y_pos)
                lbl_e.setPos(min_len,  y_pos)
                plot.addItem(lbl_s)
                plot.addItem(lbl_e)
            self.ui.DT_chart.layout().addWidget(plot)

        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to generate density chart: {e}")

    @QtCore.pyqtSlot()
    def Data_Downtime_page(self):
        self.style_button_with_shadow(
            (self.ui.DT_data_btn, self.ui.DT_import_data_btn, self.ui.DT_problem_report_btn, self.ui.DT_dashboard_btn))
        self.ui.DT_stacked_widget.setCurrentWidget(self.ui.DT_detail_data_page)
        self.style_button_with_shadow(
            (self.ui.DT_DD_error_chart_btn, self.ui.DT_DD_line_chart_btn, self.ui.DT_DD_machine_chart_btn))
        if self.ui.DT_DD_area_cbb.count() > 0:
            return
        try:
            headers = ["ID", "Date", "Line", "Start\nTime", "Technical\nStart", "Finish\nTime", "Total Loss\nTime", "Repair\nTime",
                       "Technical\nName", "Failure\nCode", "Failure\nDescription", "Reason", "Recommendation", "Machine Code"]
            self.DT_detail_model = QtGui.QStandardItemModel(0, len(headers))
            self.DT_detail_model.setHorizontalHeaderLabels(headers)
            self.ui.DT_DD_record_table.setUpdatesEnabled(False)
            self.ui.DT_DD_record_table.setSortingEnabled(False)
            self.ui.DT_DD_record_table.resizeColumnsToContents()
            self.ui.DT_summary_table.setUpdatesEnabled(False)
            self.ui.DT_summary_table.setSortingEnabled(False)
            self.ui.DT_DD_record_table.setModel(self.DT_detail_model)
            vetical_header = ["Area", "Total Failure", "Total Loss", "MTTR",
                              "MTBF", "Machine with Most Failure", "Failure Code Most Frequent"]
            self.DT_DD_summary_model = QtGui.QStandardItemModel(
                len(vetical_header), 2)
            for i in range(len(vetical_header)):
                item = QtGui.QStandardItem(vetical_header[i])
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.DT_DD_summary_model.setItem(i, 0, item)
            self.ui.DT_DD_summary_table.setModel(self.DT_DD_summary_model)
            self.ui.DT_DD_summary_table.horizontalScrollBar().setVisible(False)
            self.ui.DT_DD_summary_table.setColumnWidth(0, 180)
            self.ui.DT_DD_summary_table.setColumnWidth(
                1, self.ui.DT_DD_summary_table.width()-179)
            self.ui.DT_DD_summary_table.setUpdatesEnabled(True)
            self.ui.DT_DD_summary_table.setSortingEnabled(True)
            self.ui.DT_DD_record_table.setColumnHidden(0, True)
            self.ui.DT_DD_record_table.setUpdatesEnabled(True)
            self.ui.DT_DD_record_table.setSortingEnabled(True)
            self.ui.DT_DD_date_edit.setDisplayFormat("MMM-yyyy")
            self.DT_DD_calendar_widget = self.ui.DT_DD_date_edit.calendarWidget()
            self.table_view_of_DTcalendar = self.DT_DD_calendar_widget.findChild(
                QtWidgets.QTableView)
            if self.table_view_of_DTcalendar:
                self.table_view_of_DTcalendar.hide()
            self.DT_DD_calendar_widget.setFixedSize(250, 30)
            self.DT_DD_calendar_widget.currentPageChanged.connect(
                lambda year, month: self.update_date_from_calendar(year, month, self.ui.DT_DD_date_edit))
            self.ui.DT_DD_area_cbb.addItems(self.areas)
            self.Load_production_line_base_area()
            self.ui.DT_DD_date_edit.setDate(self.ui.today)
            self.safe_connect(
                self.ui.DT_DD_area_cbb.currentTextChanged, self.Load_production_line_base_area)
            self.safe_connect(self.ui.DT_DD_machine_code_lnedit.textChanged, lambda text: self.filter_suggestion(target=self.ui.DT_DD_machine_code_lnedit,
                                                                                                                 text="DISTINCT ( m.machine_code )", table="`downtime_records` as dr",
                                                                                                                 where=f""" JOIN `machines` as m
                                                                                                        ON dr.machine_id = m.machine_id
                                                                                                        JOIN `production_lines` as pl
                                                                                                        ON dr.line_id = pl.line_id
                                                                                                        WHERE {f'pl.line_name = "{self.ui.DT_DD_line_cbb.currentText()}" AND' if self.ui.DT_DD_line_cbb.currentText() != "All Lines" else ""} 
                                                                                                        m.machine_code LIKE '%{text}%' AND MONTH(dr.downtime_date) = {self.ui.DT_DD_date_edit.date().month()} 
                                                                                                        AND YEAR(dr.downtime_date) = {self.ui.DT_DD_date_edit.date().year()} """))
            self.safe_connect(self.ui.DT_DD_load_btn.clicked, lambda: self.Load_Downtime_data(area_name=self.ui.DT_DD_area_cbb.currentText(),
                                                                                              line_name=self.ui.DT_DD_line_cbb.currentText(),
                                                                                              machine_code=self.ui.DT_DD_machine_code_lnedit.text(),
                                                                                              date=self.ui.DT_DD_date_edit.date().toString("yyyy-MM-dd")))
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to initialize data table: {e}")
            return

    @QtCore.pyqtSlot()
    def Load_production_line_base_area(self):
        area_name = self.ui.DT_DD_area_cbb.currentText()
        try:
            result = self.database_process.query(sql='''SELECT DISTINCT pl.line_name
                                                            FROM `production_lines` pl
                                                            JOIN `downtime_areas_production_lines` dapl ON pl.line_id = dapl.line_id
                                                            JOIN `downtime_areas` da ON dapl.downtime_area_id = da.downtime_area_id
                                                            WHERE da.downtime_area_name = :area_name;''', params={'area_name': area_name})

            self.ui.DT_DD_line_cbb.clear()
            self.ui.DT_DD_line_cbb.addItem("All Lines")
            line_names = [row[0] for row in result]
            self.ui.DT_DD_line_cbb.addItems(line_names)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load production lines: {e}")

    @QtCore.pyqtSlot()
    def Load_Downtime_data(self, area_name, line_name, machine_code, date):
        try:
            query = f'''SELECT dr.Downtime_ID,dr.Date, dr.Line_Name, dr.Start_Time, dr.Start_Repair_Time, dr.End_Time, dr.Total_Loss, dr.Repair_Time, dr.Staff_Name, dr.Error_Code, er.error_description, er.reason, er.recommended_action, dr.Machine_Code
                        FROM `downtime_report` as dr
                        LEFT JOIN `error_codes_list` as er ON dr.Error_Code = er.error_code
                        WHERE dr.Downtime_Area = :area_name AND dr.Working_Month = MONTH(:date) AND YEAR(dr.Date) = YEAR(:date)'''
            query_wt = f'''SELECT lot.line_id, lot.operation_date, lot.operation_hours
                            FROM `line_operation_times` as lot
                            JOIN `downtime_areas_production_lines` as dapl ON lot.line_id = dapl.line_id
                            JOIN `downtime_areas` as da ON dapl.downtime_area_id = da.downtime_area_id
                            JOIN `production_lines` as pl ON lot.line_id = pl.line_id
                            WHERE da.downtime_area_name = :area_name AND MONTH(lot.operation_date) = MONTH(:date) AND YEAR(lot.operation_date) = YEAR(:date)'''
            params = {'area_name': area_name, 'date': date}
            if line_name != "All Lines":
                query += " AND dr.Line_Name = :line_name"
                query_wt += " AND pl.Line_Name = :line_name"
                params['line_name'] = line_name
            if machine_code:
                query += " AND dr.Machine_Code LIKE :machine_code"
                params['machine_code'] = f"%{machine_code}%"
            query += " ORDER BY dr.Date DESC, dr.Start_Time DESC"
            working_time = self.database_process.query(
                sql=query_wt, params=params)
            working_time = pd.DataFrame(working_time, columns=[
                                        "line_id", "Date", "operation_hours"])
            result = self.database_process.query(sql=query, params=params)
            self.ui.DT_DD_record_table.setUpdatesEnabled(False)
            self.ui.DT_DD_record_table.setSortingEnabled(False)
            self.DT_detail_model.removeRows(0, self.DT_detail_model.rowCount())
            self.DT_detail_model.setRowCount(len(result))
            column_names = [
                "id", "date", "line", "start_time", "technical_start_time",
                "finish_time", "total_loss_time", "repair_time",
                "technical_name", "error_code", "error_description",
                "reason", "recommended_action", "machine_code"
            ]
            dataframe_for_sumary_table = pd.DataFrame(result)
            dataframe_for_sumary_table.columns = column_names
            for r in range(len(result)):
                for c in range(len(result[0])):
                    item = QtGui.QStandardItem(str(result[r][c]))
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    item.setFlags(QtCore.Qt.ItemIsEnabled |
                                  QtCore.Qt.ItemIsSelectable)
                    self.DT_detail_model.setItem(r, c, item)
            self.ui.DT_DD_record_table.setEditTriggers(
                QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.EditKeyPressed)
            self.ui.DT_DD_record_table.setUpdatesEnabled(True)
            self.ui.DT_DD_record_table.setSortingEnabled(True)
            self.ui.DT_DD_record_table.setContextMenuPolicy(
                QtCore.Qt.CustomContextMenu)
            self.DT_summary_table_show(table=self.ui.DT_DD_summary_table, model=self.DT_DD_summary_model, area=area_name,
                                       data_frame=dataframe_for_sumary_table, working_time=working_time[["Date", "operation_hours"]])
            self.DT_summary_chart_show(self.ui.DT_DD_summary_chart_widget, "error", dataframe_for_sumary_table, (
                self.ui.DT_DD_error_chart_btn, self.ui.DT_DD_line_chart_btn, self.ui.DT_DD_machine_chart_btn))
            self.safe_connect(self.ui.DT_DD_insert_btn.clicked, lambda: self.insert_row_QTableview(
                self.DT_detail_model, self.ui.DT_DD_record_table))
            self.safe_connect(self.ui.DT_DD_error_chart_btn.clicked, lambda: self.DT_summary_chart_show(self.ui.DT_DD_summary_chart_widget, "error",
                              dataframe_for_sumary_table, (self.ui.DT_DD_error_chart_btn, self.ui.DT_DD_line_chart_btn, self.ui.DT_DD_machine_chart_btn)))
            self.safe_connect(self.ui.DT_DD_line_chart_btn.clicked, lambda: self.DT_summary_chart_show(self.ui.DT_DD_summary_chart_widget, "line",
                              dataframe_for_sumary_table, (self.ui.DT_DD_line_chart_btn, self.ui.DT_DD_error_chart_btn, self.ui.DT_DD_machine_chart_btn)))
            self.safe_connect(self.ui.DT_DD_machine_chart_btn.clicked, lambda: self.DT_summary_chart_show(self.ui.DT_DD_summary_chart_widget, "machine",
                              dataframe_for_sumary_table, (self.ui.DT_DD_machine_chart_btn, self.ui.DT_DD_error_chart_btn, self.ui.DT_DD_line_chart_btn)))
            self.safe_connect(self.ui.DT_DD_record_table.customContextMenuRequested, lambda pos: self.table_context_menu(pos, self.ui.DT_DD_record_table, actions=["edit", "separator", "delete"],
                                                                                                                         functions_dict={   "edit": lambda idx: print(f"Edit row {idx.row()} column {idx.column()}"),
                                                                                                                                            "delete": lambda idx: self.DT_detail_model.removeRow(idx.row()),
                                                                                                                                                        }))
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load downtime data: {e}")

    @QtCore.pyqtSlot()
    def insert_row_QTableview(self, model, table):
        try:
            new_row = []
            for _ in range(model.columnCount()):
                item = QtGui.QStandardItem("")
                item.setFlags(QtCore.Qt.ItemIsEnabled |
                              QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEditable)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                new_row.append(item)
            model.appendRow(new_row)
            source_index = model.index(model.rowCount() - 1, 0)
            proxy = table.model()
            view_index = proxy.mapFromSource(
                source_index) if proxy is not model else source_index
            table.scrollTo(view_index)
            table.setCurrentIndex(view_index)
            table.edit(view_index)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to insert new row: {e}")

    @QtCore.pyqtSlot()
    def Import_data_Downtime_page(self):
        self.style_button_with_shadow(
            (self.ui.DT_import_data_btn, self.ui.DT_problem_report_btn, self.ui.DT_dashboard_btn, self.ui.DT_data_btn))
        self.ui.DT_stacked_widget.setCurrentWidget(self.ui.DT_import_data_page)
        self.style_button_with_shadow(
            (self.ui.DT_error_chart_btn, self.ui.DT_line_chart_btn, self.ui.DT_machine_chart_btn))
        headers = ["Date", "Line", "Start\nTime", "Technical\nStart", "Finish\nTime",
                   "Total Loss\nTime", "Repair\nTime", "Technical\nName", "Failure\nCode", "Machine Code"]
        self.DT_model = QtGui.QStandardItemModel(0, len(headers))
        self.DT_model.setHorizontalHeaderLabels(headers)
        self.ui.DT_data_table.setModel(self.DT_model)
        self.ui.DT_data_table.setColumnWidth(0, 100)
        vetical_header = ["Area", "Total Failure", "Total Loss", "MTTR",
                          "MTBF", "Machine with Most Failure", "Failure Code Most Frequent"]
        self.DT_summary_model = QtGui.QStandardItemModel(
            len(vetical_header), 2)
        self.ui.DT_summary_table.setUpdatesEnabled(False)
        self.ui.DT_summary_table.setSortingEnabled(False)
        for i in range(len(vetical_header)):
            item = QtGui.QStandardItem(vetical_header[i])
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.DT_summary_model.setItem(i, 0, item)
        self.ui.DT_summary_table.setModel(self.DT_summary_model)
        self.ui.DT_summary_table.horizontalScrollBar().setVisible(False)
        self.ui.DT_summary_table.setColumnWidth(0, 180)
        self.ui.DT_summary_table.setColumnWidth(
            1, self.ui.DT_summary_table.width()-179)
        self.ui.DT_summary_table.setUpdatesEnabled(True)
        self.ui.DT_summary_table.setSortingEnabled(True)
        self.safe_connect(self.ui.DT_upload_data_btn.clicked,
                          self.DT_excel_upload)
        self.safe_connect(self.ui.DT_error_code_btn.clicked,
                          self.DT_error_code_show)

    @QtCore.pyqtSlot()
    def DT_excel_upload(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")

        def get_area_name():
            try:
                group_choose = Group_Area_Choose(
                    parent=self, database=self.database_process, file_path=file_path)
                if group_choose.exec() == QtWidgets.QDialog.Accepted:
                    return group_choose.selected_area, group_choose.excel_sheet_name
                return None, None
            except Exception as e:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Failed to choose area: {e}")
                return None , None
        if file_path:
            try:
                area_name, sheet_name = get_area_name()
                if area_name is None or sheet_name is None:
                    return
                excel_data_process = Downtime_Excel_Processor(
                    file_path=file_path, sheet_name=sheet_name, area_name=area_name, database=self.database_process)
                self.data, self.error_frame, self.working_time = excel_data_process.read_filter_excel()
                if self.data is not None:
                    downtime_input_dialog = Downtime_Input(parent=self, database=self.database_process, data_frame=self.data,
                                                           error_frame=self.error_frame, area_name=area_name, month_year=excel_data_process.month_year)
                    downtime_input_dialog.exec()
                if downtime_input_dialog.result() == QtWidgets.QDialog.Accepted:
                    self.ui.DT_data_table.setUpdatesEnabled(False)
                    self.ui.DT_data_table.setSortingEnabled(False)
                    self.DT_model.removeRows(0, self.DT_model.rowCount())
                    self.DT_model.setRowCount(
                        len(downtime_input_dialog.data_frame))
                    self.DT_data = downtime_input_dialog.data_frame
                    self.DT_month_year = downtime_input_dialog.month_year
                    for r in range(len(self.DT_data)):
                        for c in range(len(self.DT_data.columns)):
                            item = QtGui.QStandardItem(
                                str(self.DT_data.iat[r, c]))
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                            self.DT_model.setItem(r, c, item)
                    self.ui.DT_data_table.setUpdatesEnabled(True)
                    self.ui.DT_data_table.setSortingEnabled(True)
                    self.DT_summary_table_show(
                        self.ui.DT_summary_table, self.DT_summary_model, area_name, self.DT_data, self.working_time)
                    self.DT_summary_chart_show(self.ui.DT_summary_chart_widget, "error", self.DT_data, (
                        self.ui.DT_error_chart_btn, self.ui.DT_line_chart_btn, self.ui.DT_machine_chart_btn))
                    self.safe_connect(self.ui.DT_error_chart_btn.clicked, lambda: self.DT_summary_chart_show(
                        self.ui.DT_summary_chart_widget, "error", self.DT_data, (self.ui.DT_error_chart_btn, self.ui.DT_line_chart_btn, self.ui.DT_machine_chart_btn)))
                    self.safe_connect(self.ui.DT_line_chart_btn.clicked, lambda: self.DT_summary_chart_show(
                        self.ui.DT_summary_chart_widget, "line", self.DT_data, (self.ui.DT_line_chart_btn, self.ui.DT_machine_chart_btn, self.ui.DT_error_chart_btn)))
                    self.safe_connect(self.ui.DT_machine_chart_btn.clicked, lambda: self.DT_summary_chart_show(
                        self.ui.DT_summary_chart_widget, "machine", self.DT_data, (self.ui.DT_error_chart_btn, self.ui.DT_line_chart_btn, self.ui.DT_machine_chart_btn)))
                    self.safe_connect(self.ui.DT_import_database_btn.clicked,
                                      lambda: self.DT_import_database(self.working_time))
            except Exception as e:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Failed to upload data: {e}")

    def DT_summary_table_show(self, table, model, area, data_frame, working_time):
        try:
            area_lbl = area
            total_failure = data_frame.shape[0]
            total_loss = data_frame["total_loss_time"].sum()
            total_working_time = working_time.drop(columns=["Date"]).apply(
                pd.to_numeric, errors="coerce").fillna(0).to_numpy().sum()
            mttr = round(total_loss/total_failure,
                         2) if total_failure > 0 else 0
            mtbf = round((total_working_time*60 - total_loss) /
                         total_failure, 2) if total_failure > 0 else 0
            machine_most_failure = data_frame["machine_code"].mode(
            )[0] if not data_frame["machine_code"].mode().empty else "N/A"
            failure_code_most_frequent = data_frame["error_code"].mode(
            )[0] if not data_frame["error_code"].mode().empty else "N/A"
            summary_data = [area_lbl, f"{total_failure} times", f"{total_loss} mins",
                            f"{mttr} mins", f"{mtbf} mins", machine_most_failure, failure_code_most_frequent]
            table.setUpdatesEnabled(False)
            table.setSortingEnabled(False)
            for i in range(len(summary_data)):
                item = QtGui.QStandardItem(str(summary_data[i]))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                model.setItem(i, 1, item)
            table.setUpdatesEnabled(True)
            table.setSortingEnabled(True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to calculate summary: {e}")

    @QtCore.pyqtSlot()
    def DT_summary_chart_show(self, chart_widget, chart_type, data_frame, button_set):
        try:
            if chart_type == "error":
                self.style_button_with_shadow(button_set)
                self.DT_chart_drawing(chart_widget, data_frame, "error_code", "total_loss_time",
                                      "Top 10 Error Codes by Loss Time", "Error Code", "Total Loss Time (mins)")
            elif chart_type == "line":
                self.style_button_with_shadow(button_set)
                self.DT_chart_drawing(chart_widget, data_frame, "line", "total_loss_time",
                                      "Top 10 Lines by Loss Time", "Line", "Total Loss Time (mins)")
            elif chart_type == "machine":
                self.style_button_with_shadow(button_set)
                self.DT_chart_drawing(chart_widget, data_frame, "machine_code", "total_loss_time",
                                      "Top 10 Machines by Loss Time", "Machine", "Total Loss Time (mins)")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to show chart: {e}")

    def DT_chart_drawing(self, chart_widget, data_frame, group_by_col, value_col, title, x_label, y_label):
        layout = chart_widget.layout()
        if layout is not None:
            while layout.count():
                child = layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()
            layout.setContentsMargins(0, 0, 0, 0)
        else:
            layout = QtWidgets.QVBoxLayout(chart_widget)
            layout.setContentsMargins(0, 0, 0, 0)
            chart_widget.setLayout(layout)
        try:
            error_loss_time = data_frame.groupby(
                group_by_col)[value_col].sum().sort_values(ascending=False).head(10)
            widget_w = chart_widget.width() - 10
            widget_h = chart_widget.height() - 10
            dpi = 100
            fig_w = widget_w / dpi
            fig_h = widget_h / dpi

            fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=dpi)
            fig.patch.set_alpha(0.0)
            ax.set_facecolor("none")

            bars = ax.bar(x=range(len(error_loss_time)),
                          height=error_loss_time.values, color='#3FDAA7')

            ax.set_xticks(range(len(error_loss_time)))
            ax.set_xticklabels(error_loss_time.index,
                               rotation=90, ha='center', va='top', fontsize=7)
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
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to show chart: {e}")

    @QtCore.pyqtSlot()
    def DT_error_code_show(self):
        try:
            error_code_dialog = Error_code_management(
                parent=self, database=self.database_process)
            error_code_dialog.exec_()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to show error codes: {e}")

    @QtCore.pyqtSlot()
    def DT_import_database(self, working_time):
        if self.DT_data is None:
            QtWidgets.QMessageBox.warning(
                self, "No Data", "Please upload and review the data before importing to database.")
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
                    import_result = session.execute(
                        text(sql), import_data_list, execution_options={"executemany": True})
                    # import_working_time_result = session.execute(text(
                    #     sql_working_time), working_time_import_list, execution_options={"executemany": True})
                    session.commit()
                    QtWidgets.QMessageBox.information(
                        self, "Success", "Data has been successfully imported to the database.")
                    self.ui.DT_data_table.setUpdatesEnabled(False)
                    self.ui.DT_data_table.setSortingEnabled(False)
                    self.ui.DT_summary_table.setUpdatesEnabled(False)
                    self.ui.DT_summary_table.setSortingEnabled(False)
                    self.DT_model.removeRows(0, self.DT_model.rowCount())
                    self.DT_summary_model.removeRows(
                        0, self.DT_summary_model.rowCount())
                    self.ui.DT_summary_chart_widget.layout().deleteLater()
                    self.ui.DT_data_table.setUpdatesEnabled(True)
                    self.ui.DT_data_table.setSortingEnabled(True)
                    self.ui.DT_summary_table.setUpdatesEnabled(True)
                    self.ui.DT_summary_table.setSortingEnabled(True)
                except Exception:
                    session.rollback()
                    raise
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to import data to database: {e}")

    @QtCore.pyqtSlot()
    def Problem_report_Downtime_page(self):
        self.style_button_with_shadow(
            (self.ui.DT_problem_report_btn, self.ui.DT_dashboard_btn, self.ui.DT_data_btn, self.ui.DT_import_data_btn))
        self.ui.DT_stacked_widget.setCurrentWidget(
            self.ui.DT_problem_report_page)
        self.DT_tree = self.ui.DT_report_file_tree
        self.DT_path_model = QtGui.QStandardItemModel(0, 2)
        self.DT_path_model.setHorizontalHeaderLabels(["Select Folder", "Q'ty"])
        self.DT_tree_path_dict = {
            "Current_pos": None,
            "Current_PE": None,
        }
        root_names = [
            "In-Line Incident",
            "Customer Complaint",
            "Safety Accident",
            "MSA Request",
            "Downtime Analysis",
            "4M Change",
            "Other Incident"
        ]
        for name in root_names:
            root_item = QtGui.QStandardItem(name)
            for i in range(1, 6):
                pe_name = f"PE{i}"
                child_col0 = QtGui.QStandardItem(pe_name)
                child_col1 = QtGui.QStandardItem("")
                root_item.appendRow([child_col0, child_col1])
            self.DT_path_model.appendRow([root_item, QtGui.QStandardItem("")])
        self.DT_tree.setModel(self.DT_path_model)
        self.DT_tree.setExpandsOnDoubleClick(True)
        self.DT_tree.setColumnWidth(0, 200)
        self.DT_tree.setColumnWidth(1, self.DT_tree.width() - 205)
        self.DT_tree.setEditTriggers(
            QtWidgets.QAbstractItemView.NoEditTriggers)
        self.safe_connect(self.DT_tree.clicked,
                          self.DT_report_file_tree_clicked)
        self.DT_report_file_model = QtGui.QStandardItemModel(0, 10)
        self.DT_report_file_model.setHorizontalHeaderLabels(
            ["Report ID", "Title", "Department", "Line", "Machine", "Report Type", "Date", "Reported By", "Status", "Notes"])
        self.ui.DT_report_file_table.setModel(self.DT_report_file_model)
        self.ui.DT_report_file_table.setColumnWidth(0, 0)
        self.ui.DT_report_file_table.setColumnWidth(1, 390)
        self.ui.DT_report_file_table.setColumnWidth(2, 80)
        self.ui.DT_report_file_table.setColumnWidth(3, 60)
        self.ui.DT_report_file_table.setColumnWidth(4, 100)
        self.ui.DT_report_file_table.setColumnWidth(5, 100)
        self.ui.DT_report_file_table.setColumnWidth(6, 80)
        self.ui.DT_report_file_table.setColumnWidth(7, 80)
        self.ui.DT_report_file_table.setColumnWidth(8, 60)
        self.ui.DT_report_file_table.setColumnWidth(9, 125)
        self.ui.DT_report_file_table.verticalHeader().hide()
        self.ui.DT_report_file_table.setShowGrid(True)
        self.ui.DT_report_file_table.setGridStyle(QtCore.Qt.SolidLine)
        self.ui.DT_report_file_table.setAlternatingRowColors(True)
        self.ui.DT_report_file_table.verticalHeader().setMinimumSectionSize(40)
        self.ui.DT_report_file_table.setContextMenuPolicy(
            QtCore.Qt.CustomContextMenu)
        self.safe_connect(self.ui.DT_report_file_table.customContextMenuRequested, lambda pos: self.table_context_menu(pos, self.ui.DT_report_file_table, actions=["open", "rename", "update", "delete"],
                                                                                                                       functions_dict={
            "open": lambda idx: self.DT_open_report(idx),
        }))
        self.safe_connect(self.ui.DT_new_report_btn.clicked,
                          self.New_report_input)

    def DT_report_file_tree_clicked(self, index):
        parts = []
        current = index
        while current.isValid():
            parts.insert(0, current.data())
            current = current.parent()
        new_pe = parts[-1] if len(parts) == 2 else None
        if self.DT_tree_path_dict["Current_pos"] == parts[0] and self.DT_tree_path_dict["Current_PE"] == new_pe:
            return
        self.DT_tree_path_dict["Current_pos"] = parts[0]
        self.DT_tree_path_dict["Current_PE"] = new_pe
        filter_scripts = "rt.report_type_name = :report_type AND d.department_name = :group" if new_pe else "rt.report_type_name = :report_type"
        try:
            result = self.database_process.query(sql=f'''
                                                SELECT pr.report_id, pr.report_title, d.department_name, pl.line_name, 
                                                m.machine_code, rt.report_type_name , pr.report_date, pr.reported_by, pr.status,pr.notes,
                                                pr.issue_description, pr.corrective_action,pr.report_file_path, pr.path_type
                                                FROM problem_reports AS pr
                                                LEFT JOIN departments AS d ON pr.department_id = d.department_id
                                                LEFT JOIN production_lines AS pl ON pr.line_id = pl.line_id
                                                LEFT JOIN machines AS m ON pr.machine_id = m.machine_id
                                                LEFT JOIN report_types AS rt ON pr.report_type_id = rt.report_type_id
                                                WHERE {filter_scripts}
                                                ORDER BY pr.report_date DESC;
                ''', params={"report_type": self.DT_tree_path_dict["Current_pos"], "group": self.DT_tree_path_dict["Current_PE"]})
            self.DT_report_file_data_frame = pd.DataFrame(result, columns=["report_id", "report_title", "department_name", "line_name", "machine_code",
                                                          "report_type_name", "report_date", "reported_by", "status", "notes", "issue_description", "corrective_action", "report_file_path", "path_type"])
            self.ui.DT_report_file_table.setUpdatesEnabled(False)
            self.ui.DT_report_file_table.setSortingEnabled(False)
            self.add_data_to_model(data=result, target=self.ui.DT_report_file_table, model=self.DT_report_file_model, column_range=(0, 10),
                                   callback=lambda m=self.DT_report_file_model, d=result: self.icon_from_path(m, d), tooltip_Enable=True)
            self.ui.DT_report_file_table.setWordWrap(True)
            self.ui.DT_report_file_table.resizeRowsToContents()
            self.ui.DT_report_file_table.setUpdatesEnabled(True)
            self.ui.DT_report_file_table.setSortingEnabled(True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load files: {e}")

    def icon_from_path(self, model=None, data=None, file_extension=None):
        icon_provider = QtWidgets.QFileIconProvider()
        ext_map = {
            "PDF":    "file.pdf",
            "DOC":    "file.doc",
            "DOCX":   "file.docx",
            "XLS":    "file.xls",
            "XLSX":   "file.xlsx",
            "XLSM":   "file.xlsm",
            "PPTX":   "file.pptx",
            "JPG":    "file.jpg",
            "JPEG":   "file.jpeg",
            "PNG":    "file.png",
        }
        if file_extension:
            if file_extension in ext_map:
                file_name = ext_map[file_extension]
                icon = icon_provider.icon(QtCore.QFileInfo(file_name))
            else:
                icon = icon_provider.icon(QtWidgets.QFileIconProvider.File)
            return icon
        for row_idx, row_data in enumerate(data):
            path_type = row_data[13]
            file_path = row_data[12]
            if path_type == "FOLDER":
                icon = icon_provider.icon(QtWidgets.QFileIconProvider.Folder)
            elif path_type == "URL":
                icon = icon_provider.icon(QtWidgets.QFileIconProvider.Network)
            elif path_type in ext_map:
                icon = icon_provider.icon(QtCore.QFileInfo(ext_map[path_type]))
            else:
                icon = icon_provider.icon(QtWidgets.QFileIconProvider.File)
            title_item = model.item(row_idx, 1)
            if title_item:
                title_item.setIcon(icon)
    
    @QtCore.pyqtSlot()
    def table_context_menu(self, pos, table, actions, functions_dict=None):
        index = table.indexAt(pos)
        if not index.isValid():
            return
        default_config = {
            "open":   (resource_path("Icons\\Open.ico"),       "Open",        self.DT_open_report),
            "rename": (resource_path("Icons\\rename.ico"),     "Rename",      self.DT_rename_report),
            "update": (resource_path("Icons\\Update_doc.ico"), "Update File", self.DT_update_report),
            "edit":   (resource_path("Icons\\new_register.ico"), "Edit",      None),
            "delete": (resource_path("Icons\\delete.ico"),     "Delete",      self.DT_delete_report),
            "add": (resource_path("Icons\\Add_new.ico"),     "Add",      None)
        }
        if functions_dict:
            for k, fn in functions_dict.items():
                if k in default_config:
                    icon_path, label, _ = default_config[k]
                    default_config[k] = (icon_path, label, fn)

        menu = QtWidgets.QMenu()
        action_map = {}
        for name in actions:
            if name == "separator":
                menu.addSeparator()
                continue
            if name not in default_config:
                continue
            icon_path, label, fn = default_config[name]
            qa = menu.addAction(QtGui.QIcon(icon_path), label)
            if fn:
                action_map[qa] = fn

        chosen = menu.exec_(table.viewport().mapToGlobal(pos))
        if chosen and chosen in action_map:
            action_map[chosen](index)

    def DT_open_report(self, index):
        report_id = int(self.DT_report_file_model.item(index.row(), 0).text())
        report_path = self.DT_report_file_data_frame.loc[self.DT_report_file_data_frame[
            "report_id"] == report_id, "report_file_path"].values[0]
        os.startfile(os.path.normpath(report_path))

    def DT_update_report(self, index):
        print("Update report at row:", index.row())

    def DT_rename_report(self, index):
        print("Rename report at row:", index.row())

    def DT_delete_report(self, index):
        print("Delete report at row:", index.row())

    def New_report_input(self):
        try:
            new_report_dialog = New_Report_Input(
                parent=self, database=self.database_process, callback=self.icon_from_path)
            new_report_dialog.exec_()
            if new_report_dialog.result() == QtWidgets.QDialog.Accepted:
                self.DT_report_file_tree_clicked(
                    self.ui.DT_report_file_tree.currentIndex())
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to create new report: {e}")

# ==========================Function of Downtime page ==================================================================================END
# ==========================Function of Downtime page ==================================================================================END
# ==========================Function of Downtime page ==================================================================================END


class Machine_information(QtWidgets.QDialog):
    def __init__(self, database=None, code=None):
        super().__init__()
        self.database = database
        self.ui = Ui_Machine_detail()
        self.ui.setupUi(self)
        self.ui.mc_code.setText(code)
        try:
            data = self.database.query(sql='''
                                                SELECT m.machine_name,p.line_name,d.department_name,m.maker,m.model,m.function,m.date_receipt,m.machine_status,m.image_link
                                                FROM `Machines` as m
                                                JOIN `Production_Lines` as p
                                                ON m.line_id = p.line_id
                                                JOIN `Departments` as d
                                                ON p.department_id = d.department_id
                                                WHERE m.machine_code = :code;
                                                ''', params={'code': code})
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
            headers = ["Date", "Line", "Result", "Record"]
            monitor_model = QtGui.QStandardItemModel()
            monitor_model.setHorizontalHeaderLabels(headers)
            history = self.database.query(sql='''SELECT mr.maintenance_date, p.line_name,mr.record_link
                                                    FROM `Maintenance_records` as mr
                                                    JOIN `Production_Lines` as p
                                                    ON mr.line_id = p.line_id
                                                    JOIN `Machines` as m
                                                    ON mr.machine_id = m.machine_id 
                                                    WHERE m.machine_code = :code
                                                    ORDER BY mr.maintenance_date DESC;
                                                    ''', params={'code': code})
            for row in history:
                item = [QtGui.QStandardItem(str(row[0])), QtGui.QStandardItem(
                    str(row[1])), QtGui.QStandardItem("OK")]
                monitor_model.appendRow(item)
            self.ui.mc_history.setModel(monitor_model)
            self.ui.mc_history.setColumnWidth(0, 80)
            self.ui.mc_history.setColumnWidth(1, 60)
            self.ui.mc_history.setColumnWidth(2, 50)
            self.ui.mc_history.setColumnWidth(3, 50)
            delegate_btn = ButtonDelegate(buttons=("Link",))
            self.ui.mc_history.setItemDelegateForColumn(3, delegate_btn)
            delegate_btn.ButtonClicked.connect(
                lambda name, idx: self.on_delegate_btn_clicked(name, idx, history))
            self.ui.mc_history.setMouseTracking(True)
            self.ui.mc_history.viewport().setMouseTracking(True)
            self.ui.mc_history.resizeRowsToContents()
            headers2 = ["Code", "Name", "Stock", "Safety"]
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
            self.ui.mc_partlist.setColumnWidth(0, 80)
            self.ui.mc_partlist.setColumnWidth(1, 100)
            self.ui.mc_partlist.setColumnWidth(2, 40)
            self.ui.mc_partlist.setColumnWidth(3, 40)
            self.ui.mc_partlist.resizeRowsToContents()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot()
    def on_delegate_btn_clicked(self, name, index, history):
        row = index.row()
        try:
            self.pdf = pdf_view(history[row][2])
            self.pdf.show()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"File not found: {history[row][2]}")

class Update_machine_info(QtWidgets.QDialog):
    def __init__(self, parent=None, code=None):
        super().__init__()
        self.parent = parent
        self.database = self.parent.database_process
        self.code = code
        self.ui = Ui_Update_machine_info()
        self.ui.setupUi(self)
        self.delete_data = []
        hearders = ["Code", "Machine name", "Group", "Line name", "Maintenance Frequency",
                    "Maker", "Model", "Function", "Date receipt", "Machine status", "Image Link"]
        self.ui.machine_info_table_bf.setRowCount(len(hearders))
        self.ui.machine_info_table_af.setRowCount(len(hearders))
        self.ui.machine_info_table_af.setColumnCount(2)
        self.ui.machine_info_table_bf.setColumnCount(2)
        for r, hearder in enumerate(hearders, start=0):
            item = QtWidgets.QTableWidgetItem(hearder)
            item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.ui.machine_info_table_bf.setItem(
                r, 0, QtWidgets.QTableWidgetItem(hearder))
            self.ui.machine_info_table_af.setItem(r, 0, item)
        self.ui.machine_info_table_bf.verticalHeader().setVisible(False)
        self.ui.machine_info_table_bf.horizontalHeader().setVisible(False)
        self.ui.machine_info_table_bf.setItem(
            0, 1, QtWidgets.QTableWidgetItem(self.code))
        self.ui.machine_info_table_bf.setEditTriggers(
            QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.machine_info_table_bf.horizontalHeader(
        ).setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.machine_info_table_af.setItem(
            0, 1, QtWidgets.QTableWidgetItem(self.code))
        self.ui.machine_info_table_af.horizontalHeader().setVisible(False)
        self.ui.machine_info_table_af.verticalHeader().setVisible(False)
        self.ui.machine_info_table_af.horizontalHeader(
        ).setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.maintenance_plan_table_bf.setEditTriggers(
            QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.maintenance_plan_table_bf.verticalHeader().setVisible(False)
        self.ui.maintenance_plan_table_bf.horizontalHeader(
        ).setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.maintenance_plan_table_af.verticalHeader().setVisible(False)
        self.ui.maintenance_plan_table_af.horizontalHeader(
        ).setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        try:
            self.machine_info = self.database.query(sql='''   SELECT m.machine_name,d.department_name,p.line_name,m.maintenance_frequency,m.maker,m.model,m.function,m.date_receipt,m.machine_status,m.image_link
                                                                FROM `Machines` as m
                                                                JOIN `Production_Lines` as p
                                                                ON m.line_id = p.line_id
                                                                JOIN `Departments` as d
                                                                ON p.department_id = d.department_id
                                                                WHERE m.machine_code = :code;
                                                                ''', params={'code': self.code})
            self.maintenance_plan = self.database.query(sql='''SELECT my.month,mp.week,p.line_name
                                                                FROM `Maintenance_plan` as mp
                                                                JOIN `Production_Lines` as p
                                                                ON mp.line_id = p.line_id
                                                                JOIN `Machines` as m
                                                                ON mp.machine_id = m.machine_id
                                                                JOIN `Months_Years` as my
                                                                ON mp.month_year_id = my.month_year_id
                                                                WHERE m.machine_code = :code AND my.year = :year AND mp.maintenance_date IS NULL AND mp.status is NULL
                                                                GROUP BY my.month;
                                                                ''', params={'code': self.code, 'year': self.parent.year_num})
            self.register_form = self.database.query(sql='''  SELECT mf.form_name,mf.form_link, d.department_name
                                                                FROM `Maintenance_Form_Register` as mfr
                                                                JOIN `Maintenance_form` as mf
                                                                ON mfr.form_id = mf.form_id
                                                                JOIN `Machines` as m
                                                                ON mfr.machine_id = m.machine_id
                                                                JOIN `Departments` as d
                                                                ON mf.department_id = d.department_id
                                                                WHERE m.machine_code = :code;
                                                            ''', params={'code': self.code})
            for r, item in enumerate(self.machine_info[0], start=1):
                self.ui.machine_info_table_bf.setItem(
                    r, 1, QtWidgets.QTableWidgetItem(str(item)))
            self.ui.maintenance_plan_table_bf.setRowCount(
                len(self.maintenance_plan))
            for row in range(len(self.maintenance_plan)):
                for col in range(len(self.maintenance_plan[row])):
                    self.ui.maintenance_plan_table_bf.setItem(
                        row, col, QtWidgets.QTableWidgetItem(str(self.maintenance_plan[row][col])))
            if self.register_form:
                self.ui.form_type_lnedit_bf.setText(
                    f"{self.register_form[0][0]} : {self.register_form[0][2]}")
            self.ui.form_type_lnedit_bf.setEnabled(False)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")
        self.setup_signals()

    def setup_signals(self):
        self.ui.insert_btn.clicked.connect(self.insert_row)
        self.ui.delete_btn.clicked.connect(self.delete_row)
        self.ui.cancel_btn.clicked.connect(self.close)
        self.ui.transfer_btn.clicked.connect(self.transfer_data)
        self.ui.confirm_btn.clicked.connect(self.update_data)
        self.ui.form_type_lnedit_af.textChanged.connect(
            lambda text: self.on_text_changed(text=text))
        self.ui.maintenance_plan_table_af.itemChanged.connect(
            lambda item: self.check_line(item))

    @QtCore.pyqtSlot()
    def insert_row(self):
        current_row = self.ui.maintenance_plan_table_af.rowCount()
        if current_row < 12:
            self.ui.maintenance_plan_table_af.insertRow(current_row)

    @QtCore.pyqtSlot()
    def delete_row(self):
        current_row = self.ui.maintenance_plan_table_af.currentRow()
        try:
            if self.ui.maintenance_plan_table_af.item(current_row, 0) is not None:
                code = self.ui.machine_info_table_bf.item(0, 1).text()
                week = self.ui.maintenance_plan_table_bf.item(
                    current_row, 1).text()
                question = QtWidgets.QMessageBox.question(
                    self, "Delete", f"Are you sure to delete the maintenance plan for the machine '{code}'?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
                if question == QtWidgets.QMessageBox.Yes:
                    self.database.query(sql=''' DELETE mp
                                                    FROM `Maintenance_plan` mp
                                                    JOIN `Machines` m
                                                    ON mp.machine_id = m.machine_id
                                                    JOIN `Production_Lines` p
                                                    ON mp.line_id = p.line_id
                                                    JOIN `Months_Years` AS my
                                                    ON mp.month_year_id = my.month_year_id
                                                    WHERE m.machine_code = :code AND my.year = :year AND mp.week = :week;''',
                                        params={'code': code, 'week': week, 'year': self.parent.year_num})
                    self.maintenance_plan = self.database.query(sql='''SELECT my.month,mp.week,p.line_name
                                                FROM `Maintenance_plan` as mp
                                                JOIN `Production_Lines` as p
                                                ON mp.line_id = p.line_id
                                                JOIN `Machines` as m
                                                ON mp.machine_id = m.machine_id
                                                JOIN `Months_Years` as my
                                                ON mp.month_year_id = my.month_year_id
                                                WHERE m.machine_code = :code AND my.year = :year AND mp.maintenance_date IS NULL
                                                GROUP BY my.month;
                                                ''', params={'code': code, 'year': self.parent.year_num})
                    self.ui.maintenance_plan_table_bf.clearContents()
                    self.ui.maintenance_plan_table_bf.setRowCount(
                        len(self.maintenance_plan))
                    for row in range(len(self.maintenance_plan)):
                        for col in range(len(self.maintenance_plan[row])):
                            self.ui.maintenance_plan_table_bf.setItem(
                                row, col, QtWidgets.QTableWidgetItem(str(self.maintenance_plan[row][col])))
                    self.ui.maintenance_plan_table_af.removeRow(current_row)
        except:
            pass

    @QtCore.pyqtSlot()
    def transfer_data(self):
        self.ui.machine_info_table_af.setItem(
            0, 1, QtWidgets.QTableWidgetItem(self.code))
        for r, item in enumerate(self.machine_info[0], start=1):
            self.ui.machine_info_table_af.setItem(
                r, 1, QtWidgets.QTableWidgetItem(str(item)))
        self.ui.maintenance_plan_table_af.setRowCount(
            len(self.maintenance_plan))
        for row in range(len(self.maintenance_plan)):
            for col in range(len(self.maintenance_plan[row])):
                self.ui.maintenance_plan_table_af.setItem(
                    row, col, QtWidgets.QTableWidgetItem(str(self.maintenance_plan[row][col])))
        if self.register_form:
            self.ui.form_type_lnedit_af.setText(
                f"{self.register_form[0][0]} : {self.register_form[0][2]}")

    @QtCore.pyqtSlot()
    def on_text_changed(self, text):
        try:
            dep = self.ui.machine_info_table_af.item(2, 1).text()
            self.parent.filter_suggestion(self.ui.form_type_lnedit_af, "mf.form_name, d.department_name", "`Maintenance_form` as mf ", f'''JOIN `Departments` as d
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
            new_code = to_null(ui.item(0, 1).text())
            old_code = to_null(self.ui.machine_info_table_bf.item(0, 1).text())
            name = to_null(ui.item(1, 1).text())
            dep = to_null(ui.item(2, 1).text())
            line = to_null(ui.item(3, 1).text())
            freq = to_null(ui.item(4, 1).text())
            maker = to_null(ui.item(5, 1).text())
            model = to_null(ui.item(6, 1).text())
            function = to_null(ui.item(7, 1).text())
            receipt = to_null(ui.item(8, 1).text())
            status = to_null(ui.item(9, 1).text())
            image = to_null(ui.item(10, 1).text())
            iscorrectDep = self.database.query(sql=''' SELECT 1 FROM `Production_Lines` as p
                                                        JOIN `Departments` as d
                                                        ON p.department_id = d.department_id
                                                        WHERE p.line_name = :line AND d.department_name = :dep;
                                                        ''', params={'line': line, 'dep': dep})
            if not iscorrectDep:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Line name not found in your Group")
                return
            form = self.ui.form_type_lnedit_af.text().split(" : ")[0]
            iscorrectForm = self.database.query(sql=''' SELECT 1 FROM `Maintenance_form` as mf
                                                        JOIN `Departments` as d
                                                        ON mf.department_id = d.department_id
                                                        WHERE mf.form_name = :form AND d.department_name = :dep;
                                                        ''', params={'form': form, 'dep': dep})
            if not iscorrectForm:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Form name not found in your Group")
                return
            if dep != self.machine_info[0][1]:
                maintenance_plans = []
                for row in range(self.ui.maintenance_plan_table_af.rowCount()):
                    if self.ui.maintenance_plan_table_af.item(row, 0) is not None:
                        maintenance_plans.append((self.ui.maintenance_plan_table_af.item(row, 0).text(
                        ), self.ui.maintenance_plan_table_af.item(row, 1).text(), self.ui.maintenance_plan_table_af.item(row, 2).text()))
                payload = {'old_code': old_code, 'name': name, 'department': dep, 'line': line, 'freq': freq,
                           'maker': maker, 'model': model, 'function': function, 'receipt': receipt, 'status': status, 'image': image,
                           'maintenance': maintenance_plans, 'form': form}
                receiver_id = self.database.query(sql=''' SELECT u.user_id FROM `Users` as u
                                                            JOIN `Departments` as d
                                                            ON u.department_id = d.department_id
                                                            JOIN `Roles` as r
                                                            ON u.role_id = r.role_id
                                                            WHERE d.department_name = :dep AND r.role_level = "Supervisor";''', params={'dep': dep})
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
                QtWidgets.QMessageBox.information(
                    self, "Request Sent", "Your request has been sent to the Supervisor, please wait for confirmation.")
            else:
                for row in range(self.ui.maintenance_plan_table_af.rowCount()):
                    if row < self.ui.maintenance_plan_table_bf.rowCount():
                        self.database.query(sql='''   UPDATE `Maintenance_plan` as mp
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
                                                    ''', params={'old_code': old_code, 'week': self.ui.maintenance_plan_table_af.item(row, 1).text(),
                                                                 'line': self.ui.maintenance_plan_table_af.item(row, 2).text(), 'old_month': self.ui.maintenance_plan_table_bf.item(row, 0).text(), 'year': self.parent.year_num})
                    else:
                        self.database.query(sql='''   INSERT INTO `Maintenance_plan` 
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
                                                    ''', params={'code': old_code, 'line': self.ui.maintenance_plan_table_af.item(row, 2).text(), 'quarter': (self.parent.ui.company_week_month(self.parent.year_num, int(self.ui.maintenance_plan_table_af.item(row, 1).text())) - 1) // 3 + 1,
                                                                 'week': self.ui.maintenance_plan_table_af.item(row, 1).text(), 'original_week': self.ui.maintenance_plan_table_af.item(row, 1).text(), 'year': self.parent.year_num})
                if self.register_form:
                    self.database.query(sql='''   UPDATE `Maintenance_Form_Register` AS mfr
                                                    SET mfr.form_id = ( SELECT mf.form_id FROM `Maintenance_form` AS mf
                                                                        JOIN `Departments` AS d
                                                                        ON mf.department_id = d.department_id
                                                                        WHERE mf.form_name = :form AND d.department_name = :dep LIMIT 1)
                                                    WHERE mfr.machine_id = ( SELECT m2.machine_id FROM `Machines` AS m2 WHERE m2.machine_code = :code LIMIT 1);
                                        ''', params={'code': old_code, 'form': form, 'dep': dep})
                else:
                    self.database.query(sql='''   INSERT INTO `maintenance_form_register` (machine_id, form_id)
                                                    VALUES ((SELECT m2.machine_id FROM `machines` AS m2 WHERE m2.machine_code = :code LIMIT 1),
                                                            (SELECT mf.form_id FROM `maintenance_form` AS mf
                                                             JOIN `departments` AS d
                                                             ON mf.department_id = d.department_id
                                                             WHERE mf.form_name = :form AND d.department_name = :dep LIMIT 1));
                                        ''', params={'code': old_code, 'form': form, 'dep': dep})

                self.database.query(sql='''   UPDATE `Machines` AS m
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
                                                ''', params={'old_code': old_code, 'new_code': new_code, 'name': name, 'dep': dep,
                                                             'line': line, 'freq': freq, 'maker': maker, 'model': model, 'function': function,
                                                             'receipt': receipt, 'status': status, 'image': image})
                QtWidgets.QMessageBox.information(
                    self, "Update success", "The machine data has been updated successfully.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to update data: {e}")

    @QtCore.pyqtSlot()
    def check_line(self, item):
        dep = self.ui.machine_info_table_af.item(2, 1).text()
        col = item.column()
        row = item.row()
        line = item.text()
        if col == 2:
            isCorrectDep = self.database.query(sql=''' SELECT 1 FROM `Production_Lines` as p
                                                        JOIN `Departments` as d
                                                        ON p.department_id = d.department_id
                                                        WHERE p.line_name = :line AND d.department_name = :dep;
                                                        ''', params={'line': line, 'dep': dep})
            if not isCorrectDep:
                self.ui.maintenance_plan_table_af.itemChanged.disconnect()
                self.ui.maintenance_plan_table_af.setItem(
                    row, col, QtWidgets.QTableWidgetItem(""))
                self.ui.maintenance_plan_table_af.itemChanged.connect(
                    lambda item: self.check_line(item=item))
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Line name not found in your Group")
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
                option.palette.setColor(
                    QtGui.QPalette.Text, QtGui.QColor("#ff0000"))
            elif status == "Upcoming":
                option.palette.setColor(
                    QtGui.QPalette.Text, QtGui.QColor("#29cc00"))
            elif status == "Near due":
                option.palette.setColor(
                    QtGui.QPalette.Text, QtGui.QColor("#b04903"))

class ButtonDelegate(QtWidgets.QStyledItemDelegate):
    ButtonClicked = QtCore.pyqtSignal(str, QtCore.QModelIndex)

    def __init__(self, buttons=("Detail", "Update"),target_indexes=None, parent=None):
        super().__init__(parent)
        self._buttons = {}
        self._hovered = None
        self._button_names = buttons  # tuple các nút muốn tạo
        self._target_indexes = target_indexes
    def _is_target(self, index):
        if self._target_indexes is None:
            return True
        return (index.row(), index.column()) in self._target_indexes 
    
    def paint(self, painter, option, index):
        if not self._is_target(index):
            self._buttons.pop(index, None)
            super().paint(painter, option, index)
            return
        super().paint(painter, option, index)

        rect = option.rect
        count = len(self._button_names)
        if count == 0:
            return

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
        if not self._is_target(index):
            return super().editorEvent(event, model, option, index)

        etype = event.type()
        view = option.widget or self.parent()

        if etype == QtCore.QEvent.MouseMove:
            pos = event.pos()
            btns = self._buttons.get(index, {})
            hit = None
            for name, rect in btns.items():
                if rect.contains(pos):
                    hit = (name, QtCore.QPersistentModelIndex(index))
                    break

            if hit != self._hovered:
                self._hovered = hit
                if view:
                    view.viewport().update()
            return True

        elif etype == QtCore.QEvent.MouseButtonPress:
            return True

        elif etype == QtCore.QEvent.MouseButtonRelease:
            pos = event.pos()
            btns = self._buttons.get(index, {})
            for name, rect in btns.items():
                if rect.contains(pos):
                    self.ButtonClicked.emit(name, QtCore.QModelIndex(index))
                    return True
            return True  

        elif etype in (QtCore.QEvent.Leave, QtCore.QEvent.HoverLeave):
            if self._hovered:
                self._hovered = None
                if view:
                    view.viewport().update()
            return True

        return super().editorEvent(event, model, option, index)

class Print_selector(QtWidgets.QWidget):
    def __init__(self, parent=None, quantity=0, data=None, attached_machine=None, database=None, duplicate=None):
        super().__init__(parent)
        self.ui = Ui_print_selector()
        self.ui.setupUi(self)
        self.setup_signals()
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Window)
        self.printer_process = Printer_process()
        self.printer_name = [name[2]
                             for name in self.printer_process.printers_list]
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
        self.ui.Print_select_cbb.currentIndexChanged.connect(
            self.select_printer)

    @QtCore.pyqtSlot()
    def select_printer(self):
        self.printer_process.choice_printer(
            self.ui.Print_select_cbb.currentText())

    @QtCore.pyqtSlot()
    def start_printer(self):
        if self.quantity <= 0 or self.data == None:
            QtWidgets.QMessageBox.critical(self, "Error", f"Nothing to print")
            return
        try:
            self.worker = WorkerThread(self.print_job, self.data)
            self.progress_window = Printer_progress(
                max=len(self.data), worker=self.worker)
            self.worker.progress_changed.connect(
                lambda value: self.progress_window.update_progress(value=value))
            self.worker.finished.connect(self.progress_window.on_finished)
            self.progress_window.show()
            self.worker.start()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Print Error: {e}")
        self.close()

    def print_job(self, data):
        for i, info in enumerate(data, start=0):
            if info[0] in self.attached_machine:
                self.printer_process.send_to_printer(input_pdf=info[-1], data=[info[0], info[1], info[3], info[4], info[8], str(
                    info[7])], attached_machine=self.attached_machine[info[0]], file_index=i)
            else:
                self.printer_process.send_to_printer(
                    input_pdf=info[-1], data=[info[0], info[1], info[3], info[4], info[8], str(info[7])], file_index=i)
            if i not in self.duplicate:
                try:
                    self.database.query(sql=f''' INSERT INTO `Record_pending` (machine_id,line_id,technical,maintenance_date)
                                                    VALUES ( (SELECT machine_id FROM `Machines` WHERE machine_code = :code),
                                                            (SELECT line_id FROM `Production_Lines` WHERE line_name = :line ),
                                                            :technical , :date );''',
                                        params={'code': info[0], 'line': info[4], 'technical': info[8], 'date': info[7]})
                    if info[0] in self.attached_machine:
                        for attached_code in self.attached_machine[info[0]]:
                            self.database.query(sql=f''' INSERT INTO `Record_pending` (machine_id,line_id,technical,maintenance_date,attached_equipment)
                                                            VALUES ( (SELECT machine_id FROM `Machines` WHERE machine_code = :code),
                                                                (SELECT line_id FROM `Production_Lines` WHERE line_name = :line ),
                                                                :technical , :date ,
                                                                (SELECT machine_id FROM `Machines` WHERE machine_code = :attach_with));''',
                                                params={'code': attached_code, 'line': info[4], 'technical': info[8], 'date': info[7], 'attach_with': info[0]})
                except Exception as e:
                    QtWidgets.QMessageBox.critical(
                        self, "Error",  f"Fail to load : {e}")
                    return
            else:
                try:
                    temp = self.database.query(sql=f'''SELECT rp.rp_id,m.machine_code FROM `Record_pending` as rp
                                                                                        JOIN `Machines` as m
                                                                                        ON rp.machine_id = m.machine_id
                                                                                        WHERE rp.attached_equipment = (SELECT m2.machine_id FROM `Machines` as m2 WHERE m2.machine_code = "{info[0]}")
                                                                                        ORDER BY rp.rp_id ASC;''')
                    if info[0] in self.attached_machine:
                        code_list = self.attached_machine[info[0]]
                        if temp:
                            attach_code_current = [code[1] for code in temp]
                            for code in code_list:
                                if code in attach_code_current:
                                    self.database.query(sql=f'''UPDATE Record_pending
                                                                        SET line_id = ( SELECT line_id 
                                                                        FROM `Production_Lines`
                                                                        WHERE line_name = "{info[4]}"),
                                                                        technical = "{info[8]}",
                                                                        maintenance_date = "{info[7]}"
                                                                        WHERE machine_id = ( SELECT machine_id FROM `Machines` WHERE machine_code = "{info[0]}" );''')

                                else:
                                    self.database.query(sql=f''' INSERT INTO `Record_pending` (machine_id,line_id,technical,maintenance_date,attached_equipment)
                                                                    VALUES ( (SELECT machine_id FROM `Machines` WHERE machine_code = :code),
                                                                    (SELECT line_id FROM `Production_Lines` WHERE line_name = :line ),
                                                                    :technical , :date ,
                                                                    (SELECT machine_id FROM `Machines` WHERE machine_code = :attach_with));''',
                                                        params={'code': code, 'line': info[4], 'technical': info[8], 'date': info[7], 'attach_with': info[0]})
                            delete_code = [
                                c for c in attach_code_current if c not in code_list]
                            if delete_code:
                                delete_ids = ','.join(
                                    f"'{x}'" for x in delete_code)
                                self.database.query(sql=f'''DELETE FROM `Record_pending` 
                                                            WHERE machine_id IN (SELECT machine_id FROM Machines WHERE machine_code IN ({delete_ids}))''')
                        else:
                            for code in code_list:
                                self.database.query(sql=f'''  INSERT INTO `Record_pending` (machine_id,line_id,technical,maintenance_date,attached_equipment)
                                                                VALUES ( (SELECT machine_id FROM `Machines` WHERE machine_code = :code),
                                                                (SELECT line_id FROM `Production_Lines` WHERE line_name = :line ),
                                                                :technical , :date ,
                                                                (SELECT machine_id FROM `Machines` WHERE machine_code = :attach_with));''',
                                                    params={'code': code, 'line': info[4], 'technical': info[8], 'date': info[7], 'attach_with': info[0]})
                    else:
                        if temp:
                            self.database.query(sql=f'''  DELETE FROM `Record_pending` 
                                                            WHERE attached_equipment = (SELECT m.machine_id FROM `Machines` as m WHERE m.machine_code = "{info[0]}");''')
                except Exception as e:
                    QtWidgets.QMessageBox.critical(
                        self, "Error",  f"Fail to load data: {e}")
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
    def __init__(self, parent=None, max=0, text="printed", worker=None):
        super().__init__(parent)
        self.ui = Ui_printing_progress()
        self.ui.setupUi(self)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Window)
        self.ui.printing_progess.setRange(0, max)
        self.ui.printing_progess.setValue(0)
        self.text = text
        self.worker = worker

    @QtCore.pyqtSlot()
    def update_progress(self, value):
        self.ui.printing_progess.setValue(value)

    @QtCore.pyqtSlot()
    def on_finished(self):
        QtWidgets.QMessageBox.information(
            self, "Done", f"All files {self.text}!")
        self.close()

class Form_Modification(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_Form_Modification()
        self.ui.setupUi(self)
        self.pdf_page = Scan_record_process()
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Window)
        self.parent = parent
        self.setup_signals()
        self.result = []
        self.department_maintenance_form = None
        header = ["Machine code", "Machine name"]
        self.ui.apply_machine_table.setColumnCount(len(header))
        self.ui.apply_machine_table.setHorizontalHeaderLabels(header)
        self.ui.apply_machine_table.setAcceptDrops(True)
        self.ui.apply_machine_table.setDragDropMode(
            QtWidgets.QAbstractItemView.DropOnly)
        self.ui.apply_machine_table.dragEnterEvent = self.dragEnterEvent
        self.ui.apply_machine_table.dragMoveEvent = self.dragMoveEvent
        self.ui.apply_machine_table.dropEvent = self.dropEvent
        self.ui.apply_machine_table.setEditTriggers(
            QtWidgets.QAbstractItemView.NoEditTriggers)

    def setup_signals(self):
        self.ui.register_form_btn.clicked.connect(
            lambda _: self.register_form_page())
        self.ui.update_form_btn.clicked.connect(
            lambda _: self.update_form_page())
        self.ui.mod_cancel_btn.clicked.connect(self.close)
        self.ui.mod_insert_btn.clicked.connect(
            lambda _: self.insert_row_case())
        self.ui.mode_delete_btn.clicked.connect(lambda _: self.delete_row())
        self.ui.mod_confirm_btn.clicked.connect(
            lambda _: self.confirm_action())
        self.ui.record_name_lnedit.textChanged.connect(
            lambda text: self.on_text_changed(text, isupdate=True, iscell=False))
        self.ui.load_record_form.clicked.connect(
            lambda _: self.load_record_info())
        self.ui.list_form_btn.clicked.connect(self.list_form_page)
        self.ui.list_form_load_btn.clicked.connect(self.load_form_data)

    @QtCore.pyqtSlot()
    def register_form_page(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.register_form_page)
        self.parent.style_button_with_shadow(button=(
            self.ui.register_form_btn, self.ui.update_form_btn, self.ui.list_form_btn))
        self.ui.apply_machine_table.setEnabled(True)
        self.ui.apply_machine_table.setRowCount(0)
        self.ui.apply_machine_table.clearContents()
        self.ui.label_6.setText("Apply machines")
        self.ui.apply_machine_table.setEditTriggers(
            QtWidgets.QAbstractItemView.AllEditTriggers)
        header = ["Machine code", "Machine name"]
        self.ui.apply_machine_table.setColumnCount(len(header))
        self.ui.apply_machine_table.setHorizontalHeaderLabels(header)
        self.ui.apply_machine_table.setColumnWidth(0, 100)
        self.ui.apply_machine_table.setColumnWidth(1, 310)
        self.ui.frame_14.show()
        if self.ui.register_group_cbb.count() == 0:
            self.ui.register_group_cbb.addItems(
                [item[0] for item in self.parent.group])
            self.ui.register_group_cbb.setCurrentText(
                self.parent.login_info['department'])
        self.ui.register_group_cbb.currentTextChanged.connect(
            lambda text: setattr(self, "department_maintenance_form", text))

    @QtCore.pyqtSlot()
    def update_form_page(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.update_form_page)
        self.parent.style_button_with_shadow(button=(
            self.ui.update_form_btn, self.ui.register_form_btn, self.ui.list_form_btn))
        self.ui.apply_machine_table.clearContents()
        self.ui.apply_machine_table.setRowCount(0)
        self.ui.apply_machine_table.setEnabled(False)
        self.ui.label_6.setText("Apply machines")
        header = ["Machine code", "Machine name"]
        self.ui.apply_machine_table.setColumnCount(len(header))
        self.ui.apply_machine_table.setHorizontalHeaderLabels(header)
        self.ui.apply_machine_table.setColumnWidth(0, 100)
        self.ui.apply_machine_table.setColumnWidth(1, 310)
        self.ui.frame_14.show()
        self.ui.apply_machine_table.setEditTriggers(
            QtWidgets.QAbstractItemView.AllEditTriggers)
        self.ui.update_choice.setChecked(True)

    @QtCore.pyqtSlot()
    def insert_row_case(self):
        if self.ui.stackedWidget.currentWidget() == self.ui.register_form_page:
            self.insert_row()
        else:
            self.insert_row(isupdate=True)

    def insert_row(self, isupdate=False):
        current_row = self.ui.apply_machine_table.rowCount()
        self.ui.apply_machine_table.insertRow(current_row)
        self.ui.apply_machine_table.scrollTo(
            self.ui.apply_machine_table.model().index(current_row, 0))
        editor = QtWidgets.QLineEdit()
        editor.setStyleSheet(''' border: none;''')
        editor.textChanged.connect(
            lambda text, r=current_row, c=0: self.on_text_changed(
                text, c, r, isupdate)
        )
        editor.editingFinished.connect(
            lambda r=current_row: self.load_data(r, isupdate))
        self.ui.apply_machine_table.setCellWidget(current_row, 0, editor)

    @QtCore.pyqtSlot()
    def delete_row(self, r=None):
        if r is None:
            current_row = self.ui.apply_machine_table.currentRow()
            self.ui.apply_machine_table.removeRow(
                self.ui.apply_machine_table.currentRow())

    @QtCore.pyqtSlot()
    def on_text_changed(self, text, c=0, r=0, isupdate=False, iscell=True):
        try:
            if not isupdate:
                self.parent.filter_suggestion(self.ui.apply_machine_table.cellWidget(r, c), "m.machine_code", "`Machines` as m ", f'''JOIN `Production_Lines` as p
                                                                                                                            ON p.line_id = m.line_id
                                                                                                                            JOIN `Departments` as d
                                                                                                                            ON d.department_id = p.department_id
                                                                                                                            LEFT JOIN `Maintenance_Form_Register` as mfr
                                                                                                                            ON m.machine_id = mfr.machine_id
                                                                                                                            WHERE d.department_name = "{self.ui.register_group_cbb.currentText()}" AND mfr.machine_id IS NULL AND machine_code LIKE "%{text}%"
                                                                                                                            ''')
            else:
                if iscell:
                    self.parent.filter_suggestion(self.ui.apply_machine_table.cellWidget(r, c), "m.machine_code", "`Machines` as m ", f'''JOIN `Production_Lines` as p
                                                                                                                                ON p.line_id = m.line_id
                                                                                                                                JOIN `Departments` as d
                                                                                                                                ON d.department_id = p.department_id
                                                                                                                                LEFT JOIN `Maintenance_Form_Register` as mfr
                                                                                                                                ON m.machine_id = mfr.machine_id
                                                                                                                                WHERE d.department_name = "{self.department_update.strip()}" AND mfr.machine_id IS NULL AND machine_code LIKE "%{text}%"''')
                else:
                    self.parent.filter_suggestion(self.ui.record_name_lnedit, "mf.form_name, d.department_name", "`Maintenance_form` as mf ", f'''  JOIN `Departments` as d
                                                                                                                                                            ON d.department_id = mf.department_id
                                                                                                                                                            WHERE mf.form_name LIKE "%{text}%"
                                                                                                                                                            ''')
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot()
    def load_data(self, r, isupdate):
        try:
            code = self.ui.apply_machine_table.cellWidget(r, 0).text()
            if not isupdate:
                result = self.parent.database_process.query(sql='''SELECT m.machine_name 
                                                                    FROM `Machines` as m
                                                                    JOIN `Production_Lines` as p
                                                                    ON m.line_id = p.line_id
                                                                    JOIN `Departments` as d
                                                                    ON p.department_id = d.department_id
                                                                    WHERE m.machine_code = :code AND d.department_name = :dep;''', params={'code': code, 'dep': self.ui.register_group_cbb.currentText()})
                if r > (len(self.result) - 1):
                    self.result.append(f"'{code}'")
                else:
                    self.result[r] = f"'{code}'"
            else:
                result = self.parent.database_process.query(sql='''SELECT m.machine_name 
                                                                    FROM `Machines` as m
                                                                    JOIN `Production_Lines` as p
                                                                    ON m.line_id = p.line_id
                                                                    JOIN `Departments` as d
                                                                    ON p.department_id = d.department_id
                                                                    WHERE m.machine_code = :code AND d.department_name = :dep;''', params={'code': code, 'dep': self.department_update.strip()})
            if result:
                self.ui.apply_machine_table.setItem(
                    r, 1, QtWidgets.QTableWidgetItem(f"{result[0][0]}"))
            else:
                raise ValueError(
                    f"Machine code {code} not found in your department or has been register for other form")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data : {e}")
            # if self.ui.apply_machine_table.cellWidget(r,0):
            #     editor = self.ui.apply_machine_table.cellWidget(r,0)
            #     editor.blockSignals(True)
            #     editor.setText("")
            #     editor.blockSignals(False)
            #     self.ui.apply_machine_table.setItem(r,1,QtWidgets.QTableWidgetItem(""))

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
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Failed to load data: {e}")
                return
            if df.empty:
                QtWidgets.QMessageBox.critical(self, "Error", f"File is empty")
                return
            else:
                temp = self.parent.database_process.query(sql=''' SELECT m.machine_code 
                                                                    FROM `Machines` as m
                                                                    JOIN `Production_Lines` as p
                                                                    ON m.line_id = p.line_id
                                                                    JOIN `Departments` as d
                                                                    ON p.department_id = d.department_id
                                                                    LEFT JOIN `Maintenance_Form_Register` as mfr
                                                                    ON m.machine_id = mfr.machine_id
                                                                    WHERE d.department_name = :dep AND mfr.machine_id IS NULL; ''', params={'dep': self.ui.register_group_cbb.currentText()})
                non_register_machine = [machine[0] for machine in temp]
                machine_code = df.iloc[:, 0]
                machine_hasbeen_register = []
                for r, value in enumerate(machine_code, start=0):
                    if value in non_register_machine:
                        self.insert_row()
                        self.ui.apply_machine_table.cellWidget(
                            r, 0).setText(str(value))
                        self.load_data(r=r, isupdate=False)
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
        if (text == "") or (text is None):
            return
        self.ui.apply_machine_table.setEnabled(True)
        try:
            self.department_update = text.split(":", 1)[1]
            text = text.split(":", 1)[0]
            self.department_maintenance_form = self.department_update.strip()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Invalid form name")
            return
        try:
            self.record_info = self.parent.database_process.query(sql='''SELECT m.machine_code, m.machine_name,mfr.register_id
                                                        FROM `Machines` as m
                                                        JOIN `Maintenance_Form_Register` as mfr
                                                        ON m.machine_id = mfr.machine_id
                                                        JOIN `Maintenance_form` as mf
                                                        ON mfr.form_id = mf.form_id
                                                        WHERE mf.form_name =  :form_name;''', params={'form_name': text})
            self.form_info = self.parent.database_process.query(sql=''' SELECT form_link,form_id
                                                                    FROM `maintenance_form`    
                                                                    WHERE form_name =  :form_name;''', params={'form_name': text})
            num_machine = len(self.record_info)
            self.machines_registered = [machine[0]
                                        for machine in self.record_info]
            self.ui.apply_machine_table.clearContents()
            self.ui.apply_machine_table.setRowCount(0)
            self.ui.update_form_link.setText(self.form_info[0][0])
            if num_machine == 0:
                raise ValueError(
                    f"Not see any machine has been registered for the form {text}")
            else:
                for r in range(num_machine):
                    self.insert_row(isupdate=True)
                    editor = self.ui.apply_machine_table.cellWidget(r, 0)
                    editor.setText(f"{self.record_info[r][0]}")
                    self.ui.apply_machine_table.setItem(
                        r, 1, QtWidgets.QTableWidgetItem(f"{self.record_info[r][1]}"))
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    @QtCore.pyqtSlot()
    def confirm_action(self):
        if self.parent.login_info["role_level"] in ["Manager", "Admin"]:
            pass
        elif (self.parent.login_info["department"] == self.department_maintenance_form) and (self.parent.login_info["role_level"] in ["Supervisor"]):
            pass
        else:
            QtWidgets.QMessageBox.information(
                self, "Permission denied", "Your don't have permission to update this machine info")
            return
        if self.ui.stackedWidget.currentWidget() == self.ui.list_form_page:
            return
        if self.ui.apply_machine_table.rowCount() == 0:
            return
        new_codes = []
        for r in range(self.ui.apply_machine_table.rowCount()):
            editor = self.ui.apply_machine_table.cellWidget(r, 0)
            new_codes.append(editor.text())
        if len(new_codes) != len(set(new_codes)):
            QtWidgets.QMessageBox.critical(
                self, "Error", f"There are duplicate machine codes in the table")
            return
        if self.ui.stackedWidget.currentWidget() == self.ui.register_form_page:
            values = ",".join(self.result)
            try:
                form_name = self.ui.register_form_name.text()
                form_link = self.ui.register_form_link.text()
                page_num = self.pdf_page.return_form_page(form_link)
                if not form_name or not form_link or not form_link.lower().endswith(".pdf"):
                    QtWidgets.QMessageBox.warning(
                        self, "Error", "Please enter a valid form name and .pdf path")
                    return
                department_id = self.parent.database_process.query(sql=''' SELECT department_id FROM `Departments` WHERE department_name = :dep''', params={
                                                                   'dep': self.ui.register_group_cbb.currentText()})
                self.parent.database_process.query(
                    sql='''INSERT INTO `Maintenance_form` (form_name, form_link, department_id, page_num)
                            VALUES (:form_name, :form_link, :department_id , :num)''',
                    params={
                        'form_name': form_name,
                        'form_link': form_link,
                        'department_id': department_id[0][0],
                        'num': page_num
                    }
                )
                self.parent.database_process.query(sql=f''' INSERT INTO `Maintenance_Form_Register` (machine_id, form_id)
                                                                SELECT m.machine_id, f.form_id
                                                                FROM `Machines` AS m
                                                                JOIN `Maintenance_form` AS f 
                                                                ON f.form_name = :form_name
                                                                WHERE m.machine_code IN ({values});''', params={'form_name': self.ui.register_form_name.text()})
            except Exception as e:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Failed to register form: {e}")
                return
        elif self.ui.stackedWidget.currentWidget() == self.ui.update_form_page:
            if self.ui.update_choice.isChecked():
                try:
                    text = self.ui.record_name_lnedit.text()
                    form_name = text.split(":", 1)[0]
                    form_link = self.ui.update_form_link.text()
                    self.parent.database_process.query(sql=f''' UPDATE `Maintenance_form`
                                                                    SET form_name = :form_name , form_link = :form_link , page_num = :num
                                                                    WHERE form_id = :form_id
                                                                ''', params={'form_name': form_name, 'form_link': form_link, 'num': self.pdf_page.return_form_page(form_link), 'form_id': self.form_info[0][1]})
                    # delete
                    if len(self.record_info) > self.ui.apply_machine_table.rowCount():
                        old_codes = [row[0] for row in self.record_info]
                        codes_to_delete = set(old_codes) - set(new_codes)
                        for old_code in codes_to_delete:
                            self.parent.database_process.query(
                                sql=''' DELETE FROM `Maintenance_Form_Register`
                                        WHERE form_id = :form_id
                                        AND machine_id = (SELECT machine_id FROM `Machines` WHERE machine_code = :code) ''',
                                params={
                                    'form_id': self.form_info[0][1],
                                    'code': old_code
                                }
                            )
                        QtWidgets.QMessageBox.information(
                            self, "Action complete", "Delete maintenance form complete")
                        return
                    # update and insert
                    for r in range(self.ui.apply_machine_table.rowCount()):
                        editor = self.ui.apply_machine_table.cellWidget(r, 0)
                        new_code = editor.text()
                        # insert
                        if r >= len(self.record_info):
                            self.parent.database_process.query(sql=f''' INSERT INTO `Maintenance_Form_Register` (machine_id, form_id)
                                                                        SELECT m.machine_id, f.form_id
                                                                        FROM `Machines` AS m
                                                                        JOIN `Maintenance_form` AS f 
                                                                        ON f.form_id = :form_id
                                                                        WHERE m.machine_code = :code;''', params={'form_id': self.form_info[0][1], 'code': new_code})
                            continue
                        # update
                        if self.record_info[r][0] != new_code:
                            self.parent.database_process.query(sql=''' UPDATE `Maintenance_Form_Register`
                                                                            SET machine_id = ( SELECT machine_id FROM `Machines` WHERE `machine_code` = :code)
                                                                            WHERE register_id = :register_id ''', params={'code': new_code, 'register_id': self.record_info[r][-1]})

                except Exception as e:
                    QtWidgets.QMessageBox.critical(
                        self, "Error", f"Failed to update data: {e}")
                    return
            else:
                try:
                    self.parent.database_process.query(sql=''' DELETE FROM `Maintenance_form` 
                                                                    WHERE form_id = :form_id''', params={'form_id': self.record_info[0][3]})
                except Exception as e:
                    QtWidgets.QMessageBox.critical(
                        self, "Error", f"3Failed to update data: {e}")
                    return
        QtWidgets.QMessageBox.information(
            self, "Action complete", "Update maintenance form complete")

    @QtCore.pyqtSlot()
    def list_form_page(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.list_form_page)
        self.parent.style_button_with_shadow(button=(
            self.ui.list_form_btn, self.ui.update_form_btn, self.ui.register_form_btn))
        self.ui.apply_machine_table.clearContents()
        self.ui.apply_machine_table.setRowCount(0)
        self.ui.label_6.setText("List of record form")
        self.ui.apply_machine_table.setEnabled(True)
        self.ui.apply_machine_table.setEditTriggers(
            QtWidgets.QAbstractItemView.NoEditTriggers)
        header = ["Form name", "Machine apply\nQ'ty"]
        self.ui.apply_machine_table.setColumnCount(len(header))
        self.ui.apply_machine_table.setHorizontalHeaderLabels(header)
        self.ui.apply_machine_table.setColumnWidth(1, 100)
        self.ui.apply_machine_table.setColumnWidth(0, 280)
        self.ui.frame_14.hide()
        if self.ui.list_form_group_cbb.count() == 0:
            self.ui.list_form_group_cbb.addItems(
                [item[0] for item in self.parent.group])
            self.ui.list_form_group_cbb.setCurrentText(
                self.parent.login_info['department'])
        self.ui.list_form_group_cbb.currentTextChanged.connect(
            lambda text: setattr(self, "department_maintenance_form", text))

    @QtCore.pyqtSlot()
    def load_form_data(self):
        dep = self.ui.list_form_group_cbb.currentText()
        self.ui.apply_machine_table.clearContents()
        self.ui.apply_machine_table.setRowCount(0)
        try:
            result = self.parent.database_process.query(sql='''SELECT mf.form_name,COUNT(mfr.machine_id)
                                                                FROM `maintenance_form` as mf
                                                                JOIN `maintenance_form_register` as mfr
                                                                ON mf.form_id = mfr.form_id
                                                                JOIN `departments` as d
                                                                ON mf.department_id = d.department_id
                                                                WHERE d.department_name = :dep
                                                                GROUP BY mf.form_name
                                                        ''', params={'dep': dep})
            total_machine = 0
            if result:
                self.ui.apply_machine_table.setRowCount(len(result))
                self.ui.form_quantity_lbl.setText(f"{len(result)}")
                for r in range(len(result)):
                    self.ui.apply_machine_table.setItem(
                        r, 0, QtWidgets.QTableWidgetItem(f"{result[r][0]}"))
                    self.ui.apply_machine_table.setItem(
                        r, 1, QtWidgets.QTableWidgetItem(f"{result[r][1]}"))
                    total_machine += result[r][1]
                self.ui.form_list_machine_qty_lbl.setText(f"{total_machine}")
            else:
                self.ui.form_list_machine_qty_lbl.setText("0")
                self.ui.form_quantity_lbl.setText("0")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data: {e}")

    def closeEvent(self, event):
        self.ui.apply_machine_table.clearContents()
        super().close()
        self.deleteLater()

class Login_Dialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
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
            username = "misa"
            # password = ""
            self.ui.login_btn.setEnabled(False)
            QtWidgets.QApplication.processEvents()
            result = self.database.query(sql=''' SELECT s.user_id,s.username,s.password_hash,r.role_level,d.department_name,s.first_name,s.last_name FROM `Users` as s 
                                                    LEFT JOIN `Departments` as d
                                                    ON s.department_id = d.department_id
                                                    JOIN `Roles` as r
                                                    ON s.role_id = r.role_id
                                                    WHERE username = :username ''', params={'username': username})
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
                    QtWidgets.QMessageBox.warning(
                        self, "Warning", f"Failed to set MySQL session user: {e}")
                    return
                self.accept()
            else:
                self.ui.pass_status.setText("❌ Wrong password")
                self.ui.pass_status.setStyleSheet("color: red;")

        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to login: {e}")
        finally:
            self.ui.login_btn.setEnabled(True)

class NotificationItem(QtWidgets.QWidget):
    def __init__(self, notification_content, parent=None, isYours=False):
        super().__init__()
        self.notification_content = notification_content
        self.title = self.notification_content[4]
        self.message = self.notification_content[5]
        self.status = self.notification_content[7]
        self.comment = self.notification_content[15]
        self.parent = parent
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        lbl_title = QtWidgets.QLabel(
            f"<h3 style='color:#ff6600;'>{self.title}</h3>")
        lbl_title.setWordWrap(True)

        lbl_message = QtWidgets.QLabel(self.message)
        lbl_message.setWordWrap(True)
        lbl_message.setStyleSheet("color: gray; font-size: 12px;")

        receive_at = self.notification_content[10].strftime(
            "%Y-%m-%d %H:%M:%S")
        lbl_time = QtWidgets.QLabel(
            f"<i style='color: gray; font-size: 10px;'>Received at: {receive_at}</i>")
        lbl_time.setWordWrap(True)
        lbl_status = QtWidgets.QLabel(
            f"<b style='color: gray; font-size: 12px;'>Status: {self.status}</b>")
        frame = QtWidgets.QFrame()
        frame.setObjectName("frame")
        frame.setMinimumWidth(200)
        frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        frame.setFrameShadow(QtWidgets.QFrame.Raised)

        horizontalLayout = QtWidgets.QHBoxLayout(frame)
        horizontalLayout.setSpacing(10)
        horizontalLayout.setContentsMargins(0, 0, 0, 0)

        btn = QtWidgets.QPushButton("Details")
        btn.setStyleSheet(
            "padding: 4px 10px; background-color: #ddd; border-radius: 5px;")
        btn.clicked.connect(lambda: self.show_details(
            self.notification_content, isYours))
        btn2 = QtWidgets.QPushButton("Cancel")
        btn2.setStyleSheet(
            "padding: 4px 10px; background-color: #ddd; border-radius: 5px;")
        btn2.clicked.connect(lambda: self.cancel_request(
            self.notification_content, isYours) if isYours == True else self.reject_action())
        if (isYours) and (self.status == "ACCEPTED" or self.status == "REJECTED"):
            btn2.setText("Close")
        horizontalLayout.addWidget(btn)
        horizontalLayout.addWidget(btn2)

        layout.addWidget(lbl_title)
        layout.addWidget(lbl_message)
        layout.addWidget(lbl_time)
        layout.addWidget(lbl_status)
        if self.comment:
            lbl_comment = QtWidgets.QLabel(
                f"<b style='color: gray; font-size: 12px;'>Reason: {self.comment}</b>")
            layout.addWidget(lbl_comment)
        layout.addWidget(frame, alignment=QtCore.Qt.AlignRight)

    @QtCore.pyqtSlot()
    def show_details(self, data, isYours):
        if not isYours:
            self.parent.database_process.query(sql=''' UPDATE `Notifications`
                                                SET status = 'READ'
                                                WHERE notification_id = :nid ''', params={'nid': data[0]})
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
        btn_close.setStyleSheet(
            "padding: 6px 12px; background-color: #007acc; color: white; border-radius: 5px;")
        if not isYours:
            btn_accept = QtWidgets.QPushButton("Accept")
            btn_accept.clicked.connect(self.accept_action)
            btn_accept.setStyleSheet(
                "padding: 6px 12px; background-color: #28a745; color: white; border-radius: 5px;")
            btn_reject = QtWidgets.QPushButton("Reject")
            btn_reject.setStyleSheet(
                "padding: 6px 12px; background-color: #dc3545; color: white; border-radius: 5px;")
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
    def cancel_request(self, data, isYours):
        if not isYours:
            return
        if self.status not in ["ACCEPTED", "REJECTED"]:
            confirm = QtWidgets.QMessageBox.question(
                self, "Confirm", "Are you sure you want to cancel this request?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if confirm == QtWidgets.QMessageBox.Yes:
                try:
                    self.parent.database_process.query(sql='''DELETE FROM `Notifications`
                                                                WHERE notification_id = :nid ''', params={'nid': data[0]})
                    self.setParent(None)
                    self.deleteLater()
                except Exception as e:
                    QtWidgets.QMessageBox.critical(
                        self, "Error", f"Action Error: {e}")
        else:
            self.parent.database_process.query(sql='''   UPDATE `Notifications` 
                                                            SET lifecycle_status = "CLOSED"
                                                            WHERE notification_id = :nid ''', params={'nid': data[0]})
        self.parent.Home_page()
        self.close()

    @QtCore.pyqtSlot()
    def accept_action(self):
        try:
            isNotification = self.parent.database_process.query(
                sql=''' SELECT * FROM `Notifications` WHERE notification_id = :nid AND lifecycle_status = "PENDING"''', params={'nid': self.notification_content[0]})
            if isNotification:
                if self.notification_content[1] == "update_machine":
                    content = json.loads(self.notification_content[6])
                    self.parent.database_process.query(sql='''UPDATE `Machines` AS m
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
                                                       params={'code': content.get('old_code'),
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
                            QtWidgets.QMessageBox.critical(
                                self, "Error", f"Invalid month: {month}")
                            return
                        params_list.append({
                            'code': content.get('old_code'),
                            'line': content.get('line'),
                            'quarter': quarter,
                            'week': week,
                            'original_week': original_week,
                            'year': self.parent.year_num
                        })
                    if params_list:
                        self.parent.database_process.executemany(sql=''' INSERT INTO `Maintenance_plan` 
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
                                                                original_week = VALUES(original_week);''', params_list=params_list)
                self.parent.database_process.query(sql='''UPDATE `Notifications`
                                                            SET status = 'ACCEPTED'
                                                            WHERE notification_id = :nid ''', params={'nid': self.notification_content[0]})
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Action Error: {e}")
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
                    sql='''
                        UPDATE `Notifications`
                        SET status = 'REJECTED', comment = :reason
                        WHERE notification_id = :nid
                    ''',
                    params={
                        'nid': self.notification_content[0],
                        'reason': reason
                    }
                )
                QtWidgets.QMessageBox.information(
                    self, "Success", "Request rejected successfully.")
            except Exception as e:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Action Error: {e}")
                return
            self.parent.Home_page()
            self.close()

    def json_to_html_table(self, d: dict) -> str:
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

    def __init__(self, parent=None, line_name="None", data_list=[]):
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
            QtWidgets.QMessageBox.information(
                self, "Success", "Missing data synchronized successfully.")
            self.close()

        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to sync data: {e}")

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self._drag_pos = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == QtCore.Qt.LeftButton and self._drag_pos is not None:
            self.move(event.globalPos() - self._drag_pos)
            event.accept()

    def mouseReleaseEvent(self, event):
        self._drag_pos = None

class Group_Area_Choose(QtWidgets.QDialog):
    def __init__(self, parent=None, database=None, file_path=None):
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
            result = self.database.query(sql=''' SELECT downtime_area_name 
                                                    FROM `downtime_areas` as da
                                                    JOIN `departments` as d
                                                    ON da.department_id = d.department_id
                                                    WHERE d.department_name = :group ''', params={'group': group})
            self.ui.DT_area_input_data.clear()
            if result:
                self.ui.DT_area_input_data.addItems([row[0] for row in result])
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load areas: {e}")

    def load_groups(self):
        try:
            result = self.database.query(sql=''' SELECT d.department_name 
                                                    FROM `downtime_areas` as da
                                                    JOIN `departments` as d
                                                    ON da.department_id = d.department_id
                                                    ORDER BY d.department_name ASC ''')
            self.ui.DT_group_input_data.clear()
            if result:
                self.ui.DT_group_input_data.addItems(
                    [row[0] for row in result])
                excel_sheet = pd.ExcelFile(
                    self.file_path).sheet_names if self.file_path else None
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
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load groups: {e}")
            return

    @QtCore.pyqtSlot()
    def confirm_selection(self):
        group = self.ui.DT_group_input_data.currentText()
        area = self.ui.DT_area_input_data.currentText()
        sheet_name = self.ui.DT_sheet_name.currentText()
        if not group or not area:
            QtWidgets.QMessageBox.warning(
                self, "Warning", "Please select both Group and Area.")
            return
        self.selected_group = group
        self.selected_area = area
        self.excel_sheet_name = sheet_name
        self.accept()

class Downtime_Input(QtWidgets.QDialog):
    def __init__(self, parent=None, database=None, data_frame=None, error_frame=None, area_name=None, month_year=None):
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
        self.ui.data_table.setEditTriggers(
            QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.data_table.setSortingEnabled(False)
        self.ui.data_table.setUpdatesEnabled(False)
        self.ui.error_row_table.setSortingEnabled(False)
        self.ui.error_row_table.setUpdatesEnabled(False)
        headers = ["Date", "Line", "Start\nTime", "Technical\nStart", "Finish\nTime",
                   "Total Loss\nTime", "Wait\nTechnical", "Technical\nName", "Failure\nCode", "Machine Code"]
        self.data_model = QtGui.QStandardItemModel()
        self.data_model.setHorizontalHeaderLabels(headers)
        headers = ["Date", "Line", "Start\nTime", "Technical\nStart", "Finish\nTime",
                   "Technical\nName", "Failure\nCode", "Machine Code", "Column Error", "Error Message", "Action"]
        self.error_model = QtGui.QStandardItemModel()
        self.error_model.setHorizontalHeaderLabels(headers)
        try:
            self.ui.total_row_lbl.setText(
                str(self.data_frame.shape[0]) if self.data_frame is not None else "0")
            self.ui.total_downtime_lbl.setText(str(
                self.data_frame["total_loss_time"].sum()) if self.data_frame is not None else "0")
            if self.data_frame.empty:
                return
            else:
                for row in range(self.data_frame.shape[0]):
                    items = []
                    for col in range(self.data_frame.shape[1]):
                        value = self.data_frame.iat[row, col]
                        if col == 0:
                            value = int(value) if not pd.isna(value) else ""
                        item = QtGui.QStandardItem(
                            str(value) if value is not None and str(value) != "NaT" else "")
                        item.setTextAlignment(QtCore.Qt.AlignCenter)
                        item.setEditable(False)
                        items.append(item)
                    self.data_model.appendRow(items)
                self.ui.data_table.setModel(self.data_model)
            if self.error_frame.empty:
                pass
            else:
                for row in range(self.error_frame.shape[0]):
                    items = []
                    for col in range(self.error_frame.shape[1]):
                        value = self.error_frame.iat[row, col]
                        if col == 0:
                            value = int(value) if not pd.isna(value) else ""
                        item = QtGui.QStandardItem(
                            str(value) if value is not None and str(value) != "NaT" else "")
                        item.setTextAlignment(QtCore.Qt.AlignCenter)
                        item.setEditable(False)
                        items.append(item)
                    self.error_model.appendRow(items)
                self.ui.error_row_table.setModel(self.error_model)
            delegate_btn = ButtonDelegate(buttons=("Edit", "Delete"))
            self.ui.error_row_table.setItemDelegateForColumn(10, delegate_btn)
            self.parent.safe_connect(
                delegate_btn.ButtonClicked, lambda name, idx: self.on_delegate_btn_clicked(name, idx))
            self.ui.error_row_table.setMouseTracking(True)
            self.ui.error_row_table.viewport().setMouseTracking(True)
            self.ui.data_table.setSortingEnabled(True)
            self.ui.data_table.resizeColumnsToContents()
            self.ui.data_table.setUpdatesEnabled(True)
            self.ui.error_row_table.setSortingEnabled(True)
            self.ui.error_row_table.resizeColumnsToContents()
            self.ui.error_row_table.setColumnWidth(10, 100)
            self.ui.error_row_table.setUpdatesEnabled(True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load data into tables: {e}")

    def confirm_data(self):
        self.accept()

    def on_delegate_btn_clicked(self, name, idx):
        row = idx.row()
        if name == "Edit":
            self.new_data = pd.DataFrame(columns=self.data_frame.columns)
            try:
                row_data = {}
                for col in range(self.error_model.columnCount() - 1):
                    item = self.error_model.item(row, col)
                    if item:
                        row_data[self.error_model.headerData(
                            col, QtCore.Qt.Horizontal)] = item.text()

                edit_dialog = QtWidgets.QDialog(self)
                edit_dialog.setWindowTitle("Edit Row Data")
                edit_dialog.setMinimumWidth(500)

                layout = QtWidgets.QVBoxLayout(edit_dialog)

                form_fields = {}
                for col in range(self.error_model.columnCount() - 1):
                    header = self.error_model.headerData(
                        col, QtCore.Qt.Horizontal)

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
                    pd.concat([self.new_data, pd.DataFrame([row_data])],
                              ignore_index=True, sort=False)
                    for col, (header, line_edit) in enumerate(form_fields.items()):
                        if col < 8:
                            new_row.append(line_edit.text())
                    try:
                        start_time = pd.to_datetime(new_row[2], format='%H:%M')
                        print(f"Start Time: {start_time}")
                        technical_start = pd.to_datetime(
                            new_row[3], format='%H:%M')
                        print(f"Technical Start Time: {technical_start}")
                        finish_time = pd.to_datetime(
                            new_row[4], format='%H:%M')
                        print(f"Finish Time: {finish_time}")

                        total_loss_time = int(
                            (finish_time - start_time).total_seconds() / 60)
                        wait_technical_time = int(
                            (technical_start - start_time).total_seconds() / 60)
                    except Exception:
                        QtWidgets.QMessageBox.warning(
                            self, "Warning", "Invalid format. Please check again.")
                        return
                    new_row.insert(5, total_loss_time)
                    new_row.insert(6, wait_technical_time)
                    self.new_data.loc[len(self.new_data)] = new_row
                    edit_dialog.accept()
                confirm_btn.clicked.connect(save_changes)
                cancel_btn.clicked.connect(edit_dialog.reject)
                edit_dialog.exec_()
            except Exception as e:
                QtWidgets.QMessageBox.critical(
                    self, "Error", f"Failed to edit row: {e}")
            if edit_dialog.result() == QtWidgets.QDialog.Accepted:
                if not self.new_data.empty:
                    self.data_frame = pd.concat(
                        [self.data_frame, self.new_data], ignore_index=True, sort=False)
                    self.error_frame.drop(
                        self.error_frame.index[row], inplace=True)
                    for r in range(self.new_data.shape[0]):
                        items = []
                        for c in range(self.new_data.shape[1]):
                            value = self.new_data.iat[r, c]
                            item = QtGui.QStandardItem(
                                str(value) if value is not None and str(value) != "NaT" else "")
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                            item.setEditable(False)
                            items.append(item)
                        self.data_model.appendRow(items)
                    self.error_model.removeRow(row)
                    self.ui.total_row_lbl.setText(
                        str(self.data_frame.shape[0]))
                    self.ui.total_downtime_lbl.setText(
                        str(self.data_frame["total_loss_time"].sum()))
        elif name == "Delete":
            self.ui.error_row_table.model().removeRow(row)

class Error_code_management(QtWidgets.QDialog):
    def __init__(self, parent=None, database=None):
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
            self.ui.error_list_table.setEditTriggers(
                QtWidgets.QAbstractItemView.NoEditTriggers)
            self.ui.error_list_table.setSortingEnabled(False)
            self.ui.error_list_table.setUpdatesEnabled(False)
            result = self.database.query(sql=''' SELECT ecl.error_code, ecl.error_description, ecl.reason, ecl.recommended_action, ecl.process, da.downtime_area_name
                                                    FROM `error_codes_list` ecl
                                                    JOIN `downtime_areas` da ON ecl.downtime_area_id = da.downtime_area_id
                                                    ORDER BY ecl.error_code ASC, ecl.process COLLATE utf8mb4_unicode_ci ASC;''')
            headers = ["Error Code", "Description", "Reason",
                       "Recommended\nAction", "Process", "Downtime\nArea"]
            self.error_model = QtGui.QStandardItemModel()
            self.error_model.setHorizontalHeaderLabels(headers)
            self.add_item_to_error_list(result, self.error_model)
            self.ui.error_list_table.setModel(self.error_model)
            self.ui.error_list_table.resizeColumnsToContents()
            self.ui.error_list_table.setColumnWidth(
                3, self.ui.error_list_table.columnWidth(3)-15)
            self.ui.error_list_table.setColumnWidth(
                2, self.ui.error_list_table.columnWidth(2)-15)
            self.ui.error_list_table.resizeRowsToContents()
            self.ui.error_list_table.setSortingEnabled(True)
            self.ui.error_list_table.setUpdatesEnabled(True)
            self.ui.error_list_table.setAlternatingRowColors(True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load error codes: {e}")

    def add_item_to_error_list(self, data, model):
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
                params = None
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
                    conditions.append(
                        "(ecl.error_code LIKE :search OR ecl.error_description LIKE :search OR ecl.reason LIKE :search OR ecl.recommended_action LIKE :search)")
                    params['search'] = f"%{search_char}%"
                filter_scripts = "WHERE " + " AND ".join(conditions)
            result = self.database.query(sql=f''' SELECT ecl.error_code, ecl.error_description, ecl.reason, ecl.recommended_action, ecl.process, da.downtime_area_name
                                                FROM `error_codes_list` ecl
                                                JOIN `downtime_areas` da ON ecl.downtime_area_id = da.downtime_area_id
                                                JOIN `departments` d ON da.department_id = d.department_id
                                                {filter_scripts}
                                                ORDER BY ecl.error_code ASC, ecl.process COLLATE utf8mb4_unicode_ci ASC;''', params=params)
            self.ui.error_list_table.setSortingEnabled(False)
            self.ui.error_list_table.setUpdatesEnabled(False)
            self.add_item_to_error_list(result, self.error_model)
            self.ui.error_list_table.setSortingEnabled(True)
            self.ui.error_list_table.setUpdatesEnabled(True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to filter error codes: {e}")

    @QtCore.pyqtSlot()
    def load_area(self):
        try:
            selected_group = self.ui.Group_cbb.currentText()
            if selected_group == "All":
                result = self.database.query(sql=''' SELECT DISTINCT da.downtime_area_name
                                                        FROM `downtime_areas` da
                                                        ORDER BY da.downtime_area_name COLLATE utf8mb4_unicode_ci ASC;''')
            else:
                result = self.database.query(sql=''' SELECT DISTINCT da.downtime_area_name
                                                        FROM `downtime_areas` da
                                                        JOIN `departments` d ON da.department_id = d.department_id
                                                        WHERE d.department_name = :group
                                                        ORDER BY da.downtime_area_name COLLATE utf8mb4_unicode_ci ASC;''', params={'group': selected_group})
            self.ui.Area_cbb.clear()
            if result:
                self.ui.Area_cbb.addItems(["All"] + [row[0] for row in result])
                self.ui.Area_cbb.setCurrentText("All")
                self.load_process()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load areas: {e}")

    @QtCore.pyqtSlot()
    def load_process(self):
        try:
            area = self.ui.Area_cbb.currentText()
            if area == "All":
                result = self.database.query(sql=''' SELECT DISTINCT process 
                                                        FROM `error_codes_list` ecl
                                                        JOIN `downtime_areas` da ON ecl.downtime_area_id = da.downtime_area_id
                                                        ORDER BY process COLLATE utf8mb4_unicode_ci ASC;''')
            else:
                result = self.database.query(sql=''' SELECT DISTINCT process 
                                                        FROM `error_codes_list` ecl
                                                        JOIN `downtime_areas` da ON ecl.downtime_area_id = da.downtime_area_id
                                                        WHERE da.downtime_area_name = :area
                                                        ORDER BY process COLLATE utf8mb4_unicode_ci ASC;''', params={'area': area})
            self.ui.Process_cbb.clear()
            if result:
                self.ui.Process_cbb.addItems(
                    ["All"] + [row[0] for row in result])
                self.ui.Process_cbb.setCurrentText("All")
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to load processes: {e}")

    @QtCore.pyqtSlot()
    def delete_row(self):
        if self.error_model.rowCount() == 0:
            QtWidgets.QMessageBox.warning(
                self, "Warning", "No rows available to delete.")
            return
        index = self.ui.error_list_table.currentIndex()
        if not index.isValid():
            QtWidgets.QMessageBox.warning(
                self, "Warning", "Please select a row to delete.")
            return
        row = index.row()
        if row in self.new_errors_code:
            self.new_errors_code.remove(row)
        else:
            self.remove_error_code.append(
                (row, self.error_model.item(row, 0).text()))
        self.error_model.removeRow(row)
        self.new_errors_code = [r-1 if r >
                                row else r for r in self.new_errors_code]
        self.remove_error_code = [(r-1 if r > row else r, code)
                                  for r, code in self.remove_error_code]

    @QtCore.pyqtSlot()
    def insert_row(self):
        try:
            self.error_model.insertRow(self.error_model.rowCount())
            for col in range(self.error_model.columnCount()):
                item = QtGui.QStandardItem("")
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                item.setEditable(True)
                self.error_model.setItem(
                    self.error_model.rowCount() - 1, col, item)
            self.ui.error_list_table.setEditTriggers(
                QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.EditKeyPressed)
            self.new_errors_code.append(self.error_model.rowCount() - 1)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to insert new error code: {e}")

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
            reply = QtWidgets.QMessageBox.question(
                self, "Confirm", f'''Are you sure you want to apply changes? \n{question}''', QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.No:
                return
            for row, error_code in self.remove_error_code:
                error_code_list_remove.append({'code': error_code})
            if error_code_list_remove:
                self.database.executemany(
                    sql=''' DELETE FROM `error_codes_list` WHERE error_code = :code ''', params_list=error_code_list_remove)
            for row in self.new_errors_code:
                error_code = self.error_model.item(row, 0).text()
                description = self.error_model.item(row, 1).text()
                reason = self.error_model.item(row, 2).text()
                recommended_action = self.error_model.item(row, 3).text()
                process = self.error_model.item(row, 4).text()
                area_name = self.error_model.item(row, 5).text()
                if not error_code or not area_name:
                    QtWidgets.QMessageBox.warning(
                        self, "Warning", f"Error Code and Downtime Area cannot be empty. Please check row {row+1}.")
                    return
                self.database.query(sql=''' INSERT INTO `error_codes_list` (error_code, error_description, reason, recommended_action, process, downtime_area_id)
                                            VALUES (:code, :description, :reason, :action, :process,
                                            (SELECT downtime_area_id FROM `downtime_areas` WHERE downtime_area_name = :area LIMIT 1))
                                            ON DUPLICATE KEY UPDATE
                                            error_description = VALUES(error_description),
                                            reason = VALUES(reason),
                                            recommended_action = VALUES(recommended_action),
                                            process = VALUES(process),
                                            downtime_area_id = VALUES(downtime_area_id);''',
                                    params={
                                        'code': error_code,
                                        'description': description,
                                        'reason': reason,
                                        'action': recommended_action,
                                        'process': process,
                                        'area': area_name
                                    })
            QtWidgets.QMessageBox.information(
                self, "Success", "Changes updated successfully.")
            self.accept()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to update changes: {e}")
            return

class RotatedAxisItem(pg.AxisItem):
    def __init__(self, angle=-45, dx=-10, dy=5, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.angle = angle
        self.dx = dx
        self.dy = dy

    def drawPicture(self, p, axisSpec, tickSpecs, textSpecs):
        # if self.angle == 0:
        #     super().drawPicture(p, axisSpec, tickSpecs, textSpecs)
        #     return
        p.setRenderHint(p.Antialiasing, False)
        p.setRenderHint(p.TextAntialiasing, True)

        # pen, p1, p2 = axisSpec
        # p.setPen(pen)
        # p.drawLine(p1, p2)

        # for pen, p1, p2 in tickSpecs:
        #     p.setPen(pen)
        #     p.drawLine(p1, p2)

        if self.style['showValues']:
            p.setFont(self.style['tickFont'] or self.font())
            p.setPen(self.textPen())
            for rect, flags, text in textSpecs:
                if self.angle == 0:
                    p.drawText(rect, flags, text)
                else:
                    offset_rect = rect.translated(self.dx, self.dy)
                    p.save()
                    p.translate(offset_rect.center())
                    p.rotate(self.angle)
                    p.translate(-offset_rect.center())
                    p.drawText(offset_rect, flags, text)
                    p.restore()

class New_Report_Input(QtWidgets.QDialog):
    def __init__(self, parent=None, database=None, callback=None):
        super().__init__(parent)
        self.parent = parent
        self.database = database
        self.callback = callback
        self.ui = Ui_ReportInput()
        self.ui.setupUi(self)
        self.ui.dateEdit.setDate(QtCore.QDate.currentDate())
        self.setWindowTitle("New Report Input")
        self.setup_drop_area()
        self.ui.group_cbb.addItems([row[0] for row in self.parent.group])
        self.ui.group_cbb.setCurrentText(self.parent.login_info['department'])
        try:
            lines = self.database.query(sql='''   SELECT line_name FROM `production_lines` as pl
                                                    JOIN `departments` as d
                                                    ON pl.department_id = d.department_id
                                                    WHERE d.department_name = :department 
                                                    ORDER BY pl.line_name ASC;''', params={'department': self.parent.login_info['department']})
            if lines:
                self.ui.line_cbb.addItems([""] + [row[0] for row in lines])
            report_types = self.database.query(
                sql=''' SELECT report_type_name FROM `report_types` ''')
            if report_types:
                self.ui.report_type_cbb.addItems(
                    [row[0] for row in report_types])
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to fetch lines: {e}")
        self.setup_signals()

    def setup_signals(self):
        self.ui.Cancel_btn.clicked.connect(self.reject)
        self.ui.Confirm_btn.clicked.connect(self.confirm_report)
        self.ui.group_cbb.currentIndexChanged.connect(self.load_lines)

    def setup_drop_area(self):
        self.ui.drop_file_area_toolbtn.setAcceptDrops(True)
        self.ui.drop_file_area_toolbtn.dragEnterEvent = self.dragEnterEvent
        self.ui.drop_file_area_toolbtn.dragMoveEvent = self.dragMoveEvent
        self.ui.drop_file_area_toolbtn.dropEvent = self.dropEvent

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
            self.file_path = file_path
            file_name = file_path.split("/")[-1]
            file_extension = file_name.split(
                ".")[-1].upper() if "." in file_name else ""
            self.ui.drop_file_area_toolbtn.setText(file_name)
            self.ui.drop_file_area_toolbtn.setToolTip(file_path)
            self.ui.drop_file_area_toolbtn.setStyleSheet("""
                #drop_file_area_toolbtn {
                    background-color: rgba(0, 255, 0, 0.07);
                    border: none;
                    border-top: 1px solid rgba(0, 255, 0, 1);
                    border-bottom: 1px solid rgba(0, 255, 0, 1);
                }
            """)
            icon = self.callback(file_extension=file_extension)
            self.ui.drop_file_area_toolbtn.setIcon(icon)
            self.ui.drop_file_area_toolbtn.setIconSize(QtCore.QSize(42, 42))
            event.acceptProposedAction()
        else:
            super().dropEvent(event)

    def load_lines(self):
        try:
            lines = self.database.query(sql=''' SELECT line_name FROM `production_lines` as pl
                                                    JOIN `departments` as d
                                                    ON pl.department_id = d.department_id
                                                    WHERE d.department_name = :department 
                                                    ORDER BY pl.line_name ASC;''', params={'department': self.ui.group_cbb.currentText()})
            if lines:
                self.ui.line_cbb.clear()
                self.ui.line_cbb.addItems([""] + [row[0] for row in lines])
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to fetch lines: {e}")

    def check_before_update(self):
        if not self.ui.report_title_lnedit.text():
            QtWidgets.QMessageBox.warning(
                self, "Warning", "Report Title cannot be empty.")
            return False
        if not self.ui.group_cbb.currentText():
            QtWidgets.QMessageBox.warning(
                self, "Warning", "Group cannot be empty.")
            return False
        if not self.ui.report_type_cbb.currentText():
            QtWidgets.QMessageBox.warning(
                self, "Warning", "Report Type cannot be empty.")
            return False
        if not self.file_path:
            QtWidgets.QMessageBox.warning(
                self, "Warning", "Please upload a file.")
            return False
        if not self.ui.report_by_lnedit.text():
            QtWidgets.QMessageBox.warning(
                self, "Warning", "Report By cannot be empty.")
            return False
        return True

    def confirm_report(self):
        try:
            if not self.check_before_update():
                return
            values_scripts = f'''   "{self.ui.report_title_lnedit.text()}",
                                    "{self.ui.dateEdit.date().toString("yyyy-MM-dd")}",
                                    "{self.ui.issue_descrip_text.toPlainText() if self.ui.issue_descrip_text.toPlainText() else None}",
                                    "{self.ui.correct_act_text.toPlainText() if self.ui.correct_act_text.toPlainText() else None}",
                                    "{self.ui.notes_text.toPlainText() if self.ui.notes_text.toPlainText() else None}",
                                    "{self.file_path}",
                                    "Finish",'''
            values_scripts += f'''(SELECT staff_id FROM staff WHERE staff_name = "{self.ui.report_by_lnedit.text()}"),'''
            values_scripts += f'''(SELECT department_id FROM departments WHERE department_name = "{self.ui.group_cbb.currentText()}"),'''
            values_scripts += f'''(SELECT line_id FROM production_lines WHERE line_name = "{self.ui.line_cbb.currentText()}"),''' if self.ui.line_cbb.currentText(
            ) else "NULL,"
            values_scripts += f'''(SELECT machine_id FROM machines WHERE machine_code = "{self.ui.machine_code.text()}"),''' if self.ui.machine_code.text(
            ) else "NULL,"
            values_scripts += f'''(SELECT report_type_id FROM report_types WHERE report_type_name = "{self.ui.report_type_cbb.currentText()}"))'''
            self.database.query(sql=''' INSERT INTO `problem_reports` (report_title, report_date,issue_description, corrective_action,
                                                                         notes, report_file_path, status,reported_by,
                                                                        department_id, line_id, machine_id, report_type_id)
                                        VALUES (''' + values_scripts)
            QtWidgets.QMessageBox.information(
                self, "Success", "Report submitted successfully.")
            self.accept()
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Warning", str(e))
            return

class DonutChart(QtWidgets.QWidget):
    def __init__(self, value: float, max_value: float = 100, target_value: float = 0, background_color: str = "#E8E8E8", foreground_color: str = "#2ECC71", target_color: str = "#affaff", parameter_name: str = "OEE", scale: float = 0.7, has_Lengend: bool = True, parent=None):
        super().__init__(parent)
        self.value = value
        self.max_value = max_value
        self.target_value = target_value
        self.background_color = background_color
        self.foreground_color = foreground_color
        self.target_color = target_color
        self.parameter_name = parameter_name
        # self.setMinimumSize(200, 400)
        if self.value == self.target_value:
            self.foreground_color = "#facc15"
        elif self.value < self.target_value:
            self.foreground_color = "#f61863"
        else:
            self.foreground_color = "#2ECC71"
        self.scale = scale
        self.has_Lengend = has_Lengend

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)

        width = self.width()
        height = self.height()
        size = min(width, height)*self.scale
        x = (width - size) / 2
        y = (height - size) / 4

        stroke_width = size * 0.12
        margin = stroke_width / 2

        rect = QtCore.QRectF(x + margin, y + margin,
                             size - stroke_width, size - stroke_width)

        percentage = self.value / self.max_value
        span_angle = int(percentage * 360 * 16)
        full_angle = 360 * 16
        start_angle = 90 * 16

        target_percentage = self.target_value / self.max_value
        target_span_angle = int(target_percentage * 360 * 16)

        def draw_circle(pen_color, pen_width, cap_style, span_angle):
            pen = QtGui.QPen(QtGui.QColor(pen_color))
            pen.setWidth(int(pen_width))
            pen.setCapStyle(cap_style)
            painter.setPen(pen)
            painter.drawArc(rect, start_angle, -span_angle)

        draw_circle(self.background_color, stroke_width,
                    QtCore.Qt.PenCapStyle.FlatCap, full_angle)
        draw_circle(self.foreground_color, stroke_width,
                    QtCore.Qt.PenCapStyle.RoundCap, span_angle)
        draw_circle(self.target_color, stroke_width*0.5,
                    QtCore.Qt.PenCapStyle.RoundCap, target_span_angle)

        painter.setPen(QtGui.QColor("#333333"))
        font = QtGui.QFont("Arial", int(size * 0.18), QtGui.QFont.Weight.Bold)
        painter.setFont(font)
        painter.drawText(
            QtCore.QRectF(x, y, size, size),
            QtCore.Qt.AlignmentFlag.AlignCenter,
            f"{int(self.value)}%"
        )
        if self.has_Lengend:
        # Legend
            if self.parameter_name == "%OEE":
                dot_r = int(size * 0.045)
                legend_x = int(x - x*0.05)
                legend_y = int(y + size + dot_r * 3)

                painter.setBrush(QtGui.QColor(self.foreground_color))
                painter.setPen(QtCore.Qt.PenStyle.NoPen)
                painter.drawEllipse(legend_x, legend_y - dot_r +
                                    3, dot_r * 2, dot_r * 2)
                painter.setPen(QtGui.QColor("#555555"))
                font_legend = QtGui.QFont("Arial", int(size * 0.07))
                painter.setFont(font_legend)
                painter.drawText(
                    QtCore.QRectF(legend_x + dot_r * 2 + int(size * 0.03),
                                legend_y - dot_r,
                                size * 0.4, dot_r * 3),
                    QtCore.Qt.AlignmentFlag.AlignVCenter | QtCore.Qt.AlignmentFlag.AlignLeft,
                    f"{self.parameter_name}"
                )
                legend_x2 = legend_x + dot_r * 2 + int(size * 0.4)
                painter.setBrush(QtGui.QColor(self.target_color))
                painter.setPen(QtCore.Qt.PenStyle.NoPen)
                painter.drawEllipse(legend_x2, legend_y - dot_r +
                                    3, dot_r * 2, dot_r * 2)
                painter.setPen(QtGui.QColor("#555555"))
                painter.setFont(font_legend)
                painter.drawText(
                    QtCore.QRectF(legend_x2 + dot_r * 2 + int(size * 0.03),
                                legend_y - dot_r,
                                size * 0.8, dot_r * 3),
                    QtCore.Qt.AlignmentFlag.AlignVCenter | QtCore.Qt.AlignmentFlag.AlignLeft,
                    f"Target {self.target_value}%"
                )
            else:
                dot_r = int(size * 0.045)
                legend_x = int(x + size * 0.4)
                legend_y = int(y + size + dot_r * 3)

                painter.setBrush(QtGui.QColor(self.foreground_color))
                painter.setPen(QtCore.Qt.PenStyle.NoPen)
                painter.drawEllipse(legend_x, legend_y - dot_r +
                                    3, dot_r * 2, dot_r * 2)
                painter.setPen(QtGui.QColor("#555555"))
                font_legend = QtGui.QFont("Arial", int(size * 0.07))
                painter.setFont(font_legend)
                painter.drawText(
                    QtCore.QRectF(legend_x + dot_r * 2 + int(size * 0.03),
                                legend_y - dot_r,
                                size * 0.4, dot_r * 3),
                    QtCore.Qt.AlignmentFlag.AlignVCenter | QtCore.Qt.AlignmentFlag.AlignLeft,
                    f"{self.parameter_name}"
                )
        painter.end()

class Bullet_Status_Bar(QtWidgets.QWidget):
    def __init__(self, parent=None, value: float = 0, max_value: float = 0, target_value: float = 0, previous_value: float = 0, background_color: str = "#E8E8E8",
                  foreground_color: str = "#2ECC71", target_color: str = "#0A21EE", html_doc = False, format_time = None, label = "MTTR", height_ratio: float = 0.2):
        super().__init__(parent)
        self.value = value
        self.target_value = target_value
        self.previous_value = previous_value
        self.max_value = (self.target_value*1.5 if self.target_value > self.value else self.value*1.5) if max_value == 0 else max_value
        self.background_color = background_color
        self.foreground_color = foreground_color
        self.target_color = target_color
        self.html_doc = html_doc
        self.format_time = format_time
        self.label = label
        self.height_ratio = height_ratio
        if self.target_value > 0:
            if self.label == "MTTR":
                if (self.value < self.target_value) :
                    self.foreground_color = "#2ECC71"
                else:
                    self.foreground_color = "#f61863"
            else:
                if (self.value > self.target_value) :
                    self.foreground_color = "#2ECC71"
                else:
                    self.foreground_color = "#f61863"
        else:
            self.foreground_color = "#2ECC71"
        self.text_color = "#C9D1D9"
        self.muted_text = "#8B93A7"
    
    def make_comparison_html(self, current: float, previous: float, target: float, label: str, time_format: None) -> str:
        diff = current - previous
        diff = round(diff, 1)
        diff_target = current - target
        diff_target = round(diff_target, 1)
        def colorize(value):
            if value > 0:
                return "&#9650;" , "#27ae60" , "#e74c3c" , "+"
            elif value < 0:
                return "&#9660;" , "#e74c3c" , "#27ae60" , "-"
            else:
                return "&#9654;" , "#888888" , "#888888" , ""
        arrow, color, color_dt, sign = colorize(diff)
        arrow_target, color_target, color_dt_target, sign_target = colorize(diff_target)
    
        percent = (1 - current/target) if target != 0 else 0  
        percent_prev = (1-current/previous) if previous != 0 else 0
        time_format = time_format(previous,"m")
        html_content = f""
        if target > 0:
            html_content = f"""
            <div style='font-family: Arial; text-align: center;'>
                <span style='font-size: 10px; font-weight: bold; color: {color_dt_target if label == "MTTR" else color_target};'>
                    {f"{arrow_target} {sign_target}{abs(percent*100):.1f}% vs target"}
                </span>
            </div>"""
        if previous != 0:
            html_content += f"""
             <div style='font-family: Arial; text-align: center;'>
            <span style='font-size: 10px; font-weight: bold; color: #222;'>Previous:</span>
            <span style='font-size: 10px; font-weight: bold; color: #222;'>
                {f"{time_format['h']} hrs {time_format['m']} mins" if previous >= 60 else f"{time_format['m']} mins {time_format['s']} secs"}
            </span>
            <br/>
            <span style='font-size: 10px; color: {color_dt if label == "MTTR" else color}; font-weight: bold;'>
                {arrow} {sign}{abs(percent_prev*100):.1f}% vs prev
            </span>
        </div>
        """
        return html_content
       
    def paintEvent(self, event): 
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)

        width_bar = self.width()*0.9
        height_bar = self.height() *0.2
        left_margin = (self.width() - width_bar) / 2
        top_margin = (self.height() - height_bar) *self.height_ratio
        radius = height_bar / 2
        font_min = QtGui.QFont("Arial", 9, QtGui.QFont.Weight.Bold)
        painter.setFont(font_min)
        painter.setPen(QtGui.QColor(self.muted_text))
        painter.drawText(QtCore.QRectF(left_margin, top_margin-radius*2.1, width_bar, height_bar),
                            QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter,
                            f"0")
        
        bar_rect_background = QtCore.QRectF(left_margin, top_margin, width_bar, height_bar)
        painter.setPen(QtGui.QColor(self.muted_text))
        painter.setBrush(QtGui.QColor(self.background_color))
        painter.drawRoundedRect(bar_rect_background, radius, radius)
       
        ratio = self.value / self.max_value
        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        if self.value <=0:
            return
        elif ratio < ( radius / width_bar):
            bar_value_rect = QtCore.QRectF(left_margin, top_margin, radius , height_bar)
            painter.setBrush(QtGui.QColor(self.foreground_color))
            painter.drawRoundedRect(bar_value_rect, radius, radius)

        else:
            bar_value_rect = QtCore.QRectF(left_margin + ratio*width_bar - radius , top_margin, radius , height_bar)   
            painter.setBrush(QtGui.QColor(self.foreground_color))
            painter.drawRect(bar_value_rect)

            bar_value_rect_full = QtCore.QRectF(left_margin, top_margin, width_bar * self.value / self.max_value, height_bar)
            painter.setBrush(QtGui.QColor(self.foreground_color))
            painter.drawRoundedRect(bar_value_rect_full, radius, radius)

        if self.previous_value > 0:        
            previous = left_margin + width_bar * self.previous_value / self.max_value
            painter.setBrush(QtGui.QColor(self.muted_text))
            painter.setPen(QtCore.Qt.PenStyle.NoPen)
            painter.drawRect(QtCore.QRectF(previous - 1, top_margin, 2, height_bar))

            previous_text_rect = QtCore.QRectF(left_margin + width_bar * (self.previous_value / self.max_value)-22, top_margin+radius*2.1, 40, height_bar)
            painter.setFont(QtGui.QFont("Arial", 9, QtGui.QFont.Weight.Bold))
            painter.setPen(QtGui.QColor(self.muted_text))
            painter.drawText(previous_text_rect, QtCore.Qt.AlignmentFlag.AlignCenter, f" Prev")
        if self.target_value > 0:
            target_x = left_margin + width_bar * self.target_value / self.max_value
            painter.setBrush(QtGui.QColor(self.target_color))
            painter.setPen(QtCore.Qt.PenStyle.NoPen)
            painter.drawRect(QtCore.QRectF(target_x - 1, top_margin, 2, height_bar))
            target_text_rect = QtCore.QRectF(left_margin + width_bar * (self.target_value / self.max_value) - 20, top_margin-radius*2.1, 80, height_bar)
            painter.setFont(QtGui.QFont("Arial", 9, QtGui.QFont.Weight.Bold))
            painter.setPen(QtGui.QColor(self.target_color))
            painter.drawText(target_text_rect, QtCore.Qt.AlignmentFlag.AlignCenter, f"KPI = {int(self.target_value)}m")
    
        if self.html_doc:
            html_content = self.make_comparison_html(self.value,self.previous_value ,self.target_value, self.label, self.format_time)
            doc = QtGui.QTextDocument()
            doc.setHtml(html_content)
            doc.setTextWidth(self.width())
            painter.save()
            painter.translate(0, top_margin + height_bar + radius*1.5)
            doc.drawContents(painter)
            painter.restore()

        painter.end()

class ToggleSwitch(QtWidgets.QCheckBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(40, 20)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self._circle_pos = 3.0
        self.animation = QtCore.QPropertyAnimation(self, b"circle_pos", self)
        self.animation.setEasingCurve(QtCore.QEasingCurve.InOutCubic)
        self.animation.setDuration(200)
        self.stateChanged.connect(self.start_animation)

    @QtCore.pyqtProperty(float)
    def circle_pos(self):
        return self._circle_pos
    
    def hitButton(self, pos):
        return self.contentsRect().contains(pos)

    @circle_pos.setter
    def circle_pos(self, pos):
        self._circle_pos = pos
        self.update()  

    def start_animation(self, value):
        self.animation.stop()
        start = self._circle_pos
        end = self.width() - 17 if value else 3
        self.animation.setStartValue(start)
        self.animation.setEndValue(end)
        self.animation.start()

    def paintEvent(self, event):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        p.setPen(QtCore.Qt.NoPen)
        bg_color = QtGui.QColor("#2ecc71") if self.isChecked() else QtGui.QColor("#DADADA")
        p.setBrush(QtGui.QBrush(bg_color))
        p.drawRoundedRect(0, 0, self.width(), self.height(),
                          self.height() / 2, self.height() / 2)
        p.setBrush(QtGui.QBrush(QtGui.QColor("white")))
        p.drawEllipse(QtCore.QRectF(self._circle_pos, 3, 14, 14))

        p.end()

class OEE_Edit_Data(QtWidgets.QDialog):
    def __init__(self, parent=None, database = None, data=None, cycle_time = 0, machine_code = None):
        super().__init__(parent)
        self.ui = UI_OEE_Edit_Data()
        self.ui.setupUi(self)
        self.database = database
        self.data = data
        self.machine_code = machine_code
        for i in range(5,len(self.data)):
            if isinstance(self.data.iloc[i], Decimal):
                self.data.iloc[i] = float(self.data.iloc[i])
        self.cycle_time = cycle_time
        self.setWindowIcon(QtGui.QIcon(resource_path("Icons/OEE.ico")))
        vertical_headers = ["Date", "Working Shift", "Break Time (min)", "Setup Time (min)", "Plan time (min)", "Total Loss Time (min)", "Available Time (min)", "FGs (pcs)", "Defect (pcs)", "Availability (%)", "Performance (%)", "Quality (%)", "OEE (%)"]
        self.data_model = QtGui.QStandardItemModel(len(vertical_headers), 2)
        data_for_table = list(self.data[4:] if len(self.data) > 4 else [None] * (len(vertical_headers) - 4))
        for row in range(len(vertical_headers)):
            item = QtGui.QStandardItem(vertical_headers[row])
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.data_model.setItem(row, 0, item)
            item = QtGui.QStandardItem(str(data_for_table[row]) if data_for_table[row] is not None else "")
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.data_model.setItem(row, 1, item)
        self.ui.result_table.setModel(self.data_model)
        self.ui.result_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        class downtime_record_widgetItem(QtWidgets.QWidget):
            delete_requested = QtCore.pyqtSignal(object)
            edit_completed = QtCore.pyqtSignal(object)
            def __init__(self, record):
                super().__init__()
                self.id = record[0]
                self.date = record[1]
                self.start_time = record[2]
                self.repair_time = record[3]
                self.end_time = record[4]
                self.staff_name = record[5]
                self.error_code = record[6]
                self.machine_code = record[7]
                self.line_name = record[8]
                self.total_loss_time = self.calculate_loss_time()
                layout = QtWidgets.QVBoxLayout(self)
                layout.setContentsMargins(8, 8, 8, 8)
                layout.setSpacing(4)
                date_label = QtWidgets.QLabel(f"Date: {self.date} | Machine: {self.machine_code} ")
                date_label.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                date_label.setStyleSheet("font-weight: bold;")
                frame_1 = QtWidgets.QFrame()
                frame_1.setFrameShape(QtWidgets.QFrame.NoFrame)
                layout_1 = QtWidgets.QHBoxLayout(frame_1)
                layout_1.setContentsMargins(0, 0, 0, 0)
                layout_1.addWidget(QtWidgets.QLabel("Staff: "))
                staff_name_lnedit = QtWidgets.QLineEdit(f"{self.staff_name}")
                staff_name_lnedit.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                layout_1.addWidget(staff_name_lnedit)
                frame_2 = QtWidgets.QFrame()
                frame_2.setFrameShape(QtWidgets.QFrame.NoFrame)
                layout_2 = QtWidgets.QHBoxLayout(frame_2)
                layout_2.setContentsMargins(0, 0, 0, 0)
                layout_2.addWidget(QtWidgets.QLabel("Error Code: "))
                error_code_lnedit = QtWidgets.QLineEdit(f"{self.error_code}")
                error_code_lnedit.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                layout_2.addWidget(error_code_lnedit)
                frame_3 = QtWidgets.QFrame()
                frame_3.setFrameShape(QtWidgets.QFrame.NoFrame)
                layout_3 = QtWidgets.QHBoxLayout(frame_3)
                layout_3.setContentsMargins(0, 0, 0, 0)
                layout_3.addWidget(QtWidgets.QLabel("Start Time:"))
                self.start_time_lnedit = QtWidgets.QLineEdit(str(self.start_time) if self.start_time else "")
                self.start_time_lnedit.setContextMenuPolicy(QtCore.Qt.NoContextMenu)
                layout_3.addWidget(self.start_time_lnedit)
                frame_4 = QtWidgets.QFrame()
                frame_4.setFrameShape(QtWidgets.QFrame.NoFrame)
                layout_4 = QtWidgets.QHBoxLayout(frame_4)
                layout_4.setContentsMargins(0, 0, 0, 0)
                layout_4.addWidget(QtWidgets.QLabel("Repair Time (min):"))
                self.repair_time_lnedit = QtWidgets.QLineEdit(str(self.repair_time) if self.repair_time else "")
                self.repair_time_lnedit.setContextMenuPolicy(QtCore.Qt.NoContextMenu)
                layout_4.addWidget(self.repair_time_lnedit)
                frame_5 = QtWidgets.QFrame()
                frame_5.setFrameShape(QtWidgets.QFrame.NoFrame)
                layout_5 = QtWidgets.QHBoxLayout(frame_5)
                layout_5.setContentsMargins(0, 0, 0, 0)
                layout_5.addWidget(QtWidgets.QLabel("End Time:"))
                self.end_time_lnedit = QtWidgets.QLineEdit(str(self.end_time) if self.end_time else "")
                self.end_time_lnedit.setContextMenuPolicy(QtCore.Qt.NoContextMenu)
                layout_5.addWidget(self.end_time_lnedit)
                self.total_loss_time_label = QtWidgets.QLabel(f"Total Loss Time: {self.total_loss_time} min")
                layout.addWidget(date_label)
                layout.addWidget(frame_1)
                layout.addWidget(frame_2)
                layout.addWidget(frame_3)
                layout.addWidget(frame_4)
                layout.addWidget(frame_5)
                layout.addWidget(self.total_loss_time_label)
                self.setObjectName("downtime_card")
                self.setAttribute(QtCore.Qt.WA_StyledBackground, True)
                self.setStyleSheet("""
                QWidget#downtime_card {
                    background: #ffffff;
                    border: 1px solid #e0e0e0;
                    border-radius: 8px;
                    border-left: 4px solid #007acc;
                }
                QWidget#downtime_card:hover {
                    background: #f0f7ff;
                    border: 1px solid #007acc;
                    border-left: 4px solid #007acc;
                }
                QLabel {
                    background-color: transparent;
                    border: none;
                    color: #012d4b;
                }
                QFrame {
                    background-color: transparent;
                    border: none;
                }
                QLineEdit {
                    background-color: rgba(255, 255, 255, 0.15);
                    border: 1px solid rgba(255, 255, 255, 0.3);
                    border-radius: 4px;
                    padding: 2px 6px;
                    color: #012d4b;
                }
                QLineEdit:focus {
                    border: 1px solid rgba(255, 255, 255, 0.7);
                    background-color: rgba(255, 255, 255, 0.22);
                }               
            """)
                self.start_time_lnedit.returnPressed.connect(self.update_loss_time)
                self.end_time_lnedit.returnPressed.connect(self.update_loss_time)
                staff_name_lnedit.editingFinished.connect(lambda: setattr(self, 'staff_name', staff_name_lnedit.text().strip()))
                error_code_lnedit.editingFinished.connect(lambda: setattr(self, 'error_code', error_code_lnedit.text().strip()))
                self.repair_time_lnedit.editingFinished.connect(lambda: setattr(self, 'repair_time', self.repair_time_lnedit.text().strip()))
                self.context_menu = QtWidgets.QMenu(self)
                self.context_menu.addAction("Delete Record", lambda: self.delete_record())
                self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
                self.customContextMenuRequested.connect(lambda pos: self.show_context_menu(pos))
            
            @QtCore.pyqtSlot()
            def show_context_menu(self, pos):
                self.context_menu.exec_(self.mapToGlobal(pos))

            @QtCore.pyqtSlot()    
            def delete_record(self):
                reply = QtWidgets.QMessageBox.question(
                    self, "Confirm", "Are you sure you want to delete this downtime record?\nThis action cannot be undone.",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
                if reply == QtWidgets.QMessageBox.Yes:
                    self.delete_requested.emit(self) 
                    

            @QtCore.pyqtSlot()
            def update_loss_time(self):
                self.start_time = self.start_time_lnedit.text().strip()
                self.end_time = self.end_time_lnedit.text().strip()
                self.total_loss_time = self.calculate_loss_time()
                self.total_loss_time_label.setText(f"Total Loss Time: {self.total_loss_time} min")
                self.edit_completed.emit(self)

            @QtCore.pyqtSlot()
            def calculate_loss_time(self):
                try:
                    if self.start_time and self.end_time:
                        def to_time(t):
                            if isinstance(t, dt.timedelta):
                                total = int(t.total_seconds())
                                return dt.time(total // 3600, (total % 3600) // 60, total % 60)
                            return dt.datetime.strptime(str(t), "%H:%M:%S").time()

                        base = self.date if isinstance(self.date, dt.date) else dt.date.fromisoformat(str(self.date))
                        start_dt = dt.datetime.combine(base, to_time(self.start_time))
                        end_dt = dt.datetime.combine(base, to_time(self.end_time))

                        if end_dt < start_dt:
                            end_dt += dt.timedelta(days=1)
                        return int((end_dt - start_dt).total_seconds() / 60)
                except Exception:
                    pass
                return 0
        self.downtime_record_widgetItem = downtime_record_widgetItem 
        try:
            area =  self.data["area_name"]
            line = self.data["line_name"]
            model = self.data["model_name"]
            process = self.data["process"]
            production_date = self.data["production_date"]
            self.working_shift = self.database.query(sql=''' SELECT operation_id,operation_hours,change_model,change_from, break_time, setup_time
                                                                FROM `line_operation_times` as lot
                                                                JOIN `production_lines` as pl ON lot.line_id = pl.line_id
                                                                WHERE pl.line_name = :line AND lot.operation_date = :date AND model_running = :model'''
                                                            , params={'line': line, 'date': production_date, 'model': model})
            
            self.downtime_records  = self.database.query(sql='''SELECT dr.Downtime_ID, dr.Date, dr.Start_Time, dr.Start_Repair_Time, dr.End_Time,dr.Staff_Name, dr.Error_Code, dr.Machine_Code, dr.Line_Name
                                                            FROM `downtime_report` as dr
                                                            JOIN machines as m ON dr.Machine_Code = m.machine_code
                                                            JOIN machine_oee_register as mor ON m.machine_id = mor.machine_id
                                                            JOIN production_lines as pl ON mor.line_id = pl.line_id AND pl.line_name = dr.Line_Name
                                                            JOIN product_models_oee as pmo ON mor.model_id = pmo.model_id AND pmo.model_name = dr.Current_Model
                                                            WHERE dr.Line_Name = :line AND dr.Current_Model = :model AND dr.Date = :date AND mor.process = :process; ''', 
                                                            params={'line': line, 'model': model, 'date': production_date, 'process': process})
            self.production_outputs = self.database.query(sql = '''SELECT po.output_id, po.OK_qty, po.NG_qty 
                                                            FROM production_output as po
                                                            JOIN production_lines as pl ON po.line_id = pl.line_id
                                                            WHERE pl.line_name = :line AND po.model_name = :model 
                                                            AND po.production_date = :date;''',
                                                            params={'line': line, 'model': model, 'date': production_date})
            self.ui.working_shift_lnedit.setText(str(self.working_shift[0][1]) if self.working_shift else "")
            self.ui.working_shift_lnedit.setSuffix(f" Bf: {self.working_shift[0][1]}" if self.working_shift else "")
            self.ui.breaktime_lnedit.setText(str(self.working_shift[0][4]) if self.working_shift else "")
            self.ui.breaktime_lnedit.setSuffix(f" Bf: {self.working_shift[0][4]}" if self.working_shift else "")
            self.ui.setuptime_lnedit.setText(str(self.working_shift[0][5]) if self.working_shift else "")
            self.ui.setuptime_lnedit.setSuffix(f" Bf: {self.working_shift[0][5]}" if self.working_shift else "")
            self.downtime_records = [self.downtime_record_widgetItem(record ) for record in self.downtime_records]
            self.ui.downtime_records_widget.clear()
            for record in self.downtime_records:
                record.delete_requested.connect(lambda widget=record: self.remove_downtime_record(widget))
                record.edit_completed.connect(lambda widget=record: self.calculate_oee(edit_object="downtime_record"))
                list_item = QtWidgets.QListWidgetItem(self.ui.downtime_records_widget)
                list_item.setSizeHint(QtCore.QSize(0, 160))
                self.ui.downtime_records_widget.addItem(list_item)
                self.ui.downtime_records_widget.setItemWidget(list_item, record)
            self.ui.FGs_lnedit.setText(str(self.production_outputs[0][1]) if self.production_outputs else "")
            self.ui.FGs_lnedit.setSuffix(f" Bf: {self.production_outputs[0][1]}" if self.production_outputs else "")
            self.ui.defect_lnedit.setText(str(self.production_outputs[0][2]) if self.production_outputs else "")
            self.ui.defect_lnedit.setSuffix(f" Bf: {self.production_outputs[0][2]}" if self.production_outputs else "")
            self.DT_context_menu = QtWidgets.QMenu(self)
            self.DT_context_menu.addAction("Add Record", self.add_record)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
        
        self.setup_signals()
    
    def setup_signals(self):
        self.ui.working_shift_lnedit.returnPressed.connect(lambda: self.calculate_oee("working_shift"))
        self.ui.breaktime_lnedit.returnPressed.connect(lambda: self.calculate_oee("breaktime"))
        self.ui.setuptime_lnedit.returnPressed.connect(lambda: self.calculate_oee("setuptime"))
        self.ui.FGs_lnedit.returnPressed.connect(lambda: self.calculate_oee("FGs"))
        self.ui.defect_lnedit.returnPressed.connect(lambda: self.calculate_oee("defect"))
        self.ui.OEEsub_groupBox_2.customContextMenuRequested.connect(lambda pos: self.show_downtime_context_menu(pos))

    @QtCore.pyqtSlot()
    def calculate_oee(self,edit_object=None):
        if not self.ui.working_shift_lnedit.text() or not self.ui.FGs_lnedit.text() or not self.ui.defect_lnedit.text():
            return
        def re_cal_break_setup_time(working_shift_hours):
            if working_shift_hours < 12:
                break_time = 45
                setup_time = 10
            elif working_shift_hours < 16:
                break_time = 60
                setup_time = 15
            elif working_shift_hours < 24:
                break_time = 120
                setup_time = 30
            return break_time, setup_time
        try:
            if edit_object == "working_shift" or edit_object == "downtime_record" or edit_object == "breaktime" or edit_object == "setuptime":
                if edit_object == "working_shift":
                    working_shift = float(self.ui.working_shift_lnedit.text().strip())
                    if working_shift <= 0 or working_shift > 24:
                        QtWidgets.QMessageBox.warning(self, "Warning", "Please enter a valid number for working shift (1-24).")
                        self.ui.working_shift_lnedit.clear()
                        return
                    self.data["working_shift_hours"] = working_shift
                    break_time, setup_time = re_cal_break_setup_time(working_shift)
                    self.data["break_time"] = break_time
                    self.data["setup_time"] = setup_time
                    self.data["planed_time"] = working_shift*60 - break_time - setup_time
                    self.ui.breaktime_lnedit.setText(str(break_time))
                    self.ui.setuptime_lnedit.setText(str(setup_time))
                    item = QtGui.QStandardItem(str(working_shift))
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.data_model.setItem(1, 1, item)
                    item = QtGui.QStandardItem(str(break_time))
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.data_model.setItem(2, 1, item)
                    item = QtGui.QStandardItem(str(setup_time))
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.data_model.setItem(3, 1, item)
                    self.data_model.setItem(4, 1, QtGui.QStandardItem(str(self.data["planed_time"])))
                elif edit_object == "downtime_record":
                    self.data["total_loss_mins"] = sum(record.total_loss_time for record in self.downtime_records)
                    item = QtGui.QStandardItem(str(self.data["total_loss_mins"]))
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.data_model.setItem(5, 1, item)
                # elif edit_object == "breaktime":
                #     self.data["break_time"] = float(self.ui.breaktime_lnedit.text().strip())
                #     item = QtGui.QStandardItem(str(self.data["break_time"]))
                #     item.setTextAlignment(QtCore.Qt.AlignCenter)
                #     self.data_model.setItem(2, 1, item)
                # elif edit_object == "setuptime":
                #     self.data["setup_time"] =  float(self.ui.setuptime_lnedit.text().strip())
                #     item = QtGui.QStandardItem(str(self.data["setup_time"]))
                #     item.setTextAlignment(QtCore.Qt.AlignCenter)
                #     self.data_model.setItem(3, 1, item)
                available_time = self.data["planed_time"] - self.data["total_loss_mins"]
                self.data["available_time_mins"] = available_time
                item = QtGui.QStandardItem(str(available_time))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.data_model.setItem(6, 1, item)
                self.data["availability_percentage"] = self.data["available_time_mins"] / (self.data["planed_time"] ) * 100 if self.data["planed_time"] > 0 else 0
                item = QtGui.QStandardItem(f"{self.data['availability_percentage']:.2f}")
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.data_model.setItem(9, 1, item)
            elif edit_object == "FGs" or edit_object == "defect":
                if edit_object == "FGs":
                    fgs_output = float(self.ui.FGs_lnedit.text().strip())
                    self.data["fgs_output_pcs"] = fgs_output
                    item = QtGui.QStandardItem(str(fgs_output))
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.data_model.setItem(7, 1, item)
                else:
                    defect_output = float(self.ui.defect_lnedit.text().strip())
                    self.data["defect_pcs"] = defect_output
                    item = QtGui.QStandardItem(str(defect_output))
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.data_model.setItem(8, 1, item)
                self.data["quality_percentage"] = self.data["fgs_output_pcs"] / (self.data["fgs_output_pcs"] + self.data["defect_pcs"]) * 100 if (self.data["fgs_output_pcs"] + self.data["defect_pcs"]) > 0 else 0
                item = QtGui.QStandardItem(f"{self.data['quality_percentage']:.2f}")
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.data_model.setItem(11, 1, item)
            self.data["performance_percentage"] = self.cycle_time * (self.data["fgs_output_pcs"] + self.data["defect_pcs"]) / (self.data["available_time_mins"]*60) * 100 if self.data["available_time_mins"] > 0 else 0
            item = QtGui.QStandardItem(f"{self.data['performance_percentage']:.2f}")
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.data_model.setItem(10, 1, item)
            self.data["OEE_percentage"] = (self.data["availability_percentage"] * self.data["performance_percentage"] * self.data["quality_percentage"]) / 10000
            item = QtGui.QStandardItem(f"{self.data['OEE_percentage']:.2f}")
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.data_model.setItem(12, 1, item)
        except ValueError:
            QtWidgets.QMessageBox.warning(self, "Warning", "Please enter a valid number for working shift.")
            return             

        
    @QtCore.pyqtSlot()
    def show_downtime_context_menu(self, pos):
        self.DT_context_menu.exec_(self.ui.OEEsub_groupBox_2.mapToGlobal(pos))

    @QtCore.pyqtSlot()
    def add_record(self):
        try:
            new_record = (None, self.data["production_date"], None, None, None, None, None, self.downtime_records[0].machine_code if self.downtime_records else self.machine_code, self.data["line_name"])
            downtime_record = self.downtime_record_widgetItem(new_record)
            downtime_record.delete_requested.connect(lambda widget=downtime_record: self.remove_downtime_record(widget))
            downtime_record.edit_completed.connect(lambda widget=downtime_record: self.calculate_oee(edit_object="downtime_record"))
            self.downtime_records.append(downtime_record)
            list_item = QtWidgets.QListWidgetItem(self.ui.downtime_records_widget)
            list_item.setSizeHint(QtCore.QSize(0, 160))
            self.ui.downtime_records_widget.addItem(list_item)
            self.ui.downtime_records_widget.setItemWidget(list_item, downtime_record)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to add new downtime record: {e}")
    
    @QtCore.pyqtSlot()    
    def remove_downtime_record(self, widget):
        list_widget = self.ui.downtime_records_widget
        for i in range(list_widget.count()):
            item = list_widget.item(i)
            if list_widget.itemWidget(item) is widget:
                list_widget.takeItem(i)
                if widget in self.downtime_records:
                    if widget.id is not None:
                        try:
                            self.database.query(sql=''' DELETE FROM `downtime_records` WHERE downtime_record_id = :id ''', params={'id': widget.id})
                        except Exception as e:
                            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to delete downtime record: {e}")
                            return
                    self.downtime_records.remove(widget)
                    self.calculate_oee(edit_object="downtime_record")
                break

    def keyPressEvent(self, event):
        if event.key() in (QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter):
            event.ignore()
        else:
            super().keyPressEvent(event)

class OEE_Other_Data(QtWidgets.QDialog):
    def __init__(self, parent = None, database=None, data=None):
        super().__init__(parent)
        self.ui = UI_OEE_Other_Data()
        self.ui.setupUi(self)
        self.database = database
        self.parent = parent
        self.operation_edit_item_dict = {"new":[] , "edit":[]}
        try:
            areas = self.database.query(sql='''SELECT downtime_area_name FROM downtime_areas''')
            self.ui.OEE_OD_area_cbb.addItems([row[0] for row in areas])
            models = self.database.query(sql='''SELECT model_name 
                                                            FROM `product_models_oee` as pmo
                                                            JOIN `departments` as d
                                                            ON pmo.department_id = d.department_id
                                                            JOIN downtime_areas as da
                                                            ON da.department_id = d.department_id
                                                            WHERE da.downtime_area_name = :area_name;''', params={"area_name": areas[0][0]})
            self.ui.OEE_OD_model_cbb.addItems([row[0] for row in models])
            lines = self.database.query(sql='''SELECT DISTINCT pl.line_name
                                                            FROM `production_lines` as pl
                                                            JOIN `production_output` as po
                                                            ON pl.line_id = po.line_id
                                                            WHERE  MONTH(po.production_date) = :month AND YEAR(po.production_date) = :year;''', 
                                                            params={ "month": self.ui.OEE_OD_period_edit.date().month(), "year": self.ui.OEE_OD_period_edit.date().year()})
            self.ui.OEE_OD_line_cbb.addItems([row[0] for row in lines])
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
        self.setWindowIcon(QtGui.QIcon(resource_path("Icons/OEE.ico")))
        header = ["Cycle Time ID", "Model Name", "Machine Code", "Cycle Time (s)", "Recorded At", "Notes"]
        self.data_cycle_time_model = QtGui.QStandardItemModel(0, 5)
        self.data_cycle_time_model.setHorizontalHeaderLabels(header)
        header = ["Operation ID", "Line Name", "Operation\nDate", "Operation\nHours", "Break Time\n(min)", "Setup Time\n(min)", "Change\nModel", "Change\nFrom", "OK Qty"]
        self.data_operation_time_model = QtGui.QStandardItemModel(0 , 9)
        self.data_operation_time_model.setHorizontalHeaderLabels(header)
        self.tab_changing(self.ui.Tab_widget.currentIndex())

        self.show_data()
        self.setup_signals()
        
    def setup_signals(self):
        self.ui.cancel_btn.clicked.connect(lambda: self.reject())
        self.ui.Tab_widget.currentChanged.connect(lambda index: self.tab_changing(index))
        self.ui.OEE_OD_calendar_widget.currentPageChanged.connect(lambda year,
                              month: self.parent.update_date_from_calendar(year, month, self.ui.OEE_OD_period_edit))
        self.ui.machine_cycle_table.customContextMenuRequested.connect(lambda pos: self._cycle_table_context_menu(pos))
        self.ui.confirm_btn.clicked.connect(self.accept_action)
        self.ui.line_operation_table.doubleClicked.connect(lambda index: self.edit_operation_time_record(index))
        self.ui.OEE_OD_area_cbb.currentTextChanged.connect(lambda: self.filter_process())
        self.ui.OEE_OD_model_cbb.currentTextChanged.connect(lambda: self.show_data())
        self.ui.OEE_OD_line_cbb.currentTextChanged.connect(lambda: self.show_data())
        self.ui.OEE_OD_period_edit.dateChanged.connect(lambda: self.show_data())
    
    @QtCore.pyqtSlot()
    def tab_changing(self, index):
        if index == 0:
            self.ui.OEE_OD_groupBox_3.setHidden(False)
            self.ui.OEE_OD_groupBox.setHidden(True)
            self.ui.OEE_OD_groupBox_2.setHidden(True)
        else:
            self.ui.OEE_OD_groupBox_3.setHidden(True)
            self.ui.OEE_OD_groupBox.setHidden(False)
            self.ui.OEE_OD_groupBox_2.setHidden(False)
        self.show_data()
    
    @QtCore.pyqtSlot()
    def show_data(self):
        current_page = self.ui.Tab_widget.currentIndex()
        try:
            if current_page == 0:
                self.data_cycle_time_model.removeRows(0, self.data_cycle_time_model.rowCount())
                model_name = self.ui.OEE_OD_model_cbb.currentText()
                cycle_time_log = self.database.query(sql='''SELECT mct.cycle_time_id, pmo.model_name, m.machine_code, mct.cycle_time_seconds, mct.create_at, mct.Notes
                                                        FROM `machine_cycle_times` as mct
                                                        JOIN `product_models_oee` as pmo ON mct.model_id = pmo.model_id
                                                        JOIN `machines` as m ON mct.machine_id = m.machine_id
                                                        WHERE pmo.model_name = :model_name
                                                        ORDER BY mct.create_at DESC
                                                        LIMIT 10;''',
                                                        params={"model_name": model_name})
                self.cycle_time_df = pd.DataFrame(cycle_time_log, columns=["id", "model_name", "machine_code", "cycle_time_seconds", "create_at", "notes"])
                for row in range(self.cycle_time_df.shape[0]):
                    for col in range(self.cycle_time_df.shape[1]):
                        if col == 4:
                            item = QtGui.QStandardItem(self.cycle_time_df.iat[row, col].strftime("%Y-%m-%d") if self.cycle_time_df.iat[row, col] else "")
                        else:
                            item = QtGui.QStandardItem(str(self.cycle_time_df.iat[row, col]) if self.cycle_time_df.iat[row, col] is not None else "")
                        item.setTextAlignment(QtCore.Qt.AlignCenter)
                        item.setEditable(False)
                        item.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Normal))
                        self.data_cycle_time_model.setItem(row, col, item)
                
                self.ui.machine_cycle_table.setModel(self.data_cycle_time_model)
                self.ui.machine_cycle_table.setColumnHidden(0, True)
                self.ui.machine_cycle_table.setColumnHidden(1, True)
            else:
                line_name = self.ui.OEE_OD_line_cbb.currentText()
                month = self.ui.OEE_OD_period_edit.date().month()
                year = self.ui.OEE_OD_period_edit.date().year()
                self.data_operation_time_model.removeRows(0, self.data_operation_time_model.rowCount())
                operetion_time = self.database.query(sql = '''  SELECT lot.operation_id, pl.line_name, po.production_date, 
                                                                    lot.operation_hours, lot.break_time, lot.setup_time, 
                                                                    lot.change_model, lot.change_from , po.OK_qty
                                                                FROM production_output AS po
                                                                JOIN production_lines AS pl ON po.line_id = pl.line_id
                                                                LEFT JOIN line_operation_times AS lot 
                                                                    ON po.line_id = lot.line_id AND po.production_date = lot.operation_date AND po.model_name = lot.model_running
                                                                WHERE MONTH(po.production_date) = :month AND YEAR(po.production_date) = :year
                                                                AND pl.line_name = :line_name AND po.`OK_qty`> 100
                                                                ORDER BY po.production_date ASC;''',
                                                                params={"line_name": line_name, "month": month, "year": year})
                self.operation_time_df = pd.DataFrame(operetion_time, columns=["operation_id", "line_name", "operation_date", "operation_hours", "break_time", "setup_time", "model", "change_from", "OK_qty"])
                for row in range(self.operation_time_df.shape[0]):
                    for col in range(self.operation_time_df.shape[1]):
                        if col == 2:
                            item = QtGui.QStandardItem(self.operation_time_df.iat[row, col].strftime("%Y-%m-%d") if self.operation_time_df.iat[row, col] else "")
                        elif col == 7:
                            val = self.operation_time_df.iat[row, col]
                            if val is None or (hasattr(pd, 'NaT') and val is pd.NaT):
                                item = QtGui.QStandardItem("")
                            elif isinstance(val, pd.Timedelta) or hasattr(val, 'total_seconds'):
                                total_seconds = int(val.total_seconds())
                                hours, remainder = divmod(abs(total_seconds), 3600)
                                minutes, seconds = divmod(remainder, 60)
                                item = QtGui.QStandardItem(f"{hours:02d}:{minutes:02d}:{seconds:02d}")
                            elif hasattr(val, 'strftime'):
                                item = QtGui.QStandardItem(val.strftime("%H:%M:%S"))
                            else:
                                item = QtGui.QStandardItem(str(val))
                        elif col == 4 or col == 5:
                            item = QtGui.QStandardItem(str(self.operation_time_df.iat[row, col]) if self.operation_time_df.iat[row, col] is not None else "")
                            item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                        else:
                            item = QtGui.QStandardItem(str(self.operation_time_df.iat[row, col]) if self.operation_time_df.iat[row, col] is not None else "")
                        item.setTextAlignment(QtCore.Qt.AlignCenter)
                        # item.setEditable(False)
                        item.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Normal))
                        self.data_operation_time_model.setItem(row, col, item)
                self.ui.line_operation_table.setModel(self.data_operation_time_model)
                self.ui.line_operation_table.setColumnHidden(0, True)
                self.ui.line_operation_table.setColumnHidden(1, True)
                self.ui.line_operation_table.horizontalHeader().setSectionResizeMode(6, QtWidgets.QHeaderView.Fixed)
                self.ui.line_operation_table.setColumnWidth(6, 250)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load data: {e}")
    
    @QtCore.pyqtSlot()
    def filter_process(self):
        try:
            area_name = self.ui.OEE_OD_area_cbb.currentText()
            models = self.database.query(sql='''SELECT model_name 
                                                            FROM `product_models_oee` as pmo
                                                            JOIN `departments` as d
                                                            ON pmo.department_id = d.department_id
                                                            JOIN downtime_areas as da
                                                            ON da.department_id = d.department_id
                                                            WHERE da.downtime_area_name = :area_name;''', params={"area_name": area_name})
            self.ui.OEE_OD_model_cbb.clear()
            self.ui.OEE_OD_model_cbb.addItems([row[0] for row in models])
            month = self.ui.OEE_OD_period_edit.date().month()
            year = self.ui.OEE_OD_period_edit.date().year()
            lines = self.database.query(sql='''SELECT DISTINCT pl.line_name
                                                            FROM `production_lines` as pl
                                                            JOIN `production_output` as po
                                                            ON pl.line_id = po.line_id
                                                            WHERE MONTH(po.production_date) = :month AND YEAR(po.production_date) = :year;''', 
                                                            params={ "month": month, "year": year})
            self.ui.OEE_OD_line_cbb.clear()
            self.ui.OEE_OD_line_cbb.addItems([row[0] for row in lines])
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to filter models: {e}")
        

    @QtCore.pyqtSlot()
    def _cycle_table_context_menu(self, pos):
        index = self.ui.machine_cycle_table.indexAt(pos)
        menu = QtWidgets.QMenu(self.ui.machine_cycle_table)
        add_act = menu.addAction("Add")
        add_act.triggered.connect(lambda: self.add_cycle_time_record(index))
        if index.isValid():
            menu.addSeparator()
            del_act = menu.addAction("Delete")
            del_act.triggered.connect(lambda: self.delete_cycle_time_record(index))
        menu.exec_(self.ui.machine_cycle_table.viewport().mapToGlobal(pos))

    @QtCore.pyqtSlot()
    def add_cycle_time_record(self,index):
        self.edit_cycle_time_flag = True
        try:
            self.data_cycle_time_model.insertRow(self.cycle_time_df.shape[0])
            new_record = {"id": None, "model_name": self.ui.OEE_OD_model_cbb.currentText(), "machine_code": self.cycle_time_df.iat[0, 2], "cycle_time_seconds": None, "create_at": dt.datetime.now().date(), "notes": None}
            self.cycle_time_df = pd.concat([pd.DataFrame([new_record]), self.cycle_time_df], ignore_index=True)
            for col in range(self.cycle_time_df.shape[1]):
                item = QtGui.QStandardItem(str(self.cycle_time_df.iat[0, col]) if self.cycle_time_df.iat[0, col] is not None else "")
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                item.setEditable(True)
                item.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Normal))
                item.setBackground(QtGui.QColor(34, 57, 40, 30))
                self.data_cycle_time_model.setItem(len(self.cycle_time_df) - 1, col, item)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to add cycle time record: {e}")


    @QtCore.pyqtSlot()
    def delete_cycle_time_record(self,index):
        try:
            record_id = self.cycle_time_df.iat[index.row(), 0]
            if record_id is not None:
                reply = QtWidgets.QMessageBox.question(
                    self, "Confirm", "Are you sure you want to delete this cycle time record?\nThis action cannot be undone.",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
                if reply != QtWidgets.QMessageBox.Yes:
                    return
                self.database.query(sql=''' DELETE FROM `machine_cycle_times` WHERE cycle_time_id = :id ''', params={'id': record_id})
            self.data_cycle_time_model.removeRow(index.row())
            self.cycle_time_df = self.cycle_time_df.drop(index.row()).reset_index(drop=True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to delete cycle time record: {e}")
    
    @QtCore.pyqtSlot()
    def edit_operation_time_record(self,index):
        if index.column() in [4, 5]:
            return
        self.edit_operation_time_flag = True
        row = index.row()
        try:
            operation_id = self.operation_time_df.iat[index.row(), 0]
            for col in range(3,8):
                if col == 4 or col == 5:
                    continue
                item = self.data_operation_time_model.item(row, col)
                item.setEditable(True)
                item.setBackground(QtGui.QColor(34, 57, 40, 30))
                item.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Normal))
            if operation_id is None:
                self.operation_edit_item_dict["new"].append(row)
            else:
                self.operation_edit_item_dict["edit"].append(row)
                
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to edit operation time record: {e}")
            return

    @QtCore.pyqtSlot()
    def accept_action(self):
        insert_params = []
        affected_points = []
        def update_cycle_time():
            for row in range(self.cycle_time_df.shape[0]):
                record_id = self.data_cycle_time_model.index(row, 0).data()
                if not record_id:
                    model_name = self.data_cycle_time_model.index(row, 1).data()
                    machine_code = self.data_cycle_time_model.index(row, 2).data()
                    cycle_time_seconds = self.data_cycle_time_model.index(row, 3).data()
                    note = self.data_cycle_time_model.index(row, 5).data()
                    if cycle_time_seconds is not None:
                        insert_params.append({'model_name': model_name, 'machine_code': machine_code, 'cycle_time_seconds': cycle_time_seconds, 'notes': note})
            if insert_params:
                self.database.executemany(sql=''' INSERT INTO `machine_cycle_times` (model_id, machine_id, cycle_time_seconds, notes)
                                                    VALUES ((SELECT model_id FROM product_models_oee WHERE model_name = :model_name),
                                                    (SELECT machine_id FROM machines WHERE machine_code = :machine_code),
                                                    :cycle_time_seconds,:notes)''', params_list=insert_params)

        def rebuild_after_edit( affected_points):
            if not affected_points:
                return

            line_name = self.ui.line_operation_table.model().index(0, 1).data()

            line_id = self.database.query(
                sql="SELECT line_id FROM production_lines WHERE line_name = :line_name",
                params={'line_name': line_name}
            )[0][0]
            unique_points = {(p['operation_date'], p['change_from']) for p in affected_points}

            points_sorted = sorted(
                unique_points,
                key=lambda x: (str(x[0]), x[1] if x[1] is not None else '00:00:00')
            )
            call_params = [
                {'line_id': line_id, 'operation_date': d, 'change_from': cf}
                for (d, cf) in points_sorted
            ]
            self.database.executemany(
                sql="CALL sp_rebuild_model_running(:line_id, :operation_date, :change_from)",
                params_list=call_params
            )
        def update_operations():
            line_name = self.ui.line_operation_table.model().index(0, 1).data()
            if self.operation_edit_item_dict["new"]:
                params_list = []
                for row in set(self.operation_edit_item_dict["new"]):
                    operation_date = self.ui.line_operation_table.model().index(row, 2).data()
                    operation_hours = self.ui.line_operation_table.model().index(row, 3).data()
                    break_time = self.ui.line_operation_table.model().index(row, 4).data() 
                    setup_time = self.ui.line_operation_table.model().index(row, 5).data()
                    change_model = self.ui.line_operation_table.model().index(row, 6).data()
                    change_from = self.ui.line_operation_table.model().index(row, 7).data()
                    params_list.append({'line_name': line_name, 'operation_date': operation_date, 'operation_hours': operation_hours, 
                                        # 'break_time': break_time.strip() if break_time.strip() != "" else 0, 
                                        # 'setup_time': setup_time.strip() if setup_time.strip() != "" else 0, 
                                        'change_model': change_model.strip() if change_model.strip() != "" else None, 
                                        'change_from': change_from.strip() if change_from.strip() != "" else None})
                    
                self.database.executemany(sql = ''' INSERT INTO line_operation_times (line_id, operation_date, operation_hours, change_model, change_from)
                                                    VALUES ((SELECT line_id FROM production_lines WHERE line_name = :line_name), 
                                                            :operation_date, :operation_hours, :change_model, :change_from)'''
                                          , params_list=params_list)
            if self.operation_edit_item_dict["edit"]:
                params_list_update = []
                params_list_insert = []
                for row in set(self.operation_edit_item_dict["edit"]):
                    operation_id = self.operation_time_df.iat[row, 0]
                    operation_date = self.ui.line_operation_table.model().index(row, 2).data()
                    operation_hours = self.ui.line_operation_table.model().index(row, 3).data()
                    break_time = self.ui.line_operation_table.model().index(row, 4).data()
                    setup_time = self.ui.line_operation_table.model().index(row, 5).data()
                    change_model = self.ui.line_operation_table.model().index(row, 6).data()
                    change_from = self.ui.line_operation_table.model().index(row, 7).data()
                    cm = change_model.strip() if change_model.strip() != "" else None
                    cf = change_from.strip()  if change_from.strip()  != "" else None
                    affected_points.append({'operation_date': operation_date, 'change_from': cf}) 
                    if pd.isna(operation_id) or operation_id is None:
                        params_list_insert.append({'line_name': line_name, 'operation_date': operation_date, 'operation_hours': operation_hours, 
                                                            # 'break_time': break_time if break_time.strip() != "" else 0, 
                                                            # 'setup_time': setup_time if setup_time.strip() != "" else 0, 
                                                            'change_model': cm, 
                                                            'change_from': cf})
                        continue
                    params_list_update.append({'operation_id': operation_id, 'operation_date': operation_date, 'operation_hours': operation_hours, 
                                        # 'break_time': break_time if break_time.strip() != "" else 0, 
                                        # 'setup_time': setup_time if setup_time.strip() != "" else 0, 
                                        'change_model': cm, 
                                        'change_from': cf})
                    
                if params_list_update:
                    self.database.executemany(sql = ''' UPDATE line_operation_times SET operation_date = :operation_date, operation_hours = :operation_hours, change_model = :change_model, change_from = :change_from
                                                    WHERE operation_id = :operation_id''',
                                        params_list=params_list_update)
                if params_list_insert:
                    self.database.executemany(sql = ''' INSERT INTO line_operation_times (line_id, operation_date, operation_hours, change_model, change_from)
                                                    VALUES ((SELECT line_id FROM production_lines WHERE line_name = :line_name),
                                                            :operation_date, :operation_hours, :change_model, :change_from)'''
                                          , params_list=params_list_insert)
                rebuild_after_edit(affected_points)
        try:
            if hasattr(self, 'edit_cycle_time_flag') and hasattr(self, 'edit_operation_time_flag'):
                msgBox = QtWidgets.QMessageBox(self)
                msgBox.setWindowTitle("Update Content")
                msgBox.setText("What would you like to update?")
                cycle_time_btn = msgBox.addButton("Cycle Time", QtWidgets.QMessageBox.ActionRole)
                operations_btn = msgBox.addButton("Operations", QtWidgets.QMessageBox.ActionRole)
                both_btn = msgBox.addButton("Both", QtWidgets.QMessageBox.ActionRole)
                msgBox.exec()
                clicked_button = msgBox.clickedButton()
                if clicked_button == cycle_time_btn:
                    update_cycle_time()
                    QtWidgets.QMessageBox.information(self, "Success", "Cycle time records have been updated successfully.")
                elif clicked_button == operations_btn:
                    update_operations()
                    QtWidgets.QMessageBox.information(self, "Success", "Operation time records have been updated successfully.")
                elif clicked_button == both_btn:
                    update_cycle_time()
                    update_operations()
                    QtWidgets.QMessageBox.information(self, "Success", "Cycle time and operation time records have been updated successfully.")
                else:
                    return
            elif hasattr(self, 'edit_cycle_time_flag'):
                update_cycle_time()
                QtWidgets.QMessageBox.information(self, "Success", "Cycle time records have been updated successfully.")
            elif hasattr(self, 'edit_operation_time_flag'):
                update_operations()
                QtWidgets.QMessageBox.information(self, "Success", "Operation time records have been updated successfully.")
            else:
                return
        
            self.show_data()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to save cycle time record: {e}")

class GaugeChart(QtWidgets.QWidget):
    def __init__(self, value=0, height_scale=1.0, width_scale=1.0, parent=None):
        super().__init__(parent)
        self.value = max(0, min(100, value))
        self.height_scale = height_scale
        self.width_scale = width_scale

    def paintEvent(self, event):

        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)

        width_chart = self.width()*0.7*self.width_scale
        height_chart = self.height()*0.7*self.height_scale
        left_margin = (self.width() - width_chart) / 2
        top_margin = (self.height() - height_chart) / 2

        rect = QtCore.QRectF(
            left_margin ,
            top_margin,
            width_chart,
            height_chart*2
        )

        # Background Arc
        pen = QtGui.QPen(QtGui.QColor("#E0E0E0"), 20)
        pen.setCapStyle(QtCore.Qt.PenCapStyle.FlatCap)
        painter.setPen(pen)

        painter.drawArc(
            rect,
            180 * 16,
            -180 * 16
        )

        # Value Arc
        pen = QtGui.QPen(QtGui.QColor("#060AF3"), 20)
        pen.setCapStyle(QtCore.Qt.PenCapStyle.FlatCap)
        painter.setPen(pen)

        painter.drawArc(
            rect,
            180 * 16,
            -int(180 * 16 * (self.value / 100))
        )

        painter.setPen(QtGui.QColor("#464646"))


        font3 = QtGui.QFont()
        font3.setFamily("Agency FB")
        font3.setPointSize(28)
        font3.setBold(True)
        painter.setFont(font3)
        
        painter.drawText(
            self.rect(),
            QtCore.Qt.AlignBottom|QtCore.Qt.AlignHCenter,
            f"{self.value:.2f}%"
        )

class Downtime_Detail_dashboard(QtWidgets.QDialog):
    closed = QtCore.pyqtSignal(object)
    def __init__(self, parent=None, database=None, title="Downtime Detail Report Of ", catagory=None, data_dict = {}, button_style_call_back=None, change_time_format_call_back=None, KPI_chart_call_back=None, icon_from_path_call_back=None, extract_content_call_back = None):
        super().__init__(parent)
        self.ui = UI_DT_detail_report()
        self.ui.setupUi(self)
        self.data_dict = data_dict
        if self.data_dict.get("group_col", "") == "Line Name":
            title = title + "Line " + catagory
        elif self.data_dict.get("group_col", "") == "Machine Code":
            title = title + self.data_dict.get("name", "") + " " + catagory
        else:
            title = title + self.data_dict.get("name", "") + " " +catagory
        self.day, self.month, self.year = None, None, None
        self.have_action = False
        if data_dict.get("view_by", "") == "day":
            date = data_dict.get("target", "")
            title = title + " | Date: " + str(date)
            self.year, self.month, self.day = map(int, date.split('-'))
            date_condition = " AND for_day = :day AND for_month = :month AND for_year = :year"
        elif data_dict.get("view_by", "") == "week":
            date = data_dict.get("target", "")
            title = title + " | Week: " + str(date) + " of " + str(data_dict.get("year", ""))
        else:
            date = data_dict.get("target", "")
            title = title + " | Month: " + str(date) + "/" + str(data_dict.get("year", ""))
            self.month = date
            self.year = data_dict.get("year", "")
            date_condition = " AND for_month = :month AND for_year = :year"
        self.ui.title_lbl.setText(title)
        self.catagory = catagory
        self.database = database
        self.button_style_call_back = button_style_call_back
        self.change_time_format_call_back = change_time_format_call_back
        self.DT_KPI_chart = KPI_chart_call_back
        self.icon_from_path_call_back = icon_from_path_call_back
        self.extract_content = extract_content_call_back
        self.setWindowIcon(QtGui.QIcon(resource_path("Icons/Downtime.ico")))
        self.ui.Detail_total_downtime_value.setText(self.data_dict["total_downtime"])
        if self.data_dict.get("group_col", "") in ("Machine Code","Line Name"):
            self.ui.Detail_card3_label.setText("MTTR")
            self.ui.Detail_card4_label.setText("MTBF")
            self.ui.Detail_card3_value.setText(self.data_dict["mttr"])
            self.ui.Detail_card4_value.setText(self.data_dict["mtbf"])
            root_cause = self.Column_chart(widget= self.ui.Detail_left_chart,
                          data_df = self.data_dict["data_of_category"],
                          title = "Downtime By Error Code",
                          show_by = "Error Code")
            self.ui.Detail_left_label.setText("Downtime By Error")
            if self.data_dict.get("group_col", "") == "Line Name":
                top_trend = self.Column_chart(widget= self.ui.Detail_right_chart,
                                data_df = self.data_dict["data_of_category"],
                                title = "Downtime By Machine Code",
                                show_by = "Machine Code")
                self.ui.Detail_right_label.setText("Downtime By Machine")
            if self.data_dict.get("group_col", "") == "Machine Code":
                self.Column_chart(widget= self.ui.Detail_right_chart,
                                data_df = self.data_dict["data_of_category"],
                                title = "Downtime By Line Name",
                                show_by = "Line Name")
                top_trend = None
                self.ui.Detail_right_label.setText("Downtime By Line")
            self.DT_KPI_chart(widget=self.ui.KPI_MTTR_chart, value=self.time_to_minute(self.data_dict.get("mttr", 0)), target_value=self.time_to_minute(self.data_dict.get("mttr_target", 0)), previous_value=0, label="MTTR")
            self.DT_KPI_chart(widget=self.ui.KPI_MTBF_chart, value=float(self.time_to_minute(self.data_dict.get("mtbf", 0))), target_value=self.time_to_minute(self.data_dict.get("mtbf_target", 0)), previous_value=0, label="MTBF")
            if self.data_dict.get("group_col", "") == "Line Name":
                condition_category = f"AND pl.line_name = :category_value"
            else:
                condition_category = f"AND m.machine_code = :category_value"
            action = self.database.query(sql=f'''SELECT datc.action_id, datc.action_content, datc.action_report_link FROM downtime_actions as datc
                                                JOIN downtime_areas AS da ON datc.downtime_area_id = da.downtime_area_id
                                                LEFT JOIN production_lines AS pl ON datc.line_id = pl.line_id
                                                LEFT JOIN machines AS m ON datc.machine_id = m.machine_id
                                                LEFT JOIN error_codes_list AS ecl ON datc.error_code = ecl.error_code
                                                WHERE da.downtime_area_name = :area_name {condition_category} {date_condition} AND datc.action_for = "DT";''', params={"area_name": self.data_dict.get("area_name", ""), "category_value": self.data_dict.get("selected_category", ""), "day": self.day, "month": self.month, "year": self.year})
        else:
            self.resize(900,800)
            self.ui.Detail_card3_frame.setMinimumWidth(0)
            self.ui.Detail_card3_value.setMinimumWidth(self.ui.Detail_card3_frame.width())
            self.ui.Detail_keyinsights_frame.hide()
            root_cause = None
            top_trend = None
            self.ui.Detail_card3_label.setText("Events Count")
            self.ui.Detail_card3_value.setText(self.data_dict["events_count"])
            self.ui.Detail_card4_frame.hide()
            self.Column_chart(widget= self.ui.Detail_left_chart,
                                data_df = self.data_dict["data_of_category"],
                                title = "Downtime By Machine Code",
                                show_by = "Machine Code")
            self.ui.Detail_left_label.setText("Downtime By Machine")
            self.Column_chart(widget= self.ui.Detail_right_chart,
                                data_df = self.data_dict["data_of_category"],
                                title = "Downtime By Line Name",
                                show_by = "Line Name")
            self.ui.Detail_right_label.setText("Downtime By Line")
            self.ui.KPI_MTTR_chart.hide()
            self.ui.KPI_MTBF_chart.hide()
            condition_category = f"AND ecl.error_code = :category_value"
            action = self.database.query(sql=f'''SELECT datc.action_id, datc.action_content, datc.action_report_link FROM downtime_actions as datc
                                                JOIN downtime_areas AS da ON datc.downtime_area_id = da.downtime_area_id
                                                LEFT JOIN production_lines AS pl ON datc.line_id = pl.line_id
                                                LEFT JOIN machines AS m ON datc.machine_id = m.machine_id
                                                LEFT JOIN error_codes_list AS ecl ON datc.error_code = ecl.error_code
                                                WHERE da.downtime_area_name = :area_name {condition_category} {date_condition} AND datc.action_for = "DT";''', params={"area_name": self.data_dict.get("area_name", ""), "category_value": self.data_dict.get("selected_category", ""), "day": self.day, "month": self.month, "year": self.year})
        if action:
            self.have_action = True
            self.action_id = action[0][0]
            self.ui.action_text.setText(action[0][1])
            action_links = json.loads(action[0][2]) if action[0][2] else []
            for link in action_links:
                self.ui.action_text.append(f'<a href="{link["file_path"]}" style="color:#007acc; font-size:10px; text-decoration:underline;">'
                                            f'{link["file_name"]}</a> ')    
        self.ui.keyinsights_text.setText(self.Generate_key_insights_en(data = self.data_dict, root_cause = root_cause, top_trend = top_trend))
        self.Gauge_chart(self.data_dict["percentage"])
        self.Line_chart(widget= self.ui.Detail_bytime_chart, title="Downtime Over Time",view_by = "day" ,show_by=self.data_dict.get("group_col", ""))
        self.setup_signals()

    def setup_signals(self):
        self.ui.ByDate_btn.clicked.connect(lambda: self.Line_chart(widget= self.ui.Detail_bytime_chart, title="Downtime Over Time",view_by = "day" ,show_by=self.data_dict.get("group_col", "")))
        self.ui.ByMonth_btn.clicked.connect(lambda: self.Line_chart(widget= self.ui.Detail_bytime_chart, title="Downtime Over Time",view_by = "month" ,show_by=self.data_dict.get("group_col", "")))
        self.ui.action_text.dropEvent = lambda event: self.action_text_drop_event(event)
        self.ui.action_text.mousePressEvent = self.action_text_mouse_press_event
        self.ui.action_text.setMouseTracking(True)
        self.ui.action_text.mouseMoveEvent = self.mouseMoveEvent
    
    def action_before_closed(self):
        action_content = self.extract_content(self.ui.action_text)
        comment = ""
        link_list = []
        for key, value in enumerate(action_content):
            if value['type'] == 'link':
                link_list.append({
                    "file_name": value['text'],
                    "file_path": value['href']
                }
            )
            else:
                comment += f" {value['text']}"
        link_list = json.dumps(link_list, ensure_ascii=False)
        if self.data_dict.get("group_col", "") == "Machine Code":
            category = f"machine_id"
            part_condition = "(SELECT machine_id FROM machines WHERE machine_code = :category_value)"
        elif self.data_dict.get("group_col", "") == "Line Name":
            category = f"line_id"
            part_condition = "(SELECT line_id FROM production_lines WHERE line_name = :category_value)"
        else:
            category = f"error_code"
            part_condition = ":category_value"    
        if action_content:
            if self.have_action:
                self.database.query(sql=f''' UPDATE downtime_actions 
                                            SET action_content = :action_content, action_report_link = :action_report_link
                                            WHERE action_id = :action_id''', params={"action_content": comment, "action_report_link": link_list, "action_id": self.action_id})
            else:
                self.database.query(sql=f''' INSERT INTO downtime_actions (action_content, action_report_link, downtime_area_id, {category}, for_day, for_month, for_year, action_for)
                                            VALUES (:action_content, :action_report_link, (SELECT downtime_area_id FROM downtime_areas WHERE downtime_area_name = :area_name),
                                            {part_condition}, :day, :month, :year, "DT");''', params={"action_content": comment, "action_report_link": link_list, "area_name": self.data_dict.get("area_name", ""), "category_value": self.data_dict.get("selected_category", ""), "day": self.day, "month": self.month, "year": self.year})
        else:
            if self.have_action:
                self.database.query(sql=f''' DELETE FROM downtime_actions WHERE action_id = :action_id''', params={"action_id": self.action_id})

    def reject(self):
        self.action_before_closed()
        super().reject()
    
    def Gauge_chart(self, percent_value):
        old_layout = self.ui.Detail_percent_chart.layout()
        if old_layout is not None:
            while old_layout.count():
                item = old_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
        else:
            new_layout = QtWidgets.QVBoxLayout()
            new_layout.setContentsMargins(0, 0, 0, 0)
            new_layout.setSpacing(0)
            self.ui.Detail_percent_chart.setLayout(new_layout)
        if self.data_dict.get("group_col", "") in ("Machine Code","Line Name"):
            gauge = GaugeChart(value=percent_value)
        else:
            gauge = GaugeChart(value=percent_value, width_scale=0.8)
        self.ui.Detail_percent_chart.layout().addWidget(gauge)

    def Column_chart(self, widget ,data_df, title = "", show_by = None):
        old_layout = widget.layout()
        if old_layout is not None:
            while old_layout.count():
                item = old_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
        else:
            new_layout = QtWidgets.QVBoxLayout()
            new_layout.setContentsMargins(0, 0, 0, 0)
            new_layout.setSpacing(0)
            widget.setLayout(new_layout)
        layout = widget.layout()
        pivot_df = (
            data_df.groupby(show_by)["Total Loss Time"]
            .sum()
            .reset_index()
            .sort_values(by="Total Loss Time", ascending=False)
        )
        chart_font = QtGui.QFont("Comic Sans MS", 8)
        chart_font.setStyleStrategy(QtGui.QFont.PreferAntialias)
        chart_font.setHintingPreference(QtGui.QFont.PreferFullHinting)
        chart_font.setBold(True)
        categories = pivot_df[show_by]
        values = pivot_df["Total Loss Time"]
        x = np.arange(len(categories))
        if show_by == "Machine Code":
            _angle = -45
            dx = -15
            dy = 5
        elif show_by == "Error Code" and len(categories) > 10:
            _angle = -40
            dx = 0
            dy = 0
        else:
            _angle = 0
            dx = 0
            dy = 0
        x_axis = RotatedAxisItem(
                angle=_angle, dx=dx, dy=dy, orientation='bottom')
        plot = pg.PlotWidget(axisItems={'bottom': x_axis})
        plot.setBackground(None)
        plot.showGrid(x=False, y=False)
        bar_graph = pg.BarGraphItem(x=range(len(categories)), height=values, width=0.6, brush='b', pen=pg.mkPen(None))
        plot.addItem(bar_graph)
        plot.getAxis('bottom').setTextPen(pg.mkPen(color='#474747'))
        plot.getAxis('bottom').setTickFont(chart_font)
        plot.getAxis('bottom').setTicks([list(zip(range(len(categories)), categories))])
        left_axis = plot.getAxis('left')
        left_axis.setLabel("Time (minutes)")
        left_axis.setStyle(tickTextOffset=10, tickLength=0)
        left_axis.setPen(None)
        left_axis.setTextPen(pg.mkPen(color='#474747'))
        left_axis.setTickFont(chart_font)

        # left_axis.hide()
        y_min = 0
        y_max = max(values) * 1.1 if len(values) > 0 else 1
        plot.setYRange(y_min, y_max*1.2, padding=0)
        x_axis = plot.getAxis('bottom')
        if _angle == -45:
            x_axis.setStyle(tickTextHeight=10,tickTextOffset=15)
            x_axis.setHeight(50)
        if len(categories) <= 20:
            for i, val in enumerate(values):
                text_item = pg.TextItem(text=str(val), color="black", anchor=(0.5, 1))
                text_item.setFont(QtGui.QFont("Comic Sans MS", 8, QtGui.QFont.Bold))
                offset = max(values) * 0.02 if max(values) > 0 else 1
                text_item.setPos(i, val + offset)
                plot.addItem(text_item)
        else:
            for i, val in enumerate(values):
                text_item = pg.TextItem(text=str(val), color="black", anchor=(0.5, 1), angle=-90)
                text_item.setFont(QtGui.QFont("Comic Sans MS", 8, QtGui.QFont.Bold))
                offset = max(values) * 0.15 if max(values) > 0 else 1
                text_item.setPos(i - 0.7, val + offset)
                plot.addItem(text_item)
        plot.setMouseEnabled(x=False, y=False)
        plot.hideButtons()
        plot.setAntialiasing(True)
        layout.addWidget(plot)
        return pivot_df

    def Line_chart(self, widget, title = "", view_by = None , show_by = None):
        old_layout = widget.layout()
        if old_layout is not None:
            while old_layout.count():
                item = old_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
        else:
            new_layout = QtWidgets.QVBoxLayout()
            new_layout.setContentsMargins(0, 0, 0, 0)
            new_layout.setSpacing(0)
            widget.setLayout(new_layout)
        
        def on_fetch_more():
            try:
                if show_by == "Machine Code":
                    filter_scripts = " Machine_Code = :object_name"
                if show_by == "Line Name":
                    filter_scripts = " Line_Name = :object_name"
                if show_by == "Error Code":
                    filter_scripts = " Error_Code = :object_name"
                if view_by == "day":
                    data = self.database.query(sql=f'''SELECT Date , SUM(Total_Loss)
                                                        FROM downtime_report
                                                        WHERE {filter_scripts}
                                                        GROUP BY Date
                                                        ORDER BY Date ASC;''',
                                                    params={"object_name": self.catagory})
                    record_df = pd.DataFrame(data, columns=["Date", "Total_Loss"])
                    return record_df
                else:
                    data = self.database.query(sql=f'''SELECT Working_Month , YEAR(Date), SUM(Total_Loss)
                                                        FROM downtime_report
                                                        WHERE {filter_scripts}
                                                        GROUP BY Working_Month, YEAR(Date)
                                                        ORDER BY YEAR(Date) ASC, Working_Month ASC;''',
                                                    params={"object_name": self.catagory})
                    record_df = pd.DataFrame(data, columns=["Working_Month", "Year", "Total_Loss"])
                    return record_df
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to fetch more data: {e}")
        
        def update_line_chart(widget, record_df):
            if view_by == "day":
                labels = record_df["Date"].apply(lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) else "").tolist()
                self.button_style_call_back((self.ui.ByDate_btn, self.ui.ByMonth_btn))
                y = record_df["Total_Loss"].to_numpy().astype(float)
            else:
                month_df = (
                    record_df
                    .groupby(["Year", "Working_Month"], as_index=False)["Total_Loss"]
                    .sum()
                )

                month_df["MonthDate"] = pd.to_datetime(
                    month_df["Year"].astype(str)
                    + "-"
                    + month_df["Working_Month"].astype(str)
                    + "-01"
                )
                start = month_df["MonthDate"].min()
                last = month_df["MonthDate"].max()
                end = pd.Timestamp(year=last.year, month=12, day=1)
                full = pd.DataFrame({
                    "MonthDate": pd.date_range(start, end, freq="MS")
                })
                month_df = full.merge(
                    month_df[["MonthDate", "Total_Loss"]],
                    on="MonthDate",
                    how="left"
                )
                labels = month_df["MonthDate"].dt.strftime("%m/%Y").tolist()
                y = month_df["Total_Loss"].to_numpy(dtype=float)
                self.button_style_call_back((self.ui.ByMonth_btn, self.ui.ByDate_btn))
                
            x = np.arange(len(labels))
            layout = widget.layout()
            plot = pg.PlotWidget()
            plot.setAntialiasing(True)
            plot.plot(
                x=x,
                y=y,
                pen=pg.mkPen(color='#09D4F8', width=2),
                symbol='o',
                symbolBrush='#09D4F8',
                symbolSize=4
            )
            VISIBLE = 30 if view_by == "day" else 12
            STEP_LABEL = 7 if len(labels) > 30 else 1
            ticks = [(i, labels[i]) for i in range(0, len(labels), STEP_LABEL)]
            if len(labels) > 0 and (len(labels) - 1) % STEP_LABEL != 0:
                ticks.append((len(labels) - 1, labels[-1]))
            chart_font = QtGui.QFont("Comic Sans MS", 8)
            chart_font.setStyleStrategy(QtGui.QFont.PreferAntialias)
            chart_font.setHintingPreference(QtGui.QFont.PreferFullHinting)
            chart_font.setBold(True)
            plot.getAxis("bottom").setTextPen(pg.mkPen(color='#474747'))
            plot.getAxis("bottom").setTickFont(chart_font)
            plot.getAxis("bottom").setTicks([ticks])
            plot.setXRange(0, min(VISIBLE-1, len(x)-1), padding=0)
            plot.setMouseEnabled(x=True, y=False)
            plot.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
            plot.hideButtons()
            plot.setBackground(None)
            y_axis = plot.getAxis("left")
            y_axis.setLabel("Time (minutes)")
            y_axis.setStyle(tickTextOffset=10, tickLength=0)
            y_axis.setPen(None)
            y_axis.setTextPen(pg.mkPen(color='#474747'))
            y_axis.setTickFont(chart_font)
            plot.showGrid(x=False, y=False)
            mask = ~np.isnan(y)
            x_fill = x[mask]
            y_fill = y[mask]
            curve = pg.PlotCurveItem(x_fill, y_fill, pen=pg.mkPen(color = '#52B7FB', width=1.5))
            fill = pg.FillBetweenItem(
                curve,
                pg.PlotCurveItem(x_fill, np.zeros_like(y_fill)),
                brush=pg.mkBrush(116, 185, 232, 80)
            )
            main_vb = plot.getViewBox()
            main_vb.wheelEvent = lambda event , axis = 0: None
            scroll = QtWidgets.QScrollBar(QtCore.Qt.Horizontal)
            scroll.setRange(
                0,
                max(0, len(x)-VISIBLE + 2)
            )
            def wheel(ev, axis=None):
                nonlocal VISIBLE
                modifiers = ev.modifiers()
                if modifiers == QtCore.Qt.ControlModifier:
                    if ev.delta() > 0:
                        VISIBLE = max(5, VISIBLE - 2)
                    else:
                        VISIBLE = min(len(x), VISIBLE + 2)
                    on_scroll(scroll.value())
                else:
                    step = -1 if ev.delta() > 0 else 1
                    value = scroll.value() + step
                    value = max(scroll.minimum(), min(scroll.maximum(), value))
                    scroll.setValue(value)
                ev.accept()

            main_vb.wheelEvent = wheel

            def on_scroll(value):
                plot.setXRange(value, value+VISIBLE-1, padding=0)

            plot.addItem(fill)
            scroll.valueChanged.connect(on_scroll)
            layout.addWidget(plot)
            layout.addWidget(scroll)
            scroll.setValue(scroll.maximum() if view_by == "day" else last.month)

        self.DT_Detail_worker = WorkerThread(on_fetch_more)
        self.DT_Detail_worker.finished.connect(lambda res: update_line_chart(widget, res))
        self.DT_Detail_worker.start()

    def time_to_minute(self,time_str):
            if not time_str or not isinstance(time_str, str):
                return time_str
            try:
                parts = list(map(int, time_str.split(':')))
                if len(parts) == 3:
                    return parts[0] * 60 + parts[1] + parts[2]/60
                elif len(parts) == 2:
                    return parts[0] * 60 + parts[1]
            except ValueError:
                return time_str
            
    def Generate_key_insights_en(self,data , root_cause = None, top_trend = None ):
        insights = []          
        final_conclusion = ""     
        if data.get("group_col") in ("Machine Code", "Line Name"):
            mttr_target = self.time_to_minute(data.get("mttr_target", 0))
            mttr = self.time_to_minute(data.get("mttr", 0))
            mttr_time = self.change_time_format_call_back(time_value = mttr,input_unit = "m", output_unit = True)
            mtbf_target = self.time_to_minute(data.get("mtbf_target", 0))
            mtbf = self.time_to_minute(data.get("mtbf", 0))
            mtbf_time = self.change_time_format_call_back(time_value = mtbf,input_unit = "m", output_unit = True)
            insights.append(f"<p style='margin: 0px; margin-top: 8px;'>• <b>Overview:</b></p>")
            
            style_sub_overview = "style='margin: 0px; margin-left: 25px;'"
            
            if mttr <= mttr_target:
                mttr_text = f"<p {style_sub_overview}>• <b>MTTR:</b> <span style='color: #10b981; font-weight: bold;'> Reached target ({mttr_time})</span></p>"
            else:
                mttr_text = f"<p {style_sub_overview}>• <b>MTTR:</b> <span style='color:red; font-weight: bold;'> Missed target ({mttr_time})</span></p>"
            insights.append(mttr_text)
            
            if mtbf >= mtbf_target:
                mtbf_text = f"<p {style_sub_overview}>• <b>MTBF:</b> <span style='color: #10b981; font-weight: bold;'> Reached target ({mtbf_time})</span></p>"
            else:
                mtbf_text = f"<p {style_sub_overview}>• <b>MTBF:</b> <span style='color:red; font-weight: bold;'> Missed target ({mtbf_time})</span></p>"
            insights.append(mtbf_text)
            
            insights.append(f"<p style='margin: 0px; margin-top: 14px;'>• <b>Reference:</b></p>")
            style_sub_ref = "style='margin: 0px; margin-left: 25px; text-align: justify;'"
            
            is_mttr_good = mttr <= mttr_target 
            is_mtbf_good = mtbf >= mtbf_target 
            
            if is_mttr_good and not is_mtbf_good:
                conclusion = (
                    f"<p {style_sub_ref}>• <b>Maintenance Efficiency:</b> High disruption frequency detected. A short MTBF (<b>{mtbf_time}</b>) "
                    f"indicates frequent equipment instability; however, the technical team demonstrated rapid responsiveness "
                    f"with a swift MTTR of <b>{mttr_time}</b>.</p>"
                )
                final_conclusion = (
                    f" Need review."
                )
            elif not is_mttr_good and is_mtbf_good:
                conclusion = (
                    f"<p {style_sub_ref}>• <b>Maintenance Efficiency:</b> Robust asset reliability but weak maintenance recovery. While the asset "
                    f"runs stably with an excellent MTBF of <b>{mtbf_time}</b>, the technical team's response is lagging, "
                    f"leading to a prolonged MTTR of <b>{mttr_time}</b>.</p>"
                )
                final_conclusion = (
                    f" Need review."
                )
            elif is_mttr_good and is_mtbf_good:
                conclusion = (
                    f"<p {style_sub_ref}>• <b>Maintenance Efficiency:</b> Outstanding operational and maintenance performance. The asset demonstrates "
                    f"high reliability with a strong MTBF of <b>{mtbf_time}</b>, while the maintenance crew maintains peak efficiency "
                    f"with a minimal MTTR of <b>{mttr_time}</b>.</p>"
                )
                final_conclusion = (
                    f" No need for immediate action."
                )
            else:
                conclusion = (
                    f"<p {style_sub_ref}>• <b>Maintenance Efficiency:</b> <span style='color:#ef4444;'><b>CRITICAL O&M RISK!</b></span> "
                    f"The line suffers from poor reliability with a failing MTBF of <b>{mtbf_time}</b>, compounded by "
                    f"excessive repair delays showing a critical MTTR of <b>{mttr_time}</b>.</p>"
                )
                final_conclusion = (
                    f"<span style='font-weight: bold; color: #ef4444;'>Need for immediate action.</span>"
                )
            insights.append(conclusion)
            
        def analyze_top_drivers(data, threshold=0.1, type_analyze="root_cause"):
            if data is None or data.empty:
                return "No data available."
                
            total_value = data["Total Loss Time"].sum()
            data = list(zip(data.iloc[:, 0], data.iloc[:, 1]))
            max_value = data[0][1]
            top_cluster = []
            for name, value in data:
                if (max_value - value) / max_value <= threshold:
                    top_cluster.append((name, value))
                else:
                    break
                    
            cluster_len = len(top_cluster)
            cluster_names = [item[0] for item in top_cluster]
            cluster_total_val = sum(item[1] for item in top_cluster)
            cluster_share = round((cluster_total_val / total_value) * 100, 1)
            
            if type_analyze == "root_cause":
                if cluster_len == 1:
                    return f"Error code <b>{cluster_names[0]}</b> is the dominant root cause, accounting for <b>{cluster_share}%</b> of all occurrences. Targeted engineering focus is required here."
                elif cluster_len >= 3:
                    formatted_names = ", ".join(cluster_names[:-1]) + f", and {cluster_names[-1]}"
                    return (f"Downtime is <b>evenly distributed</b> across multiple assets ({formatted_names}), "
                            f"collectively driving <b>{cluster_share}%</b> of the total losses. "
                            f"This indicates a <b>systemic issue</b> rather than isolated machine failures.")                     
                else:
                    return f"Primary losses are co-driven by <b>{cluster_names[0]}</b> and <b>{cluster_names[1]}</b>, contributing <b>{cluster_share}%</b> of total impact."
            elif type_analyze == "top_trend":
                if cluster_len == 1:
                    return f"Machine <b>{cluster_names[0]}</b> is the main contributor to this trend, accounting for <b>{cluster_share}%</b> of the total issues."
                elif cluster_len >= 3:
                    formatted_names = ", ".join(cluster_names[:-1]) + f", and {cluster_names[-1]}"
                    return f"The main issues are driven by machines <b>{formatted_names}</b>, collectively accounting for <b>{cluster_share}%</b> of the total."
                else:
                    return f"Machines <b>{cluster_names[0]}</b> and <b>{cluster_names[1]}</b> are the primary sources of the problem, accounting for <b>{cluster_share}%</b> of the total."
        if root_cause is not None:       
            analysis_result = analyze_top_drivers(data=root_cause, threshold=0.1, type_analyze="root_cause")
            insights.append(f"<p style='margin: 0px; margin-left: 25px; margin-top: 4px; text-align: justify;'>• <b>Root Cause Analysis:</b> {analysis_result}</p>")
            
        if top_trend is not None:
            analysis_result = analyze_top_drivers(data=top_trend, threshold=0.1, type_analyze="top_trend")
            insights.append(f"<p style='margin: 0px; margin-left: 25px; margin-top: 4px; text-align: justify;'>• <b>Top Trend Analysis:</b> {analysis_result}</p>")
        if final_conclusion:
            insights.append(f"<p style='margin: 0px; margin-top: 14px;'>• <b>Final Recommendation:</b> {final_conclusion}</p>")
        return "".join(insights)

    def action_text_drop_event(self, event):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0].toLocalFile()
            extracted_text = url.split("/")[-1]
            file_url = QtCore.QUrl.fromLocalFile(url).toString()
            cursor = self.ui.action_text.textCursor()
            cursor.movePosition(QtGui.QTextCursor.End)
            self.ui.action_text.setTextCursor(cursor)
            self.ui.action_text.append(
                f'<a href="{file_url}" style="color:#007acc; font-size:10px; text-decoration:underline;">'
                f'{extracted_text}</a> '
            )
            event.acceptProposedAction()

    def action_text_mouse_press_event(self, event):
        anchor = self.ui.action_text.anchorAt(event.pos())
        if anchor:
            url = QtCore.QUrl(anchor)
            path = url.toLocalFile() if url.isLocalFile() else anchor
            path = os.path.normpath(path)  
            if os.path.exists(path):
                try:
                    os.startfile(path)
                except OSError as e:
                    QtWidgets.QMessageBox.warning(
                        self, "Lỗi", f"Không mở được file:\n{e}")
            else:
                QtWidgets.QMessageBox.warning(
                    self, "Không tìm thấy file",
                    f"File không tồn tại hoặc không truy cập được:\n{path}")
            return
        QtWidgets.QTextEdit.mousePressEvent(self.ui.action_text, event)
    
    def mouseMoveEvent(self, event):
        if self.ui.action_text.anchorAt(event.pos()):
            self.ui.action_text.viewport().setCursor(QtCore.Qt.PointingHandCursor)
        else:
            self.ui.action_text.viewport().setCursor(QtCore.Qt.IBeamCursor)
        super().mouseMoveEvent(event)

class MetricMiniBar(QtWidgets.QWidget):
    def __init__(self, label, value, parent=None):
        super().__init__(parent)
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(0, 2, 0, 2)
        layout.setSpacing(10)
        FONT_FAMILY = "Segoe UI"
        COLOR_PRIMARY = "#0F4C81"   # Deep classic blue
        COLOR_SUCCESS = "#2ECC71"   # Emerald Green (OEE >= 85%)
        COLOR_WARNING = "#F1C40F"   # Sun Yellow (OEE 65% - 85%)
        COLOR_DANGER = "#E74C3C"    # Alizarin Red (OEE < 65%)
        COLOR_BG = "#F3F4F6"        # Neutral light background
        COLOR_CARD_BG = "#FFFFFF"   # White cards
        COLOR_TEXT_MAIN = "#1F2937" # Dark graphite
        COLOR_TEXT_MUTED = "#6B7280"
        lbl = QtWidgets.QLabel(label, self)
        lbl.setFont(QtGui.QFont(FONT_FAMILY, 9, QtGui.QFont.Bold))
        lbl.setStyleSheet(f"color: {COLOR_TEXT_MUTED}; border: none;")
        lbl.setFixedWidth(15)
        self.value = round(value*100, 1)
        self.bar = QtWidgets.QProgressBar(self)
        self.bar.setValue(int(self.value))
        self.bar.setTextVisible(False)
        
        # color = COLOR_SUCCESS if self.value >= 85 else (COLOR_WARNING if self.value >= 65 else COLOR_DANGER)
        color = COLOR_SUCCESS
        self.bar.setStyleSheet(f"""
            QProgressBar {{
                background-color: #E5E7EB;
                border-radius: 4px;
                border: none;
            }}
            QProgressBar::chunk {{
                background-color: {color};
                border-radius: 4px;
            }}
        """)
        
        val_lbl = QtWidgets.QLabel(f"{self.value:.1f}%", self)
        val_lbl.setFont(QtGui.QFont(FONT_FAMILY, 10, QtGui.QFont.Bold))
        val_lbl.setStyleSheet(f"color: {COLOR_TEXT_MAIN}; border: none;")
        val_lbl.setFixedWidth(45)
        layout.addWidget(lbl)
        layout.addWidget(self.bar)
        layout.addWidget(val_lbl)

class Mini_Card(QtWidgets.QFrame):
    doubleClicked = QtCore.pyqtSignal()
    def __init__(self, data=None, format_time=None, parent=None):
        super().__init__(parent)
        self.ui = Ui_OEE_mini_card()
        self.ui.setupUi(self)
        self.format_time = format_time
        self.line_name = data["line_name"]
        self.model_name = data["model_name"]
        self.process_name = data["process"]
        self.oee_target = data["oee_target"]
        self.oee_value = data["oee_value"]*100
        self.a_value = data["availability_value"]
        self.p_value = data["performance_value"]
        self.q_value = data["quality_value"]
        self.mttr_value = data["mttr_value"]
        self.mttr_tarrget = data["mttr_target"]
        self.mtbf_value = data["mtbf_value"]
        self.mtbf_target = data["mtbf_target"]
        self.max_mttr_value = data["max_mttr_value"]
        self.max_mtbf_value = data["max_mtbf_value"]
        self.information()
        self.donut_chart(percent_value=self.oee_value)
        self.ui.metrics_layout.addWidget(MetricMiniBar("A", self.a_value))
        self.ui.metrics_layout.addWidget(MetricMiniBar("P", self.p_value))
        self.ui.metrics_layout.addWidget(MetricMiniBar("Q", self.q_value))
        mttr_chart = Bullet_Status_Bar(value=self.mttr_value, max_value=self.max_mttr_value, target_value=self.mttr_tarrget, previous_value=0, html_doc=True, label="MTTR", format_time=self.format_time, height_ratio = 0.35)
        self.ui.mttr_chart_layout.addWidget(mttr_chart)
        mtbf_chart = Bullet_Status_Bar(value=self.mtbf_value,max_value=self.max_mtbf_value ,target_value=self.mtbf_target, previous_value=0, html_doc=True, label="MTBF", format_time=self.format_time, height_ratio = 0.35)
        self.ui.mtbf_chart_layout.addWidget(mtbf_chart)
        self.setCursor(QtCore.Qt.PointingHandCursor)

    def information(self):
        self.ui.line_lbl.setText(self.line_name)
        self.ui.model_lbl.setText(self.model_name)
        self.ui.process_lbl.setText(self.process_name)

    def donut_chart(self, percent_value):
        old_layout = self.ui.OEE_chart.layout()
        if old_layout is not None:
            while old_layout.count():
                item = old_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
        else:
            new_layout = QtWidgets.QVBoxLayout()
            new_layout.setContentsMargins(0, 0, 0, 0)
            new_layout.setSpacing(0)
            self.ui.OEE_chart.setLayout(new_layout)
        layout = self.ui.OEE_chart.layout()
        chart = DonutChart(value=percent_value,target_value = self.oee_target, parameter_name="%OEE", scale = 0.9, has_Lengend=False)
        layout.addWidget(chart)

    def mouseDoubleClickEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.doubleClicked.emit()
            event.accept()
        else:
            super().mouseDoubleClickEvent(event)
        

    
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

    QtWidgets.QApplication.setAttribute(
        QtCore.Qt.AA_EnableHighDpiScaling, True)
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
    try:
        main()
    except Exception as e:
        print(f"An error occurred: {e}")
        with open("error_log.txt", "w", encoding="utf-8") as f:
            f.write(f"App bị crash do lỗi:\n{str(e)}\n\n")
            f.write("Chi tiết Traceback:\n")
            traceback.print_exc(file=f)

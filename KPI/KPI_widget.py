# -*- coding: utf-8 -*-
"""
TỔNG QUAN KPIs - KPI Dashboard (PyQt5)
Recreates the TOKIN factory KPI dashboard with the "fan / petal" card
layout: four cards curve around a central factory badge (each card has a
concave notch cut from its inner corner), a Maintenance donut, an OEE
half-donut gauge with sub-metric bars, and MTTR / MTBF speedometer dials
whose readouts sit inside the dial.

Run:
    python kpi_dashboard.py
"""

import sys
import math
from datetime import datetime

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QGridLayout, QFrame, QGraphicsDropShadowEffect, QSizePolicy
)
from PyQt5.QtCore import Qt, QRectF, QPointF
from PyQt5.QtGui import (
    QPainter, QColor, QPen, QBrush, QFont, QPainterPath, QRadialGradient, QPixmap
)


# ----------------------------------------------------------------------
#  Palette
# ----------------------------------------------------------------------
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
TARGET_COLOR = "#6bffff"



def make_card_path(w, h, r, nr, notch):
    iTL = nr if notch == "tl" else r
    iTR = nr if notch == "tr" else r
    iBR = nr if notch == "br" else r
    iBL = nr if notch == "bl" else r

    p = QPainterPath()
    p.moveTo(iTL, 0)
    p.lineTo(w - iTR, 0)

    # top-right corner
    if notch == "tr":
        p.arcTo(w - nr, -nr, 2 * nr, 2 * nr, 180, 90) 
    else:
        p.arcTo(w - 2 * r, 0, 2 * r, 2 * r, 90, -90)      
    p.lineTo(w, h - iBR)

    # bottom-right corner
    if notch == "br":
        p.arcTo(w - nr, h - nr, 2 * nr, 2 * nr, 90, 90)   
    else:
        p.arcTo(w - 2 * r, h - 2 * r, 2 * r, 2 * r, 0, -90)
    p.lineTo(iBL, h)

    # bottom-left corner
    if notch == "bl":
        p.arcTo(-nr, h - nr, 2 * nr, 2 * nr, 0, 90)    
    else:
        p.arcTo(0, h - 2 * r, 2 * r, 2 * r, 270, -90)
    p.lineTo(0, iTL)

    # top-left corner
    if notch == "tl":
        p.arcTo(-nr, -nr, 2 * nr, 2 * nr, 270, 90)   
    else:
        p.arcTo(0, 0, 2 * r, 2 * r, 180, -90)
    p.closeSubpath()
    return p

class ArcGauge(QWidget):
    def __init__(self, value, max_value, target ,color, ticks=None,
                 start_angle=225.0, span_angle=-270.0,
                 minor_ticks=48, center_lines=None, ratio=1.3, parent=None):
        super().__init__(parent)
        self.value = value
        self.max_value = max_value
        self.target = target
        self.color = QColor(color)
        self.ticks = ticks or []
        self.start_angle = start_angle
        self.span_angle = span_angle
        self.minor_ticks = minor_ticks
        self.center_lines = center_lines or []   # list of dicts
        self.setMinimumSize(200, 170)
        self.ratio = ratio

    def setValue(self, value, max_value=None):
        self.value = value
        if max_value is not None:
            self.max_value = max_value
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        w, h = self.width(), self.height()
        side = min(w, h * self.ratio)
        thickness = max(9, side * 0.11)
        radius = side / 2 - thickness
        cx = w / 2
        cy = h * 0.52 if self.span_angle <= -260 else h*0.78
        rect = QRectF(cx - radius, cy - radius, radius * 2, radius * 2)

        # --- background track ---
        pen = QPen(QColor(TRACK), thickness, Qt.SolidLine, Qt.RoundCap)
        p.setPen(pen)
        p.drawArc(rect, int(self.start_angle * 16), int(self.span_angle * 16))

        def draw_arc(value, color, pen_thickness):
            frac = max(0.0, min(1.0, value / self.max_value)) if self.max_value else 0
            if frac > 0:
                pen.setWidthF(pen_thickness)
                pen.setColor(QColor(color))
                p.setPen(pen)
                p.drawArc(rect, int(self.start_angle * 16),
                          int(self.span_angle * frac * 16))
        if self.target > self.value:
            draw_arc(self.target, TARGET_COLOR, thickness * 0.5)
            draw_arc(self.value, self.color, thickness)
        else:
            draw_arc(self.value, self.color, thickness)
            draw_arc(self.target, TARGET_COLOR, thickness * 0.5)
        # --- fine dotted minor ticks (just outside the ring) ---
        if self.minor_ticks:
            p.setPen(QPen(QColor(TEXT_GRAY), 1))
            r_in = radius + thickness * 0.60
            r_out = radius + thickness * 0.95
            for i in range(self.minor_ticks + 1):
                f = i / self.minor_ticks
                ang = math.radians(self.start_angle + self.span_angle * f)
                x1 = cx + r_in * math.cos(ang)
                y1 = cy - r_in * math.sin(ang)
                x2 = cx + r_out * math.cos(ang)
                y2 = cy - r_out * math.sin(ang)
                p.drawLine(QPointF(x1, y1), QPointF(x2, y2))

        # --- major tick labels ---
        if self.ticks:
            p.setPen(QColor(TEXT_GRAY))
            p.setFont(QFont("Segoe UI", max(7, int(radius * 0.11))))
            n = len(self.ticks) - 1
            r_lbl = radius + thickness * 1.55
            for i, t in enumerate(self.ticks):
                f = i / n if n else 0
                ang = math.radians(self.start_angle + self.span_angle * f)
                tx = cx + r_lbl * math.cos(ang)
                ty = cy - r_lbl * math.sin(ang)
                p.drawText(QRectF(tx - 16, ty - 9, 32, 18),
                           Qt.AlignCenter, str(t))

        # --- centred text lines (title + value) ---
        if self.center_lines:
            total = sum(l.get("size", 14) + 6 for l in self.center_lines)
            y = cy - total / 2
            for line in self.center_lines:
                size = line.get("size", 14)
                p.setPen(QColor(line.get("color", TEXT_DARK)))
                f = QFont("Segoe UI", size,
                          QFont.Bold if line.get("bold") else QFont.Normal)
                p.setFont(f)
                p.drawText(QRectF(cx - radius, y, radius * 2, size + 8),
                           Qt.AlignCenter, line["text"])
                y += size + 6

class Arc_Health_Gauge(QWidget):
    def __init__(self, value, max_value, target ,color_desc = True, ticks=None,
                 start_angle=225.0, span_angle=-270.0,
                 minor_ticks=48, center_lines=None, ratio=1.3, parent=None):
        super().__init__(parent)
        self.value = value
        self.max_value = max_value
        self.target = target
        if color_desc:
            self.LIGHT_COLOR = [
                            "#28851F", "#55A028", "#63B72F", "#8BCA35", 
                            "#FDD03C", "#F8B334", "#F3962D", "#EF7A26",
                            "#E95B20", "#E33F1A", "#DD2616", "#D61314"
                        ]
            self.DARK_COLOR = [
                            "#14420F", "#2A5014", "#315B17", "#46661A",
                            "#80691E", "#7D591A", "#794B16", "#774014",
                            "#732E12", "#701F0D", "#6E120B", "#6B090A"
                        ]
        else:
            self.LIGHT_COLOR = [
                            "#D61314", "#DD2616", "#E33F1A", "#E95B20",
                            "#EF7A26", "#F3962D", "#F8B334", "#FDD03C",
                            "#8BCA35",  "#63B72F", "#55A028", "#28851F"
                        ]
            self.DARK_COLOR = [
                            "#6B090A", "#6E120B", "#701F0D", "#732E12", 
                            "#774014", "#794B16", "#7D591A", "#80691E",
                            "#46661A", "#315B17", "#2A5014", "#14420F"
                        ]
        self.ticks = ticks or []
        self.start_angle = start_angle
        self.span_angle = span_angle
        self.minor_ticks = minor_ticks
        self.center_lines = center_lines or []
        self.setMinimumSize(200, 170)
        self.ratio = ratio

    def setValue(self, value, max_value=None):
        self.value = value
        if max_value is not None:
            self.max_value = max_value
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        w, h = self.width(), self.height()
        side = min(w, h * self.ratio)
        thickness = max(9, side * 0.01)
        radius = side / 2 - thickness
        cx = w / 2
        cy = h * 0.52 if self.span_angle <= -260 else h*0.78
        rect = QRectF(cx - radius, cy - radius, radius * 2, radius * 2)

        # --- background track ---
        self.split_part_arc = int(self.span_angle / len(self.LIGHT_COLOR))
        frac = max(0.0, min(1.0, self.value / self.max_value)) if self.max_value else 0
        color = None
        for i in range(len(self.LIGHT_COLOR)):
            if self.value > self.max_value * i / len(self.LIGHT_COLOR):
                color = self.LIGHT_COLOR[i]
            pen = QPen(QColor(self.LIGHT_COLOR[i]), thickness, Qt.SolidLine, Qt.FlatCap)
            p.setPen(pen)
            p.drawArc(rect, int((self.start_angle + self.split_part_arc * i) * 16), int(self.split_part_arc * 16))

        offset = thickness * 2
        rect_inner = rect.adjusted(offset, offset, -offset, -offset)
        pen = QPen(QColor(TRACK), thickness*2, Qt.SolidLine, Qt.FlatCap)
        p.setPen(pen)
        p.drawArc(rect_inner, int(self.start_angle * 16), int(self.span_angle * 16))

        pen.setColor(QColor(color))
        p.setPen(pen)
        p.drawArc(rect_inner, int(self.start_angle * 16), int(self.span_angle * frac * 16))

        if self.minor_ticks:
            p.setPen(QPen(QColor(TEXT_GRAY), 1))
            r_in = radius + thickness * 0.60
            r_out = radius + thickness * 0.95
            for i in range(self.minor_ticks + 1):
                f = i / self.minor_ticks
                ang = math.radians(self.start_angle + self.span_angle * f)
                x1 = cx + r_in * math.cos(ang)
                y1 = cy - r_in * math.sin(ang)
                x2 = cx + r_out * math.cos(ang)
                y2 = cy - r_out * math.sin(ang)
                p.drawLine(QPointF(x1, y1), QPointF(x2, y2))

        # --- major tick labels ---
        if self.ticks:
            p.setPen(QColor(TEXT_GRAY))
            p.setFont(QFont("Segoe UI", max(7, int(radius * 0.11))))
            n = len(self.ticks) - 1
            r_lbl = radius + thickness * 2.0
            for i, t in enumerate(self.ticks):
                f = i / n if n else 0
                ang = math.radians(self.start_angle + self.span_angle * f)
                tx = cx + r_lbl * math.cos(ang)
                ty = cy - r_lbl * math.sin(ang)
                p.drawText(QRectF(tx - 16, ty - 9, 32, 18),
                           Qt.AlignCenter, str(t))

        # --- centred text lines (title + value) ---
        if self.center_lines:
            total = sum(l.get("size", 14) + 6 for l in self.center_lines)
            y = cy - total / 2
            for line in self.center_lines:
                size = line.get("size", 14)
                p.setPen(QColor(line.get("color", TEXT_DARK)))
                f = QFont("Segoe UI", size,
                          QFont.Bold if line.get("bold") else QFont.Normal)
                p.setFont(f)
                p.drawText(QRectF(cx - radius, y, radius * 2, size + 8),
                           Qt.AlignCenter, line["text"])
                y += size + 6

class PetalCard(QFrame):
    def __init__(self, notch, notch_radius=96, radius=22, shadow_color=QColor(30, 60, 90, 45), parent=None):
        super().__init__(parent)
        self.notch = notch
        self.nr = notch_radius
        self.radius = radius
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(30)
        shadow.setOffset(0, 5)
        shadow.setColor(QColor(shadow_color))
        self.setGraphicsEffect(shadow)
        self.setSizePolicy(QSizePolicy.Expanding,
                           QSizePolicy.Expanding)
        # # keep content clear of the notched corner
        # m = 26
        # pad = self.nr - 40
        # left = pad if notch in ("tl", "bl") else m
        # right = pad if notch in ("tr", "br") else m
        # top = pad if notch in ("tl", "tr") else m
        # bottom = pad if notch in ("bl", "br") else m
        self.vbox = QVBoxLayout(self)
        self.vbox.setContentsMargins(10, 10, 10, 10)
        self.vbox.setSpacing(8)

    def set_shadow_color(self, color):
        shadow = self.graphicsEffect()
        if isinstance(shadow, QGraphicsDropShadowEffect):
            shadow.setColor(QColor(color))
            self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        path = make_card_path(self.width() - 2, self.height() - 2,
                              self.radius, self.nr, self.notch)
        path.translate(1, 1)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(CARD_BG))
        p.drawPath(path)

    def add_header(self, icon_link, title, align_right=False):
        row = QHBoxLayout()
        row.setSpacing(10)
        row.setContentsMargins(0, 0, 0, 0)
        lbl = QLabel(title)
        lbl.setStyleSheet(f"color:{TEXT_DARK}; font-size:18px; font-weight:700;")
        font = QFont("Segoe UI", 18, QFont.Bold)
        lbl.setFont(font)
        icon = QLabel()
        pixmap = QPixmap(icon_link).scaled(
            32, 32,
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation
        )
        icon.setPixmap(pixmap)
        if align_right:
            row.addStretch(1)
            row.addWidget(lbl)
            row.addWidget(icon)
        else:
            row.addWidget(icon)
            row.addWidget(lbl)
            row.addStretch(1)
        self.vbox.addLayout(row)

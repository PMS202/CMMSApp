
from PyQt5 import QtCore, QtWidgets, QtGui
from dataclasses import dataclass
from typing import Optional
import os
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

@dataclass
class ProcessData:

    sequence_no: int
    process_name: str

    machine_count: int

    actual_ct_sec: Optional[float]
    capacity_pcs_h: Optional[float]

    utilization_pct: Optional[float]

    status: str
    is_bottleneck: bool


class ProcessCard(QtWidgets.QGroupBox):

    clicked = QtCore.pyqtSignal(object)

    def __init__(self, process, parent=None):
        super().__init__(parent)

        self.process = process
        self.setFixedSize(140, 175)
        self.setTitle("")
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.build_ui()
        self.update_status()
        
    def build_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.nametag_frame = QtWidgets.QFrame(self)
        self.nametag_frame.setObjectName(u"nametag_frame")
        self.nametag_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.nametag_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.verticalLayout = QtWidgets.QVBoxLayout(self.nametag_frame)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.verticalLayout.setSpacing(6)
        self.verticalLayout.setContentsMargins(7, 7, 7, 7)
        self.sequence_no_label = QtWidgets.QLabel(self.nametag_frame)
        self.sequence_no_label.setObjectName("sequence_no_label")
        self.sequence_no_label.setText(f"P{self.process.sequence_no:02d}")
        self.sequence_no_label.setAlignment(QtCore.Qt.AlignCenter)
        self.verticalLayout.addWidget(self.sequence_no_label)
        self.name_label = QtWidgets.QLabel(self.nametag_frame)
        self.name_label.setObjectName("name_label")
        self.name_label.setText(f"{self.process.process_name}")
        self.name_label.setAlignment(QtCore.Qt.AlignCenter)
        self.name_label.setWordWrap(True)
        self.verticalLayout.addWidget(self.name_label)
        layout.addWidget(self.nametag_frame,3)
        self.process_infor_frame = QtWidgets.QFrame(self)
        self.process_infor_frame.setObjectName(u"process_infor_frame")
        self.process_infor_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.process_infor_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.verticalLayout_1 = QtWidgets.QVBoxLayout(self.process_infor_frame)
        self.verticalLayout_1.setObjectName(u"verticalLayout_2")
        self.verticalLayout_1.setSpacing(6)
        self.verticalLayout_1.setContentsMargins(7, 7, 7, 7)
        self.ct_label = QtWidgets.QLabel(self.process_infor_frame)
        self.ct_label.setObjectName("ct_label")
        self.verticalLayout_1.addWidget(self.ct_label)
        self.capacity_label = QtWidgets.QLabel(self.process_infor_frame)
        self.capacity_label.setObjectName("capacity_label")
        self.verticalLayout_1.addWidget(self.capacity_label)
        self.machine_label = QtWidgets.QLabel(self.process_infor_frame)
        self.machine_label.setObjectName("machine_label")
        self.verticalLayout_1.addWidget(self.machine_label)
        for label in [
            self.machine_label,
            self.ct_label,
            self.capacity_label
        ]:
            label.setAlignment(QtCore.Qt.AlignLeft)
        layout.addWidget(self.process_infor_frame,7)
        self.refresh_data()

    def refresh_data(self):

        p = self.process

        self.machine_label.setText(
            f"⚙ M: {p.machine_count} (Ref)"
        )

        self.ct_label.setText(
            f"◷ CT: {p.actual_ct_sec:.1f} s"
            if p.actual_ct_sec is not None
            else "◷ CT: N/A"
        )

        if p.capacity_pcs_h is None:
            self.capacity_label.setText("▤ Cap: N/A")
        else:
            if p.status == "critical": 
                self.capacity_label.setText(f"▤ Cap: {p.capacity_pcs_h:.1f} pcs/h \n( Minimum! )")
                self.capacity_label.setWordWrap(True)
            else:
                self.capacity_label.setText(f"▤ Cap: {p.capacity_pcs_h:.1f} pcs/h")

    def update_status(self):

        colors = {
            "normal": "#22A06B",
            "warning": "#E5A323",
            "critical": "#F86167",
            "unknown": "#8B95A5"
        }

        nametag_colors = {
            "normal": "#E5F7ED",
            "warning": "#F9F4E6",
            "critical": "#FCE0E5",
            "unknown": "#CFCFCF"
        }
        process_infor_colors = {
            "normal": "#F6FAFC",
            "warning": "#FAFBFB",
            "critical": "#FCF0F5",
            "unknown": "#F0F0F0"
        }
        
        color = colors.get(self.process.status, "#8B95A5")
        nametag_color = nametag_colors.get(self.process.status,"#8B95A5")
        process_infor_color = process_infor_colors.get(self.process.status,"#8B95A5")
        
        self.setStyleSheet(f"""
            ProcessCard {{
                background-color: white;
                border: 1.5px solid {color};
                border-radius: 9px;
            }}
            QLabel {{
                color: #263445;
                border: none;
                background: transparent;
                font-size: 9pt;
            }}
        """)
        self.nametag_frame.setStyleSheet(f"""
                                    QFrame {{
                                        background: {nametag_color};
                                        border: transparent;
                                        border-top-left-radius: 8px;
                                        border-top-right-radius: 8px;
                                        border-bottom-left-radius: 0px;
                                        border-bottom-right-radius: 0px;
                                    }}
                    """
                    )
        self.process_infor_frame.setStyleSheet(f"""
                                    QFrame {{
                                        background: {process_infor_color};
                                        border: transparent;
                                        border-top-left-radius: 0px;
                                        border-top-right-radius: 0px;
                                        border-bottom-left-radius: 8px;
                                        border-bottom-right-radius: 8px;
                                    }}
                    """ )
        if self.process.is_bottleneck:
            self.sequence_no_label.setStyleSheet("""
                            QLabel {
                                color: #C40101;
                                font-weight: bold;
                            }
                        """)
            self.name_label.setStyleSheet("""
                QLabel {
                    color: #C40101;
                    font-weight: bold;
                }
            """)
            self.capacity_label.setStyleSheet("""
                QLabel {
                    color: #C40101;
                    font-weight: bold;
                }
            """)
            self.setStyleSheet("""
                QGroupBox {
                    background-color: #FFF1F2;

                    border: 1.5px solid #FF4555;
                    border-radius: 9px;
                }

                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top center;
                    top: -5px;
                    background-color: #E53945;
                    color: white;
                    padding: 4px 18px;
                    font-size: 18px;
                    font-weight: bold;
                }
            """)

    def mousePressEvent(self, event):

        if event.button() == QtCore.Qt.LeftButton:
            self.clicked.emit(self.process)

        super().mousePressEvent(event)


# class FlowPage(QtWidgets.QWidget):

#     process_selected = QtCore.pyqtSignal(object)

#     def __init__(self, parent=None):
#         super().__init__(parent)

#         self.cards = []

#         self.build_ui()

#     def build_ui(self):

#         layout = QtWidgets.QVBoxLayout(self)
#         layout.setContentsMargins(0, 0, 0, 0)

#         self.scroll = QtWidgets.QScrollArea()

#         self.scroll.setWidgetResizable(False)

#         self.scroll.setHorizontalScrollBarPolicy(
#             QtCore.Qt.ScrollBarAsNeeded
#         )

#         self.scroll.setVerticalScrollBarPolicy(
#             QtCore.Qt.ScrollBarAlwaysOff
#         )

#         self.scroll.setFrameShape(
#             QtWidgets.QFrame.NoFrame
#         )

#         self.content = QtWidgets.QWidget()

#         self.flow_layout = QtWidgets.QHBoxLayout(
#             self.content
#         )

#         self.flow_layout.setContentsMargins(
#             12, 15, 12, 15
#         )

#         self.flow_layout.setSpacing(8)

#         self.flow_layout.setAlignment(
#             QtCore.Qt.AlignLeft |
#             QtCore.Qt.AlignVCenter
#         )

#         self.scroll.setWidget(self.content)

#         layout.addWidget(self.scroll)

#     def set_processes(self, processes):
#         while self.flow_layout.count():
#             item = self.flow_layout.takeAt(0)
#             widget = item.widget()
#             if widget:
#                 widget.deleteLater()
#         self.cards.clear()
#         processes = sorted(
#             processes,
#             key=lambda p: p.sequence_no
#         )
#         for index, process in enumerate(processes):
#             card = ProcessCard(process)
#             card.clicked.connect(
#                 self.process_selected.emit
#             )
#             self.flow_layout.addWidget(card)
#             self.cards.append(card)
#             if index < len(processes) - 1:
#                 arrow = QtWidgets.QLabel("➜")
#                 arrow.setAlignment(
#                     QtCore.Qt.AlignCenter
#                 )
#                 arrow.setFixedWidth(25)
#                 arrow.setStyleSheet("""
#                     QLabel {
#                         color: #465466;
#                         font-size: 18px;
#                         font-weight: bold;
#                         border: none;
#                     }
#                 """)
#                 self.flow_layout.addWidget(arrow)

#         self.flow_layout.activate()
#         self.content.setMinimumSize(
#             self.flow_layout.sizeHint()
#         )
#         self.content.adjustSize()



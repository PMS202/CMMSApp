import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QProgressBar, QFrame)
from PyQt5.QtCore import Qt

class KPIStatCard(QFrame):
    def __init__(self, title, current_value, target_value, progress_percent, color_hex="#1565C0"):
        """
        Khởi tạo KPI Card
        :param title: Tiêu đề card (VD: "MAINTENANCE ACHIEVEMENT")
        :param current_value: Giá trị hiển thị lớn (VD: "100.0%")
        :param target_value: Giá trị mục tiêu (VD: "100%")
        :param progress_percent: Phần trăm thanh tiến độ (0 - 100)
        :param color_hex: Màu chủ đạo của chữ và progress bar
        """
        super().__init__()
        self.initUI(title, current_value, target_value, progress_percent, color_hex)

    def initUI(self, title, current_value, target_value, progress_percent, color_hex):
        # 1. CSS cho Card (Nền trắng, viền xám nhạt, bo góc)
        self.setStyleSheet("""
            KPIStatCard {
                background-color: #FFFFFF;
                border-radius: 10px;
                border: 1px solid #E0E0E0;
            }
        """)
        
        # Layout chính của thẻ
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)  # Khoảng trống (padding)
        main_layout.setSpacing(10)

        # 2. Label Tiêu đề
        lbl_title = QLabel(title.upper())
        lbl_title.setStyleSheet("""
            color: #757575; 
            font-weight: bold; 
            font-size: 13px; 
            font-family: 'Segoe UI', Arial, sans-serif;
            border: none;
        """)

        # 3. Layout cho Giá trị & Target
        val_layout = QHBoxLayout()
        val_layout.setSpacing(10)
        
        lbl_value = QLabel(str(current_value))
        lbl_value.setStyleSheet(f"""
            color: {color_hex}; 
            font-weight: 800; 
            font-size: 28px; 
            font-family: 'Segoe UI', Arial, sans-serif;
            border: none;
        """)

        lbl_target = QLabel(f"Target: {target_value}")
        lbl_target.setStyleSheet("""
            color: #9E9E9E; 
            font-size: 12px; 
            font-weight: 600;
            font-family: 'Segoe UI', Arial, sans-serif;
            border: none;
        """)
        lbl_target.setAlignment(Qt.AlignBottom | Qt.AlignLeft)

        val_layout.addWidget(lbl_value)
        val_layout.addWidget(lbl_target)
        val_layout.addStretch() # Đẩy các thành phần sang trái

        # 4. Progress Bar
        progress_bar = QProgressBar()
        progress_bar.setValue(progress_percent)
        progress_bar.setTextVisible(False)  # Ẩn chữ % mặc định của Qt
        progress_bar.setFixedHeight(8)      # Dáng thanh mỏng, hiện đại
        
        # CSS cho Progress Bar
        progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: none;
                border-radius: 4px;
                background-color: #EEEEEE;
            }}
            QProgressBar::chunk {{
                background-color: {color_hex};
                border-radius: 4px;
            }}
        """)

        # 5. Lắp ráp các thành phần vào layout
        main_layout.addWidget(lbl_title)
        main_layout.addLayout(val_layout)
        main_layout.addWidget(progress_bar)

        self.setLayout(main_layout)

# --- PHẦN TEST GIAO DIỆN ---
if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    # Tạo một window giả lập chứa các card
    window = QWidget()
    window.setWindowTitle("Industrial Dashboard Example")
    window.setStyleSheet("background-color: #F5F7FA;") # Nền xám nhạt cho dashboard
    layout = QHBoxLayout()
    
    # Tạo Card 1: Maintenance (Màu xanh lá)
    card_maintenance = KPIStatCard(
        title="Maintenance Achievement",
        current_value="100.0%",
        target_value="100%",
        progress_percent=100,
        color_hex="#2E7D32" # Xanh lá chuẩn doanh nghiệp
    )
    
    # Tạo Card 2: MTTR (Màu đỏ cảnh báo do vượt chỉ tiêu)
    card_mttr = KPIStatCard(
        title="Mean Time To Repair (MTTR)",
        current_value="17m 14s",
        target_value="15m",
        progress_percent=100, # Fill full báo động hoặc tính tỷ lệ tuỳ logic
        color_hex="#E53935" # Đỏ pastel
    )

    layout.addWidget(card_maintenance)
    layout.addWidget(card_mttr)
    window.setLayout(layout)
    window.resize(600, 200)
    window.show()
    
    sys.exit(app.exec_())
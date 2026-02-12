import sys
import os
import subprocess
import json
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QFrame, QTextEdit, QScrollArea,
    QSizePolicy, QMessageBox
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QIcon, QFont

# === CONFIGURATION ===
APP_TITLE = "Master Control Program"

# Program Definitions
PROGRAMS = {
    "UVNK": {
        "folder": "UVNK",                     
        "entry_point": "main_gui.py",         
        "help_file": "📄 Руководство пользователя.md", 
        "description": "Получение технологических параметров из отчетов печей УВНК-9А1 и УВНК-9А2."
    },
    "UPPF": {
        "folder": "UPPF",
        "entry_point": "main_gui.py",
        "help_file": "📄 Руководство пользователя.md",
        "description": "Автоматическая обработка данных из Excel-отчетов печей УППФ."
    },
    "PointDeviation": {
        "folder": "point_deviation",
        "entry_point": "main_gui.py",
        "help_file": "📄 Руководство пользователя.md",
        "description": "Сортировка протоколов измерений (PDF/HTML) по величине отклонений."
    },
    "GreenTable": {
        "folder": "green_table",
        "entry_point": "main_gui.py",
        "help_file": "📄 Руководство пользователя.md",
        "description": "Сбор данных из отчетов плавок в единую таблицу."
    },
    "FeatherTurn": {
        "folder": "feather_turn",
        "entry_point": "main_gui.py",
        "help_file": "📄 Руководство пользователя.md",
        "description": "Обработка замеров разворота лопаток (первичный/перезамер)."
    },
    "TermodatReports": {
        "folder": "termodat_reports",
        "entry_point": "main_gui.py",
        "help_file": "📄 Руководство пользователя.md",
        "description": "Объединение CSV-отчетов с прибора Термодат в единые файлы."
    }
}

# === UNIFIED THEME ===
THEME_STYLESHEET = """
    QMainWindow, QWidget#central_widget {
        background-color: #2b2b2b;
    }
    
    /* Sidebar */
    QFrame#sidebar {
        background-color: #1e1e1e;
        border-right: 1px solid #333;
    }
    QLabel#sidebar_title {
        color: #3498db;
        font-family: "Segoe UI";
        font-size: 22px;
        font-weight: bold;
        padding: 20px;
    }
    
    /* Sidebar Buttons */
    QPushButton.sidebar_btn {
        background-color: transparent;
        color: #aaaaaa;
        text-align: left;
        padding: 15px 20px;
        border: none;
        font-size: 14px;
        font-family: "Segoe UI";
    }
    QPushButton.sidebar_btn:hover {
        background-color: #2c2c2c;
        color: white;
    }
    QPushButton.sidebar_btn:checked {
        background-color: #2b2b2b; /* Matches main content */
        color: #3498db;
        border-right: 3px solid #3498db;
        font-weight: bold;
    }
    
    /* Main Content */
    QLabel#program_title {
        color: #ffffff;
        font-family: "Segoe UI";
        font-size: 32px;
        font-weight: bold;
    }
    QLabel#program_desc {
        color: #cccccc;
        font-family: "Segoe UI";
        font-size: 16px;
    }
    QLabel#help_header {
        color: #ffffff;
        font-family: "Segoe UI";
        font-size: 18px;
        font-weight: bold;
        margin-top: 20px;
    }
    
    /* Launch Button */
    QPushButton#launch_btn {
        background-color: #3498db;
        color: white;
        font-size: 16px;
        font-weight: bold;
        padding: 15px 30px;
        border-radius: 8px;
        border: none;
    }
    QPushButton#launch_btn:hover {
        background-color: #2980b9;
    }
    QPushButton#launch_btn:pressed {
        background-color: #1f618d;
    }
    
    /* Help Viewer */
    QTextEdit#help_viewer {
        background-color: #1e1e1e;
        color: #dddddd;
        border-radius: 8px;
        border: 1px solid #333;
        padding: 15px;
        font-family: "Consolas", "Segoe UI", sans-serif;
        font-size: 13px;
    }
"""


class SidebarButton(QPushButton):
    def __init__(self, text, program_key, parent=None):
        super().__init__(text, parent)
        self.program_key = program_key
        self.setCheckable(True)
        self.setAutoExclusive(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setProperty("class", "sidebar_btn") 

class CustomTitleBar(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.window_parent = parent
        self.setFixedHeight(35)
        self.setStyleSheet("""
            QFrame { background-color: #1e1e1e; border-bottom: 1px solid #333; }
            QLabel { color: #aaaaaa; font-family: "Segoe UI"; font-size: 12px; font-weight: bold; }
            QPushButton { background: transparent; border: none; color: #aaaaaa; font-family: "Segoe UI"; font-size: 14px; }
            QPushButton:hover { background-color: #333; color: white; }
            QPushButton#btn_close:hover { background-color: #e81123; color: white; }
        """)
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 0, 0, 0)
        layout.setSpacing(0)
        
        # Icon/Title
        title = QLabel(APP_TITLE)
        layout.addWidget(title)
        layout.addStretch()
        
        # Window Controls
        self.btn_min = QPushButton("─")
        self.btn_min.setFixedSize(45, 35)
        self.btn_min.clicked.connect(self.window_parent.showMinimized)
        
        self.btn_max = QPushButton("☐")
        self.btn_max.setFixedSize(45, 35)
        self.btn_max.clicked.connect(self.toggle_max)
        
        self.btn_close = QPushButton("✕")
        self.btn_close.setObjectName("btn_close")
        self.btn_close.setFixedSize(45, 35)
        self.btn_close.clicked.connect(self.window_parent.close)
        
        layout.addWidget(self.btn_min)
        layout.addWidget(self.btn_max)
        layout.addWidget(self.btn_close)
        
        # Dragging variables
        self.click_pos = None

    def toggle_max(self):
        if self.window_parent.isMaximized():
            self.window_parent.showNormal()
            self.btn_max.setText("☐")
        else:
            self.window_parent.showMaximized()
            self.btn_max.setText("❐")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.click_pos = event.globalPosition().toPoint()

    def mouseMoveEvent(self, event):
        if self.click_pos is not None:
            delta = event.globalPosition().toPoint() - self.click_pos
            self.window_parent.move(self.window_parent.pos() + delta)
            self.click_pos = event.globalPosition().toPoint()

    def mouseReleaseEvent(self, event):
        self.click_pos = None

from PyQt6.QtWidgets import QSizeGrip

class NavigationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1100, 750)
        
        # Remove default frame
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowSystemMenuHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        # Main Container with Border for resize visibility
        self.main_container = QFrame()
        self.main_container.setObjectName("MainContainer")
        self.main_container.setStyleSheet(f"""
            QFrame#MainContainer {{
                background-color: #2b2b2b;
                border: 1px solid #444;
                border-radius: 0px;
            }}
            {THEME_STYLESHEET}
        """)
        self.setCentralWidget(self.main_container)
        
        # Main Vertical Layout
        self.main_vbox = QVBoxLayout(self.main_container)
        self.main_vbox.setContentsMargins(0, 0, 0, 0)
        self.main_vbox.setSpacing(0)
        
        # 1. Custom Title Bar
        self.title_bar = CustomTitleBar(self)
        self.main_vbox.addWidget(self.title_bar)
        
        # 2. Content Area
        self.content_widget = QWidget()
        self.content_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.main_vbox.addWidget(self.content_widget)
        
        self.setup_inner_ui(self.content_widget)
        
        # 3. Resize Grip (Bottom Right)
        bottom_bar = QHBoxLayout()
        bottom_bar.addStretch()
        self.grip = QSizeGrip(self.main_container) # Grip needs to be child of container to show
        bottom_bar.addWidget(self.grip, 0, Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignRight)
        # Using a specialized overlay for grip is better but this simple way works for corner
        
        # To make resize grip generic we often overlay it, but for simplicity let's rely on standard
        # A standard status bar has one.
        # Let's just create a small wrapper layout for the main content that includes the grip at bottom-right overlay
        # Or easier: minimal bottom margin logic.
        
        # Actually simpler: Just set window flags. 
        # But QSizeGrip needs to be in a layout or absolute position. 
        # We'll just leave it; the layout above puts it in a strip at bottom. 
        # Better visual: overlay.
        self.grip.setFixedSize(20, 20)
        self.grip.setStyleSheet("background: transparent;")
        
        # Re-doing main layout to put grip ON TOP of content or in corner
        # We will use the StatusBar trick which is easy for MainWindow
        self.setStatusBar(None) 
        
        # Manual grip placement
        self.grip.setParent(self.main_container)
        self.grip.move(self.width() - 20, self.height() - 20)
        
        self.current_program = None
        
        # Select first program
        if PROGRAMS:
            first = list(PROGRAMS.keys())[0]
            self.load_program(first)
            # Find the button and check it
            for btn in self.sidebar_layout.parentWidget().findChildren(SidebarButton):
                if btn.program_key == first:
                    btn.setChecked(True)
                    break

    def resizeEvent(self, event):
        # Update grip position
        if hasattr(self, 'grip'):
            self.grip.move(self.width() - 20, self.height() - 20)
        super().resizeEvent(event)

    def setup_inner_ui(self, parent_widget):
        main_layout = QHBoxLayout(parent_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # === SIDEBAR ===
        sidebar_frame = QFrame()
        sidebar_frame.setObjectName("sidebar")
        sidebar_frame.setFixedWidth(260)
        
        self.sidebar_layout = QVBoxLayout(sidebar_frame)
        self.sidebar_layout.setContentsMargins(0, 0, 0, 0)
        self.sidebar_layout.setSpacing(2)
        
        # Title
        title = QLabel("NAVIGATOR")
        title.setObjectName("sidebar_title")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.sidebar_layout.addWidget(title)
        
        # Program List
        for prog_name in PROGRAMS.keys():
            btn = SidebarButton(prog_name, prog_name)
            btn.clicked.connect(lambda checked, p=prog_name: self.load_program(p))
            self.sidebar_layout.addWidget(btn)
            
        self.sidebar_layout.addStretch() # Push items up
        
        # === MAIN CONTENT ===
        right_panel = QWidget()
        content_layout = QVBoxLayout(right_panel)
        content_layout.setContentsMargins(40, 40, 40, 40)
        content_layout.setSpacing(20)
        
        # Header Section
        self.lbl_title = QLabel("Select Program")
        self.lbl_title.setObjectName("program_title")
        content_layout.addWidget(self.lbl_title)
        
        self.lbl_desc = QLabel("Description...")
        self.lbl_desc.setObjectName("program_desc")
        self.lbl_desc.setWordWrap(True)
        content_layout.addWidget(self.lbl_desc)
        
        # Launch Button
        self.btn_launch = QPushButton("🚀 ЗАПУСТИТЬ")
        self.btn_launch.setObjectName("launch_btn")
        self.btn_launch.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_launch.setFixedWidth(200)
        self.btn_launch.clicked.connect(self.launch_current_program)
        content_layout.addWidget(self.btn_launch)
        
        # Help Section
        content_layout.addWidget(QLabel("Справка / Документация:", objectName="help_header"))
        
        self.help_viewer = QTextEdit()
        self.help_viewer.setObjectName("help_viewer")
        self.help_viewer.setReadOnly(True)
        content_layout.addWidget(self.help_viewer)
        
        # Add to main layout
        main_layout.addWidget(sidebar_frame)
        main_layout.addWidget(right_panel)

    def load_program(self, prog_key):
        self.current_program = prog_key
        config = PROGRAMS[prog_key]
        
        self.lbl_title.setText(prog_key)
        self.lbl_desc.setText(config.get("description", ""))
        
        # Load Help
        base_path = os.path.dirname(os.path.abspath(__file__))
        help_path = os.path.join(base_path, config["folder"], config["help_file"])
        
        if os.path.exists(help_path):
            try:
                with open(help_path, "r", encoding="utf-8") as f:
                    content = f.read()
                self.help_viewer.setMarkdown(content)
            except Exception as e:
                self.help_viewer.setPlainText(f"Ошибка чтения справки: {e}")
        else:
            self.help_viewer.setPlainText(f"Файл справки не найден:\n{help_path}")

    def launch_current_program(self):
        if not self.current_program:
            return
            
        config = PROGRAMS[self.current_program]
        base_path = os.path.dirname(os.path.abspath(__file__))
        prog_dir = os.path.join(base_path, config["folder"])
        script_path = os.path.join(prog_dir, config["entry_point"])
        
        if not os.path.exists(script_path):
            QMessageBox.critical(self, "Ошибка", f"Файл не найден:\n{script_path}")
            return
            
        self.btn_launch.setText("⏳ ЗАПУСК...")
        self.btn_launch.setEnabled(False)
        self.repaint() # Force update
        
        try:
            # Launch detached process
            subprocess.Popen([sys.executable, script_path], cwd=prog_dir)
            
            # Reset button state
            self.btn_launch.setText("🚀 ЗАПУСТИТЬ")
            self.btn_launch.setEnabled(True)
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка запуска", str(e))
            self.btn_launch.setText("🚀 ЗАПУСТИТЬ")
            self.btn_launch.setEnabled(True)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = NavigationApp()
    window.show()
    sys.exit(app.exec())

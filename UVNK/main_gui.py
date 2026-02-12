import sys
import threading
import os
import json
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QFileDialog, QTextEdit, QVBoxLayout, QHBoxLayout,
    QFrame, QMessageBox
)
from PyQt6.QtCore import Qt, pyqtSignal
from data_orchestrator import process_session_to_excel

# === UNIFIED THEME CONSTANTS ===
THEME_STYLESHEET = """
    QWidget {
        background-color: #2b2b2b;
        color: #ffffff;
        font-family: "Segoe UI", sans-serif;
        font-size: 14px;
    }
    
    /* Panels & Frames */
    QFrame#input_frame, QFrame#output_frame {
        background-color: #1e1e1e;
        border-radius: 10px;
        border: 1px solid #333333;
    }
    
    /* Labels */
    QLabel {
        color: #e0e0e0;
        font-weight: 600;
    }
    QLabel#header_label {
        font-size: 18px;
        color: #3498db;
        font-weight: bold;
    }
    
    /* Inputs */
    QLineEdit {
        padding: 8px 12px;
        border: 1px solid #3d3d3d;
        border-radius: 6px;
        background-color: #333333;
        color: white;
        font-size: 13px;
    }
    QLineEdit:focus {
        border-color: #3498db;
    }
    
    /* Buttons */
    QPushButton {
        padding: 8px 16px;
        border-radius: 6px;
        background-color: #3a3a3a;
        color: white;
        border: 1px solid #3d3d3d;
        font-weight: 600;
    }
    QPushButton:hover {
        background-color: #444444;
        border-color: #888888;
    }
    QPushButton:pressed {
        background-color: #2a2a2a;
    }
    
    /* Primary Action Button */
    QPushButton#primary_btn {
        background-color: #3498db;
        border-color: #2980b9;
    }
    QPushButton#primary_btn:hover {
        background-color: #2980b9;
        border-color: #1f618d;
    }
    
    /* Logs */
    QTextEdit {
        border-radius: 6px;
        border: 1px solid #3d3d3d;
        padding: 10px;
        background-color: #1e1e1e;
        color: #cccccc;
        font-family: "Consolas", monospace;
        font-size: 12px;
    }
"""

# === CUSTOM TITLE BAR CLASS ===
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
        self.title_label = QLabel("UVNK: Обработка данных")
        layout.addWidget(self.title_label)
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

from PyQt6.QtWidgets import QSizeGrip, QSizePolicy

class App(QWidget):
    log_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("UVNK: Обработка данных температур")
        self.resize(900, 700)
        
        # Frameless Setup
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowSystemMenuHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        # Main Layout (Fill entire widget)
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)
        
        # Main Container (Visible Border)
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
        outer_layout.addWidget(self.main_container)
        
        # Container Layout
        container_layout = QVBoxLayout(self.main_container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)
        
        # 1. Custom Title Bar
        self.title_bar = CustomTitleBar(self)
        container_layout.addWidget(self.title_bar)
        
        # 2. Content Area
        content_widget = QWidget()
        container_layout.addWidget(content_widget)
        
        self.init_inner_ui(content_widget)
        
        self.log_signal.connect(self.log_text.append)
        
        # 3. Resize Grip (Manual placement)
        self.grip = QSizeGrip(self.main_container)
        self.grip.setFixedSize(20, 20)
        self.grip.setStyleSheet("background: transparent;")
        
        self.load_last_paths()

    def resizeEvent(self, event):
        if hasattr(self, 'grip'):
            self.grip.move(self.width() - 20, self.height() - 20)
        super().resizeEvent(event)

    def init_inner_ui(self, parent_widget):
        main_layout = QVBoxLayout(parent_widget)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)

        # Header
        header = QLabel("Параметры обработки")
        header.setObjectName("header_label")
        main_layout.addWidget(header)

        # === Входная папка ===
        input_frame = QFrame()
        input_frame.setObjectName("input_frame")
        input_layout = QVBoxLayout(input_frame) 
        
        lbl_in = QLabel("📂 Папка с данными (Pasport/Reports)")
        
        h_in = QHBoxLayout()
        self.input_path = QLineEdit()
        self.input_path.setPlaceholderText("Выберите директорию с исходными файлами...")
        self.btn_input = QPushButton("Обзор...")
        self.btn_input.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_input.clicked.connect(self.select_input_folder)
        
        h_in.addWidget(self.input_path)
        h_in.addWidget(self.btn_input)
        
        input_layout.addWidget(lbl_in)
        input_layout.addLayout(h_in)
        main_layout.addWidget(input_frame)

        # === Выходной файл ===
        output_frame = QFrame()
        output_frame.setObjectName("output_frame")
        output_layout = QVBoxLayout(output_frame)
        
        lbl_out = QLabel("📄 Файл результата (.xlsx)")
        
        h_out = QHBoxLayout()
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("Укажите, куда сохранить результат...")
        self.btn_output = QPushButton("Выбрать...")
        self.btn_output.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_output.clicked.connect(self.select_output_file)
        
        h_out.addWidget(self.output_path)
        h_out.addWidget(self.btn_output)
        
        output_layout.addWidget(lbl_out)
        output_layout.addLayout(h_out)
        main_layout.addWidget(output_frame)

        # === Кнопка запуска ===
        self.run_button = QPushButton("🚀 ЗАПУСТИТЬ ОБРАБОТКУ")
        self.run_button.setObjectName("primary_btn")
        self.run_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.run_button.setMinimumHeight(50)
        self.run_button.clicked.connect(self.start_processing)
        main_layout.addWidget(self.run_button)

        # === Логи ===
        main_layout.addWidget(QLabel("Журнал событий:"))
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)

        folder_path = QFileDialog.getExistingDirectory(self, "Выберите папку сессии", "")
        if folder_path:
            self.input_path.setText(folder_path)
            self.save_last_paths()

    def select_output_file(self):
        current_path = self.output_path.text().strip()
        start_dir = current_path if current_path and os.path.exists(os.path.dirname(current_path)) else ""
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить отчет как...", start_dir,
            "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            if not file_path.lower().endswith('.xlsx'):
                file_path += '.xlsx'
            self.output_path.setText(file_path)
            self.save_last_paths()
            
    def load_last_paths(self):
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_paths.json")
        if os.path.exists(config_file):
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.input_path.setText(data.get("input", ""))
                    self.output_path.setText(data.get("output", ""))
            except Exception:
                pass

    def save_last_paths(self):
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_paths.json")
        data = {
            "input": self.input_path.text().strip(),
            "output": self.output_path.text().strip()
        }
        try:
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def start_processing(self):
        input_folder = self.input_path.text().strip()
        output_file = self.output_path.text().strip()
        
        if not input_folder or not output_file:
            QMessageBox.warning(self, "Внимание", "Необходимо заполнить оба поля:\n1. Папка с данными\n2. Файл результата")
            return

        # Disable UI to prevent double click
        self.run_button.setEnabled(False)
        self.run_button.setText("⏳ Выполняется обработка...")
        
        thread = threading.Thread(target=self.run_processing, args=(input_folder, output_file))
        thread.start()

    def run_processing(self, input_folder, output_file):
        try:
            process_session_to_excel(
                root_folder=input_folder,
                output_excel_path=output_file,
                log_callback=lambda msg: self.log_signal.emit(msg)
            )
            self.log_signal.emit("\n✅ Обработка успешно завершена!")
        except Exception as e:
            self.log_signal.emit(f"\n❌ КРИТИЧЕСКАЯ ОШИБКА: {str(e)}")
        finally:
            # Not easy to re-enable button from thread in PyQt thread-safety rules generally, 
            # but for simple cases in Python it sometimes works. 
            pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())
import sys
import threading
import os
import time
import json
import traceback
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QFileDialog, QTextEdit, QVBoxLayout, QHBoxLayout,
    QFrame, QMessageBox, QProgressBar, QSizeGrip
)
from PyQt6.QtCore import Qt, pyqtSignal

# Добавляем родительскую директорию в путь для импорта общих модулей
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from cache_manager import CacheManager
except ImportError:
    CacheManager = None

# Import logic from the processor file
from green_table_processor import GreenTableProcessor, clear_cache_for_output

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
        color: #2ecc71;
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
        border-color: #2ecc71;
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
        background-color: #2ecc71;
        border-color: #27ae60;
    }
    QPushButton#primary_btn:hover {
        background-color: #27ae60;
        border-color: #2ecc71;
    }

    /* Secondary Button (Clear Cache) */
    QPushButton#secondary_btn {
        background-color: #e67e22;
        border-color: #d35400;
    }
    QPushButton#secondary_btn:hover {
        background-color: #d35400;
        border-color: #a04000;
    }
    QPushButton#secondary_btn:disabled {
        background-color: #444444;
        color: #888888;
        border-color: #333333;
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

    /* Progress Bar */
    QProgressBar {
        border: 1px solid #3d3d3d;
        border-radius: 5px;
        text-align: center;
        background-color: #1e1e1e;
    }
    QProgressBar::chunk {
        background-color: #2ecc71;
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
        self.title_label = QLabel(self.window_parent.windowTitle())
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

class GreenTableApp(QWidget):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()
    progress_signal = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Green Table Data Collector")
        self.resize(1000, 700)
        
        # Frameless Setup
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowSystemMenuHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        # Main Layout
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)
        
        # Main Container
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
        self.progress_signal.connect(self.progress_bar.setValue)
        self.finished_signal.connect(self.on_processing_finished)

        # 3. Resize Grip
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
        header = QLabel("Обработка Excel файлов")
        header.setObjectName("header_label")
        main_layout.addWidget(header)

        # === Source Folder ===
        input_frame = QFrame()
        input_frame.setObjectName("input_frame")
        input_layout = QVBoxLayout(input_frame) 
        
        lbl_in = QLabel("📂 Папка с файлами (.xls, .xlsx)")
        
        h_in = QHBoxLayout()
        self.source_path = QLineEdit()
        self.source_path.setPlaceholderText("Выберите директорию для сканирования...")
        self.btn_source = QPushButton("Обзор...")
        self.btn_source.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_source.clicked.connect(self.select_source_folder)
        
        h_in.addWidget(self.source_path)
        h_in.addWidget(self.btn_source)
        
        input_layout.addWidget(lbl_in)
        input_layout.addLayout(h_in)
        main_layout.addWidget(input_frame)

        # === Output File ===
        output_frame = QFrame()
        output_frame.setObjectName("output_frame")
        output_layout = QVBoxLayout(output_frame)
        
        lbl_out = QLabel("📄 Итоговый Excel файл")
        
        h_out = QHBoxLayout()
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("Укажите путь для сохранения результата...")
        self.btn_output = QPushButton("Обзор...")
        self.btn_output.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_output.clicked.connect(self.select_output_file)
        
        h_out.addWidget(self.output_path)
        h_out.addWidget(self.btn_output)
        
        output_layout.addWidget(lbl_out)
        output_layout.addLayout(h_out)
        main_layout.addWidget(output_frame)

        # === Progress Bar ===
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)

        # === Action Buttons ===
        btns_layout = QHBoxLayout()
        
        self.run_button = QPushButton("🚀 НАЧАТЬ СБОР ДАННЫХ")
        self.run_button.setObjectName("primary_btn")
        self.run_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.run_button.setMinimumHeight(50)
        self.run_button.clicked.connect(self.start_processing)
        
        self.clear_cache_button = QPushButton("🧹 Очистить кэш")
        self.clear_cache_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_cache_button.setMinimumHeight(50)
        self.clear_cache_button.clicked.connect(self.clear_cache)
        
        btns_layout.addWidget(self.run_button, 4)
        btns_layout.addWidget(self.clear_cache_button, 1)
        main_layout.addLayout(btns_layout)

        # === Logs ===
        main_layout.addWidget(QLabel("Журнал событий:"))
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)

    def select_source_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите директорию")
        if folder:
            self.source_path.setText(folder)
            self.save_last_paths()

    def select_output_file(self):
        # Используем Variant 1: DontConfirmOverwrite
        options = QFileDialog.Option.DontConfirmOverwrite
        file, _ = QFileDialog.getSaveFileName(self, "Сохранить файл", self.output_path.text(), "Excel Files (*.xlsx)", options=options)
        if file:
            self.output_path.setText(file)
            self.save_last_paths()
            
    def load_last_paths(self):
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_paths.json")
        if os.path.exists(config_file):
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.source_path.setText(data.get("source", ""))
                    self.output_path.setText(data.get("output", "output.xlsx"))
            except Exception:
                pass

    def save_last_paths(self):
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_paths.json")
        data = {
            "source": self.source_path.text().strip(),
            "output": self.output_path.text().strip()
        }
        try:
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def clear_cache(self):
        output = self.output_path.text().strip()
        if not output:
             QMessageBox.warning(self, "Ошибка", "Укажите выходной файл для определения кэша!")
             return
        
        reply = QMessageBox.question(self, 'Очистка кэша', 
                                   f"Вы уверены, что хотите очистить кэш для файла {os.path.basename(output)}?\nЭто приведет к полной перечитке всех данных при следующем запуске.",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
                                   QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            clear_cache_for_output(output)
            self.log_text.append(f"[{time.strftime('%H:%M:%S')}] Кэш очищен.")
            QMessageBox.information(self, "Готово", "Кэш успешно очищен.")

    def start_processing(self):
        source = self.source_path.text().strip()
        output = self.output_path.text().strip()
        
        if not source:
            QMessageBox.warning(self, "Ошибка", "Выберите исходную директорию!")
            return
        if not output:
            QMessageBox.warning(self, "Ошибка", "Укажите выходной файл!")
            return

        self.run_button.setEnabled(False)
        self.clear_cache_button.setEnabled(False)
        self.run_button.setText("⏳ Идёт обработка...")
        self.log_text.clear()
        self.progress_bar.setValue(0)
        
        thread = threading.Thread(target=self.run_processing, args=(source, output))
        thread.start()

    def run_processing(self, source, output):
        processor = GreenTableProcessor(log_callback=self.log_callback)
        try:
            processor.process_directory(source, output, update_progress=self.update_progress)
        except Exception as e:
            self.log_signal.emit(f"Критическая ошибка: {e}")
            traceback.print_exc()
        finally:
            self.finished_signal.emit()

    def log_callback(self, message):
        self.log_signal.emit(message)
    
    def update_progress(self, value):
        self.progress_signal.emit(value)

    def on_processing_finished(self):
        self.run_button.setEnabled(True)
        self.clear_cache_button.setEnabled(True)
        self.run_button.setText("🚀 НАЧАТЬ СБОР ДАННЫХ")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GreenTableApp()
    window.show()
    sys.exit(app.exec())

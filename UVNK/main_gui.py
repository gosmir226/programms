import sys
import threading
import os
import json
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QFileDialog, QTextEdit, QVBoxLayout, QHBoxLayout,
    QFrame, QMessageBox, QSizeGrip
)
from PyQt6.QtCore import Qt, pyqtSignal

# Добавляем родительскую директорию в путь для импорта общих модулей
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from cache_manager import CacheManager
except ImportError:
    CacheManager = None

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

class App(QWidget):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("UVNK: Обработка данных температур")
        self.resize(900, 750)
        
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
        self.finished_signal.connect(self.on_processing_finished)
        
        # 3. Resize Grip
        self.grip = QSizeGrip(self.main_container)
        self.grip.setFixedSize(20, 20)
        
        self.load_last_paths()
        self.update_cache_button_state()

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
        input_layout.addWidget(QLabel("📂 Папка с данными (Pasport/Reports)"))
        
        h_in = QHBoxLayout()
        self.input_path = QLineEdit()
        self.input_path.setPlaceholderText("Выберите директорию с исходными файлами...")
        self.btn_input = QPushButton("Обзор...")
        self.btn_input.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_input.clicked.connect(self.select_input_folder)
        h_in.addWidget(self.input_path)
        h_in.addWidget(self.btn_input)
        input_layout.addLayout(h_in)
        main_layout.addWidget(input_frame)

        # === Выходной файл ===
        output_frame = QFrame()
        output_frame.setObjectName("output_frame")
        output_layout = QVBoxLayout(output_frame)
        output_layout.addWidget(QLabel("📄 Файл результата (.xlsx)"))
        
        h_out = QHBoxLayout()
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("Укажите, куда сохранить результат...")
        self.output_path.textChanged.connect(self.update_cache_button_state)
        self.btn_output = QPushButton("Выбрать...")
        self.btn_output.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_output.clicked.connect(self.select_output_file)
        h_out.addWidget(self.output_path)
        h_out.addWidget(self.btn_output)
        output_layout.addLayout(h_out)
        
        # === Кнопка очистки кэша ===
        self.btn_clear_cache = QPushButton("🧹 Очистить кэш")
        self.btn_clear_cache.setObjectName("secondary_btn")
        self.btn_clear_cache.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_clear_cache.clicked.connect(self.clear_cache)
        self.btn_clear_cache.setEnabled(False)
        output_layout.addWidget(self.btn_clear_cache)
        
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

    def select_input_folder(self):
        current_path = self.input_path.text().strip()
        start_dir = current_path if current_path and os.path.exists(current_path) else ""
        folder_path = QFileDialog.getExistingDirectory(self, "Выберите папку сессии", start_dir)
        if folder_path:
            self.input_path.setText(folder_path)
            self.save_last_paths()

    def select_output_file(self):
        current_path = self.output_path.text().strip()
        start_dir = current_path if current_path and os.path.exists(os.path.dirname(current_path)) else ""
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить отчет как...", start_dir,
            "Excel Files (*.xlsx);;All Files (*)",
            options=QFileDialog.Option.DontConfirmOverwrite
        )
        if file_path:
            if not file_path.lower().endswith('.xlsx'):
                file_path += '.xlsx'
            self.output_path.setText(file_path)
            self.save_last_paths()
            
    def update_cache_button_state(self):
        path = self.output_path.text().strip()
        self.btn_clear_cache.setEnabled(bool(path))

    def clear_cache(self):
        output_file = self.output_path.text().strip()
        if not output_file: return
            
        if CacheManager:
            cache = CacheManager("UVNK", output_file)
            if cache.clear_cache():
                self.log_text.append("✅ Кэш успешно очищен.")
                QMessageBox.information(self, "Успех", "Кэш для данного файла очищен.")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось очистить кэш.")
        else:
            QMessageBox.critical(self, "Ошибка", "Модуль управления кэшем не найден.")

    def load_last_paths(self):
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_paths.json")
        if os.path.exists(config_file):
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.input_path.setText(data.get("input", ""))
                    self.output_path.setText(data.get("output", ""))
            except Exception: pass

    def save_last_paths(self):
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_paths.json")
        data = {"input": self.input_path.text().strip(), "output": self.output_path.text().strip()}
        try:
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception: pass

    def start_processing(self):
        input_folder = self.input_path.text().strip()
        output_file = self.output_path.text().strip()
        
        if not input_folder or not output_file:
            QMessageBox.warning(self, "Внимание", "Необходимо заполнить оба поля!")
            return

        self.run_button.setEnabled(False)
        self.btn_clear_cache.setEnabled(False)
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
            self.finished_signal.emit()

    def on_processing_finished(self):
        self.run_button.setEnabled(True)
        self.update_cache_button_state()
        self.run_button.setText("🚀 ЗАПУСТИТЬ ОБРАБОТКУ")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())

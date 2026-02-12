import sys
import threading
import os
import time
import json
import traceback
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QFileDialog, QTextEdit, QVBoxLayout, QHBoxLayout,
    QFrame, QMessageBox
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtWidgets import QSizeGrip

# Import logic from the processor file
import point_processor as processor

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
        self.title_label = QLabel("Сортировка по отклонениям (Point Deviation)")
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

class PointDeviationApp(QWidget):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Point Deviation Sorter")
        self.resize(900, 700)
        
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
        header = QLabel("Параметры сортировки")
        header.setObjectName("header_label")
        main_layout.addWidget(header)

        # === Source Folder ===
        input_frame = QFrame()
        input_frame.setObjectName("input_frame")
        input_layout = QVBoxLayout(input_frame) 
        
        lbl_in = QLabel("📂 Исходная папка (PDF/HTML)")
        
        h_in = QHBoxLayout()
        self.source_path = QLineEdit()
        self.source_path.setPlaceholderText("Папка с протоколами измерений...")
        self.btn_source = QPushButton("Обзор...")
        self.btn_source.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_source.clicked.connect(self.select_source_folder)
        
        h_in.addWidget(self.source_path)
        h_in.addWidget(self.btn_source)
        
        input_layout.addWidget(lbl_in)
        input_layout.addLayout(h_in)
        main_layout.addWidget(input_frame)

        # === Target Folder ===
        output_frame = QFrame()
        output_frame.setObjectName("output_frame")
        output_layout = QVBoxLayout(output_frame)
        
        lbl_out = QLabel("📂 Целевая папка (Результат)")
        
        h_out = QHBoxLayout()
        self.target_path = QLineEdit()
        self.target_path.setPlaceholderText("Папка для отсортированных файлов...")
        self.btn_target = QPushButton("Обзор...")
        self.btn_target.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_target.clicked.connect(self.select_target_folder)
        
        h_out.addWidget(self.target_path)
        h_out.addWidget(self.btn_target)
        
        output_layout.addWidget(lbl_out)
        output_layout.addLayout(h_out)
        main_layout.addWidget(output_frame)

        # === Action Button ===
        self.run_button = QPushButton("🚀 НАЧАТЬ ОБРАБОТКУ")
        self.run_button.setObjectName("primary_btn")
        self.run_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.run_button.setMinimumHeight(50)
        self.run_button.clicked.connect(self.start_processing)
        main_layout.addWidget(self.run_button)

        # === Logs ===
        main_layout.addWidget(QLabel("Журнал событий:"))
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)

    def select_source_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите исходную папку")
        if folder:
            self.source_path.setText(folder)
            self.save_last_paths()

    def select_target_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите целевую папку")
        if folder:
            self.target_path.setText(folder)
            self.save_last_paths()
            
    def load_last_paths(self):
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_paths.json")
        if os.path.exists(config_file):
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.source_path.setText(data.get("source", ""))
                    self.target_path.setText(data.get("target", ""))
            except Exception:
                pass

    def save_last_paths(self):
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_paths.json")
        data = {
            "source": self.source_path.text().strip(),
            "target": self.target_path.text().strip()
        }
        try:
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def start_processing(self):
        source = self.source_path.text().strip()
        target = self.target_path.text().strip()
        
        if not source or not target:
            QMessageBox.warning(self, "Ошибка", "Необходимо выбрать обе папки!")
            return

        self.run_button.setEnabled(False)
        self.run_button.setText("⏳ Идёт обработка...")
        self.log_text.clear()
        
        thread = threading.Thread(target=self.run_processing, args=(source, target))
        thread.start()

    def run_processing(self, source_folder, target_folder):
        try:
            self.log_signal.emit(f"Начало обработки...")
            self.log_signal.emit(f"Источник: {source_folder}")
            self.log_signal.emit(f"Цель: {target_folder}\n")
            
            start_total = time.time()
            
            # 1. Load Decision Table
            self.log_signal.emit("1. Загрузка таблицы решений...")
            try:
                decision_table = processor.load_decision_table()
                self.log_signal.emit("✅ Таблица решений загружена.\n")
            except Exception as e:
                self.log_signal.emit(f"❌ Ошибка загрузки таблицы решений: {e}")
                return

            # 2. Process Files
            self.log_signal.emit("2. Сканирование и анализ файлов (это может занять время)...")
            result_df = processor.process_folder_parallel(source_folder, decision_table)
            self.log_signal.emit(f"✅ Анализ завершен. Обработано записей: {len(result_df)}\n")

            # 3. Apply Decision Table
            self.log_signal.emit("3. Определение статусов...")
            grouped_files = processor.apply_decision_table(result_df, decision_table)
            self.log_signal.emit("✅ Статусы определены.\n")

            # 4. Copy Files
            self.log_signal.emit("4. Копирование файлов в целевую структуру...")
            
            success_count = 0
            fail_count = 0
            
            for filename, data in grouped_files.items():
                if 'file_path' in data and data['file_path']:
                    src_path = data['file_path']
                else:
                    src_path = os.path.join(source_folder, filename)
                
                relative_path = data.get("relative_path", "")
                status = data["final_status"]
                
                if os.path.exists(src_path):
                    ok, dest = processor.move_or_copy_file_with_structure(
                        src_path, source_folder, target_folder, relative_path, status
                    )
                    
                    if ok:
                        success_count += 1
                        self.log_signal.emit(f"[OK] {filename} -> {status}")
                    else:
                        fail_count += 1
                        self.log_signal.emit(f"[ERR] {filename}")
                else:
                    self.log_signal.emit(f"[MISSING] {filename}")

            total_time = time.time() - start_total
            self.log_signal.emit(f"\n{'='*30}")
            self.log_signal.emit(f"🏁 ГОТОВО за {total_time:.2f} сек")
            self.log_signal.emit(f"✅ Скопировано: {success_count}")
            self.log_signal.emit(f"❌ Ошибки: {fail_count}")

            # Status Summary
            status_counts = {}
            for data in grouped_files.values():
                st = data["final_status"]
                status_counts[st] = status_counts.get(st, 0) + 1
            
            self.log_signal.emit("\n📊 Статистика:")
            for st, cnt in status_counts.items():
                self.log_signal.emit(f"   • {st}: {cnt}")

        except Exception as e:
            self.log_signal.emit(f"\n❌ КРИТИЧЕСКАЯ ОШИБКА: {str(e)}")
            self.log_signal.emit(traceback.format_exc())
            
        finally:
            self.finished_signal.emit()

    def on_processing_finished(self):
        self.run_button.setEnabled(True)
        self.run_button.setText("🚀 НАЧАТЬ ОБРАБОТКУ")

if __name__ == "__main__":
    # Multiprocessing support for Windows frozen apps
    from multiprocessing import freeze_support
    freeze_support()
    
    app = QApplication(sys.argv)
    window = PointDeviationApp()
    window.show()
    sys.exit(app.exec())

import os
import csv
import chardet
import datetime
import json
import traceback
import sys

# Добавляем родительскую директорию в путь для импорта общих модулей
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from excel_manager import ExcelManager
    from cache_manager import CacheManager
except ImportError:
    ExcelManager = None
    CacheManager = None

class TermodatProcessor:
    def __init__(self, input_dir, output_dir, log_callback=None, progress_callback=None):
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.log_callback = log_callback
        self.progress_callback = progress_callback

    def log_message(self, message):
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)

    def update_progress(self, value):
        if self.progress_callback:
            self.progress_callback(value)

    def detect_encoding(self, file_path):
        with open(file_path, 'rb') as f:
            raw_data = f.read()
            result = chardet.detect(raw_data)
            return result['encoding']

    def check_header_format(self, header_fields):
        # Ожидаемые поля в заголовке
        expected_fields = ["Timestamp", "TC405A_Real", "TC406A_Real", "AVERAGE_KILN_TEMPERATURE"]
        
        if len(header_fields) != len(expected_fields):
            return False
            
        for i, field in enumerate(header_fields):
            clean_field = field.strip().lower()
            clean_expected = expected_fields[i].strip().lower()
            
            if clean_field != clean_expected:
                return False
                
        return True

    def is_content_row(self, row):
        """Проверяет, является ли строка содержательной"""
        if not row or len(row) == 0:
            return False
        
        full_row = "".join(row)
        for char in full_row:
            if char not in [',', ';', '.', ' ', '\t', '\n', '\r']:
                return True
                
        return False

    def parse_timestamp(self, timestamp_str):
        """Парсит временную метку из строки"""
        try:
            if '.' in timestamp_str:
                return datetime.datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S.%f')
            else:
                return datetime.datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')
        except ValueError:
            try:
                return datetime.datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')
            except ValueError:
                return None

    def get_first_content_row(self, file_path, separator):
        """Возвращает первую содержательную строку файла"""
        encoding = self.detect_encoding(file_path)
        if not encoding:
            encoding = 'utf-8'
        
        with open(file_path, 'r', encoding=encoding, errors='replace') as f:
            f.readline() # Skip header
            
            for line in f:
                line = line.strip()
                if not line: continue
                    
                if separator == ',':
                    row = line.split(',')
                else:
                    row = line.split(';')
                
                if self.is_content_row(row):
                    return row
        return None

    def get_last_content_row(self, file_path, separator):
        """Возвращает последнюю содержательную строку файла"""
        encoding = self.detect_encoding(file_path)
        if not encoding:
            encoding = 'utf-8'
        
        with open(file_path, 'r', encoding=encoding, errors='replace') as f:
            f.readline() # Skip header
            
            lines = f.readlines()
            for line in reversed(lines):
                line = line.strip()
                if not line: continue
                    
                if separator == ',':
                    row = line.split(',')
                else:
                    row = line.split(';')
                
                if self.is_content_row(row):
                    return row
        return None

    def can_merge_with_previous(self, current_chain, next_file_path):
        """Проверяет, можно ли объединить следующий файл с текущей цепочкой"""
        if not current_chain:
            return False
        
        last_file_path = current_chain[-1][1]
        
        encoding = self.detect_encoding(last_file_path)
        if not encoding: encoding = 'utf-8'
        
        with open(last_file_path, 'r', encoding=encoding, errors='replace') as f:
            first_line = f.readline().strip()
        
        separator = ',' if ',' in first_line else ';'
        
        last_row = self.get_last_content_row(last_file_path, separator)
        if not last_row: return False
        
        first_row_next = self.get_first_content_row(next_file_path, separator)
        if not first_row_next: return False
        
        last_time = self.parse_timestamp(last_row[0])
        first_time_next = self.parse_timestamp(first_row_next[0])
        
        if not last_time or not first_time_next: return False
        
        time_diff = (first_time_next - last_time).total_seconds()
        
        # Допустимая разница: от 58 до 62 секунд
        return 58 <= time_diff <= 62

    def merge_and_process_chain(self, file_chain, output_path):
        """Объединяет и обрабатывает цепочку файлов"""
        try:
            import csv
            all_rows = []
            header = None
            separator = None
            
            for file_path in file_chain:
                encoding = self.detect_encoding(file_path)
                if not encoding: encoding = 'utf-8'
                
                # Читаем первую строку для определения разделителя
                with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                    first_line = f.readline().strip()
                
                if separator is None:
                    separator = ',' if ',' in first_line else ';'
                
                # Читаем файл
                with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                    if separator == ',':
                        reader = csv.reader(f)
                    else:
                        reader = csv.reader(f, delimiter=separator)
                    rows = list(reader)
                
                if not rows: continue
                
                # Проверяем заголовок
                if header is None:
                    if self.check_header_format(rows[0]):
                        header = rows[0]
                        all_rows.extend(rows[1:])
                    else:
                        self.log_message(f"Неверный заголовок в файле: {file_path}")
                        return False
                else:
                    all_rows.extend(rows[1:])
            
            if not header or not all_rows:
                self.log_message("Нет данных для обработки")
                return False
            
            # Преобразуем данные в legacy-формат
            new_header = ["Дата", "Время", "TC405A_Real", "TC406A_Real", "AVERAGE_KILN_TEMPERATURE"]
            new_rows = [new_header]
            
            for row in all_rows:
                if len(row) < 4: continue
                
                try:
                    # Парсим временную метку
                    dt = self.parse_timestamp(row[0])
                    if not dt: continue
                    
                    # Форматируем как в оригинале
                    date_str = dt.strftime('%d.%m.%Y')
                    time_str = dt.strftime('%H:%M:%S')
                    
                    # Заменяем точки на запятые в числах
                    new_row = [
                        date_str,
                        time_str,
                        str(row[1]).replace('.', ','),
                        str(row[2]).replace('.', ','),
                        str(row[3]).replace('.', ',')
                    ]
                    new_rows.append(new_row)
                except Exception:
                    continue
            
            # Записываем результат именно как CSV с разделителем ';'
            output_csv = output_path if output_path.lower().endswith('.csv') else output_path + '.csv'
            
            with open(output_csv, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerows(new_rows)
            
            self.log_message(f"Сохранено: {os.path.basename(output_csv)}")
            return True
            
        except Exception as e:
            self.log_message(f"Ошибка при обработке цепочки: {str(e)}")
            self.log_message(f"Трассировка: {traceback.format_exc()}")
            return False

    def process(self):
        try:
            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)
            
            # Инициализируем CacheManager
            # В TermodatReports кэш хранится в выходной директории
            cache_path = os.path.join(self.output_dir, '.termodat_cache.json')
            if CacheManager:
                cache = CacheManager("termodat_reports", cache_path)
            else:
                cache = None
            
            date_folders = []
            for item in os.listdir(self.input_dir):
                item_path = os.path.join(self.input_dir, item)
                if os.path.isdir(item_path):
                    # Проверяем, есть ли CSV внутри и изменился ли он
                    csv_files = [f for f in os.listdir(item_path) if f.lower().endswith('.csv')]
                    if csv_files:
                        csv_path = os.path.abspath(os.path.join(item_path, csv_files[0]))
                        if cache and not cache.is_file_changed(csv_path):
                            continue
                        date_folders.append(item)
            
            date_folders.sort()
            self.log_message(f"Найдено новых/измененных папок: {len(date_folders)}")
            
            if len(date_folders) == 0:
                self.log_message("Все данные актуальны.")
                return 0
            
            processed_count = 0
            current_chain = []
            newly_processed_files = []
            
            for i, folder in enumerate(date_folders):
                folder_path = os.path.join(self.input_dir, folder)
                csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]
                
                if not csv_files: continue
                
                csv_file = os.path.join(folder_path, csv_files[0])
                abs_csv_path = os.path.abspath(csv_file)
                
                if self.can_merge_with_previous(current_chain, csv_file):
                    current_chain.append((folder, csv_file))
                    self.log_message(f"  Файл добавлен в цепочку: {folder}")
                else:
                    if current_chain:
                        res = self.process_chain(current_chain, cache)
                        if res: processed_count += 1
                    
                    current_chain = [(folder, csv_file)]
                    self.log_message(f"  Начата новая цепочка: {folder}")
                
                self.update_progress(int((i + 1) / len(date_folders) * 100))
            
            if current_chain:
                res = self.process_chain(current_chain, cache)
                if res: processed_count += 1
            
            return processed_count
            
        except Exception as e:
            self.log_message(f"Ошибка в процессе обработки: {str(e)}")
            self.log_message(f"Трассировка: {traceback.format_exc()}")
            return 0

    def process_chain(self, chain, cache):
        start_date = chain[0][0]
        end_date = chain[-1][0]
        output_filename = f"{start_date}-{end_date}_merged.csv"
        output_path = os.path.join(self.output_dir, output_filename)
        
        file_paths = [file_info[1] for file_info in chain]
        if self.merge_and_process_chain(file_paths, output_path):
            if cache:
                for fp in file_paths:
                    cache.update_file(os.path.abspath(fp))
            return True
        return False

def clear_cache_for_output(output_dir):
    cache_path = os.path.join(output_dir, '.termodat_cache.json')
    if CacheManager:
        cache = CacheManager("termodat_reports", cache_path)
        cache.clear_cache()

import os
import csv
import chardet
import datetime
import json
import traceback

class TermodatProcessor:
    def __init__(self, input_dir, output_dir, log_callback=None, progress_callback=None):
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.log_callback = log_callback
        self.progress_callback = progress_callback
        self.cache_file = os.path.join(output_dir, '.processing_cache.json')

    def log_message(self, message):
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)

    def update_progress(self, value):
        if self.progress_callback:
            self.progress_callback(value)

    def load_processed_cache(self):
        """Загружает кэш обработанных папок"""
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r') as f:
                    return set(json.load(f))
            except:
                return set()
        return set()

    def save_processed_cache(self, processed_folders):
        """Сохраняет кэш обработанных папок"""
        try:
            with open(self.cache_file, 'w') as f:
                json.dump(list(processed_folders), f)
        except Exception as e:
            self.log_message(f"Ошибка при сохранении кэша: {str(e)}")

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
            all_rows = []
            header = None
            separator = None
            
            for file_path in file_chain:
                encoding = self.detect_encoding(file_path)
                if not encoding: encoding = 'utf-8'
                
                with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                    first_line = f.readline().strip()
                
                if separator is None:
                    separator = ',' if ',' in first_line else ';'
                
                with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                    if separator == ',':
                        reader = csv.reader(f)
                    else:
                        reader = csv.reader(f, delimiter=separator)
                    rows = list(reader)
                
                if not rows: continue
                
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
            
            new_rows = [["Дата", "Время", "TC405A_Real", "TC406A_Real", "AVERAGE_KILN_TEMPERATURE"]]
            
            for row in all_rows:
                if len(row) < 4: continue
                
                try:
                    timestamp_str = row[0]
                    if '.' in timestamp_str:
                        dt = datetime.datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S.%f')
                    else:
                        dt = datetime.datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')
                    
                    date_str = dt.strftime('%d.%m.%Y')
                    time_str = dt.strftime('%H:%M:%S')
                    
                    new_row = [
                        date_str,
                        time_str,
                        row[1].replace('.', ','),
                        row[2].replace('.', ','),
                        row[3].replace('.', ',')
                    ]
                    new_rows.append(new_row)
                except ValueError as e:
                    continue
            
            with open(output_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerows(new_rows)
            
            return True
            
        except Exception as e:
            self.log_message(f"Ошибка при обработке цепочки: {str(e)}")
            self.log_message(f"Трассировка: {traceback.format_exc()}")
            return False

    def process(self):
        try:
            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)
            
            processed_folders = self.load_processed_cache()
            self.log_message(f"Загружено обработанных папок из кэша: {len(processed_folders)}")
            
            date_folders = []
            for item in os.listdir(self.input_dir):
                item_path = os.path.join(self.input_dir, item)
                if os.path.isdir(item_path) and item not in processed_folders:
                    date_folders.append(item)
            
            date_folders.sort()
            self.log_message(f"Найдено новых папок для обработки: {len(date_folders)}")
            
            if len(date_folders) == 0:
                return 0
            
            processed_count = 0
            current_chain = []
            newly_processed = set()
            
            for i, folder in enumerate(date_folders):
                folder_path = os.path.join(self.input_dir, folder)
                
                csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]
                
                if not csv_files:
                    self.log_message(f"CSV файл не найден в папке: {folder_path}")
                    continue
                
                csv_file = os.path.join(folder_path, csv_files[0])
                
                if self.can_merge_with_previous(current_chain, csv_file):
                    current_chain.append((folder, csv_file))
                    self.log_message(f"Файл {csv_file} добавлен в цепочку объединения")
                else:
                    if current_chain:
                        start_date = current_chain[0][0]
                        end_date = current_chain[-1][0]
                        output_filename = f"{start_date}-{end_date}_merged.csv"
                        output_path = os.path.join(self.output_dir, output_filename)
                        
                        file_paths = [file_info[1] for file_info in current_chain]
                        if self.merge_and_process_chain(file_paths, output_path):
                            processed_count += 1
                            for folder_info in current_chain:
                                newly_processed.add(folder_info[0])
                            self.log_message(f"Обработана цепочка из {len(current_chain)} файлов")
                    
                    current_chain = [(folder, csv_file)]
                    self.log_message(f"Начата новая цепочка с файла: {csv_file}")
                
                self.update_progress(int((i + 1) / len(date_folders) * 100))
            
            if current_chain:
                start_date = current_chain[0][0]
                end_date = current_chain[-1][0]
                output_filename = f"{start_date}-{end_date}_merged.csv"
                output_path = os.path.join(self.output_dir, output_filename)
                
                file_paths = [file_info[1] for file_info in current_chain]
                if self.merge_and_process_chain(file_paths, output_path):
                    processed_count += 1
                    for folder_info in current_chain:
                        newly_processed.add(folder_info[0])
                    self.log_message(f"Обработана цепочка из {len(current_chain)} файлов")
            
            if newly_processed:
                processed_folders.update(newly_processed)
                self.save_processed_cache(processed_folders)
                self.log_message(f"Добавлено в кэш: {len(newly_processed)} папок")
            
            return processed_count
            
        except Exception as e:
            self.log_message(f"Ошибка в процессе обработки: {str(e)}")
            self.log_message(f"Трассировка: {traceback.format_exc()}")
            return 0

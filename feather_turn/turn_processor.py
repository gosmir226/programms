import os
import pandas as pd
import json
import hashlib
import re
from datetime import datetime
from openpyxl import load_workbook
import traceback

class TurnProcessor:
    def __init__(self, directory, log_callback=None, progress_callback=None):
        self.directory = directory
        self.log_callback = log_callback
        self.progress_callback = progress_callback
        self.cache_file = os.path.join(directory, 'processing_cache.json')
        self.cache_data = {}

    def log_message(self, message):
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)

    def update_progress(self, value):
        if self.progress_callback:
            self.progress_callback(value)

    def get_file_hash(self, filepath):
        """Вычисляет хэш содержимого файла"""
        try:
            hash_sha = hashlib.sha256()
            with open(filepath, "rb") as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_sha.update(chunk)
            return hash_sha.hexdigest()
        except Exception as e:
            self.log_message(f"Ошибка при вычислении хэша {filepath}: {str(e)}")
            return None

    def load_cache(self):
        """Загружает кэш из JSON файла"""
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    self.cache_data = json.load(f)
                self.log_message(f"Загружено {len(self.cache_data)} записей из кэша")
                return True
            except Exception as e:
                self.log_message(f"Ошибка при загрузке кэша: {str(e)}")
                self.cache_data = {}
                return False
        else:
            self.cache_data = {}
            return True

    def save_cache(self):
        """Сохраняет кэш в JSON файл"""
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.cache_data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            self.log_message(f"Ошибка при сохранении кэша: {str(e)}")
            return False

    def is_file_modified(self, file_path):
        """Проверяет, был ли файл изменен с момента последней обработки"""
        if file_path not in self.cache_data:
            return True  # Файл новый
        
        cache_entry = self.cache_data[file_path]
        
        try:
            # Проверяем время модификации
            current_mtime = os.path.getmtime(file_path)
            if current_mtime != cache_entry.get("last_modified"):
                return True
            
            # Проверяем хэш содержимого
            current_hash = self.get_file_hash(file_path)
            if current_hash != cache_entry.get("hash"):
                return True
            
            return False  # Файл не изменился
            
        except Exception as e:
            self.log_message(f"Ошибка при проверке файла {file_path}: {str(e)}")
            return True  # В случае ошибки считаем файл измененным

    def update_cache(self, file_path):
        """Обновляет кэш для файла"""
        try:
            current_mtime = os.path.getmtime(file_path)
            current_hash = self.get_file_hash(file_path)
            
            self.cache_data[file_path] = {
                "last_modified": current_mtime,
                "hash": current_hash,
                "last_processed": datetime.now().isoformat(),
                "file_size": os.path.getsize(file_path)
            }
            # Сохраняем кэш после каждого обновления
            self.save_cache()
            
        except Exception as e:
            self.log_message(f"Ошибка при обновлении кэша для {file_path}: {str(e)}")

    def clean_header(self, header):
        """Очищает заголовок от лишних пробелов"""
        if pd.isna(header):
            return ""
        header_str = str(header).strip()
        header_str = re.sub(r'\s+', ' ', header_str)
        return header_str

    def get_visible_sheets(self, file_path):
        """Возвращает список видимых листов в файле Excel"""
        try:
            wb = load_workbook(file_path, read_only=True)
            visible_sheets = []
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                if sheet.sheet_state == 'visible':
                    visible_sheets.append(sheet_name)
            wb.close()
            return visible_sheets
        except Exception as e:
            self.log_message(f"Ошибка при определении видимых листов {file_path}: {str(e)}")
            return []

    def find_table_start(self, df):
        """Находит начальные координаты таблицы в DataFrame"""
        for row_idx in range(len(df)):
            for col_idx in range(len(df.columns)):
                if pd.notna(df.iloc[row_idx, col_idx]):
                    row_values = df.iloc[row_idx, col_idx:]
                    if row_values.notna().sum() >= 2:
                        return row_idx, col_idx
        return None, None

    def standardize_headers(self, headers):
        """Стандартизирует заголовки согласно mapping с обработкой дубликатов"""
        column_mapping = {
            'Плавка': ['Плавка'],
            'Отливка': ['Отливка', '№ отливки'],
            'Замер факт': ['Замер факт', 'Значение разворота по приспособлению ОДК-Климов', 'Значение разворота по приспособлению к4', 'Значение разворота по приспособлению к2'],
            'Статус': ['Статус4', 'Статус2', 'Статус'],
            'Контролер': ['Контролер', 'Исполнитель'],
            'Замер с поправкой': ['Замер с поправкой'],
            'Коэффициент': ['Коэффициент'],
            'Замер после доработки': ['Замер после доработки'],
            'Статус после доработки': ['Статус после доработки'],
            'после АТОС': ['после АТОС']
        }
        
        reverse_mapping = {}
        for standard_name, variants in column_mapping.items():
            for variant in variants:
                reverse_mapping[self.clean_header(variant)] = standard_name
        
        standardized_headers = []
        used_headers = set()
        
        for header in headers:
            clean_header = self.clean_header(header)
            
            if clean_header in reverse_mapping:
                standard_name = reverse_mapping[clean_header]
            else:
                standard_name = clean_header
            
            if standard_name in used_headers:
                counter = 1
                new_name = f"{standard_name}_{counter}"
                while new_name in used_headers:
                    counter += 1
                    new_name = f"{standard_name}_{counter}"
                standardized_headers.append(new_name)
                used_headers.add(new_name)
            else:
                standardized_headers.append(standard_name)
                used_headers.add(standard_name)
        
        return standardized_headers

    def process_coefficient_column(self, df, headers):
        """Обрабатывает столбец Коэффициент - берет значение из первой строки"""
        if 'Коэффициент' in headers:
            try:
                coeff_idx = list(headers).index('Коэффициент')
                if len(df) > 0:
                    first_value = df.iloc[0, coeff_idx]
                    if pd.notna(first_value):
                        df.iloc[:, coeff_idx] = first_value
            except (ValueError, IndexError):
                pass
        return df

    def extract_melt_info(self, melt_string):
        """Извлекает год и номер плавки из строки формата 25В11"""
        if pd.isna(melt_string):
            return None, None
        
        melt_str = str(melt_string).strip()
        
        match = re.search(r'(\d+)[В](\d+)', melt_str, re.IGNORECASE)
        if match:
            year = match.group(1)
            number = match.group(2)
            return year, number
        else:
            parts = re.split(r'[^\d]', melt_str, 1)
            if len(parts) >= 2:
                return parts[0], parts[1]
            elif len(parts) == 1:
                return parts[0], None
            else:
                return None, None

    def determine_status(self, value):
        """Определяет статус на основе числового значения"""
        try:
            num_value = float(value)
            
            if (-50 <= num_value <= -45) or (35 <= num_value <= 50):
                return "ОТРЫВ"
            elif (-40 <= num_value <= -30) or (20 <= num_value <= 30):
                return "ТР"
            elif -25 <= num_value <= 15:
                return "ГОД"
            else:
                return "БРАК"
        except (ValueError, TypeError):
            return "БРАК"

    def process_atos_rejection(self, df):
        """Обрабатывает столбцы для определения брака АТОС и обновления статуса"""
        standard_statuses = ["ГОД", "БРАК", "ТР", "ОТРЫВ", "год", "брак", "тр", "отрыв"]
        df['Брак АТОС'] = None
        
        atos_columns = ['Замер', 'Замер факт', 'Значение разворота по приспособлению ОДК-Климов', 
                       'Значение разворота по приспособлению к4', 'Значение разворота по приспособлению к2']
        
        correction_columns = ['Замер с поправкой']
        
        for idx, row in df.iterrows():
            current_status = str(row.get('Статус', '')).strip() if pd.notna(row.get('Статус')) else ''
            contains_viz = 'ВИЗ' in current_status.upper()
            
            atos_detected = False
            for col in atos_columns:
                if col in df.columns and pd.notna(row.get(col)):
                    cell_value = str(row[col]).lower()
                    if 'атос' in cell_value or 'atos' in cell_value:
                        atos_detected = True
                        break
            
            if current_status not in standard_statuses and not contains_viz:
                measurement_value = None
                
                if atos_detected:
                    for col in correction_columns:
                        if col in df.columns and pd.notna(row.get(col)):
                            try:
                                measurement_value = float(row[col])
                                break
                            except (ValueError, TypeError):
                                continue
                else:
                    for col in correction_columns:
                        if col in df.columns and pd.notna(row.get(col)):
                            try:
                                measurement_value = float(row[col])
                                break
                            except (ValueError, TypeError):
                                continue
                    
                    if measurement_value is None:
                        for col in atos_columns:
                            if col in df.columns and pd.notna(row.get(col)):
                                try:
                                    measurement_value = float(row[col])
                                    break
                                except (ValueError, TypeError):
                                    continue
                
                if measurement_value is not None:
                    new_status = self.determine_status(measurement_value)
                    df.at[idx, 'Брак АТОС'] = "Брак АТОС"
                    df.at[idx, 'Статус'] = new_status
        
        return df

    def determine_file_types_in_folder(self, folder_path):
        """Определяет тип файлов в папке"""
        xlsx_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.xlsx')]
        file_types = {}
        
        primary_by_name = []
        remeasure_by_name = []
        undetermined = []
        
        for filename in xlsx_files:
            lower_name = filename.lower()
            if 'перезамер' in lower_name or 'доработк' in lower_name:
                remeasure_by_name.append(filename)
                file_types[filename] = 'перезамер'
            elif 'угол разворота' in lower_name:
                primary_by_name.append(filename)
                file_types[filename] = 'первичный'
            else:
                undetermined.append(filename)
        
        if undetermined:
            if len(undetermined) == 1:
                file_types[undetermined[0]] = 'первичный'
            else:
                file_stats = []
                for filename in undetermined:
                    file_path = os.path.join(folder_path, filename)
                    sheets_count, total_rows = self.analyze_file_structure(file_path)
                    file_stats.append({
                        'filename': filename,
                        'sheets_count': sheets_count,
                        'total_rows': total_rows
                    })
                
                file_stats.sort(key=lambda x: (x['sheets_count'], x['total_rows']), reverse=True)
                
                max_sheets = max(stats['sheets_count'] for stats in file_stats)
                files_with_max_sheets = [stats for stats in file_stats if stats['sheets_count'] == max_sheets]
                
                if len(files_with_max_sheets) == 1:
                    file_types[files_with_max_sheets[0]['filename']] = 'перезамер'
                    for stats in file_stats:
                        if stats['filename'] != files_with_max_sheets[0]['filename']:
                             file_types[stats['filename']] = 'первичный'
                else:
                    files_with_max_sheets.sort(key=lambda x: x['total_rows'], reverse=True)
                    file_types[files_with_max_sheets[0]['filename']] = 'первичный'
                    for stats in files_with_max_sheets[1:]:
                        file_types[stats['filename']] = 'перезамер'
        
        return file_types

    def analyze_file_structure(self, file_path):
        """Анализирует структуру файла"""
        try:
            workbook = load_workbook(file_path, read_only=True)
            sheets_count = len(workbook.sheetnames)
            total_rows = 0
            
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                row_count = 0
                for row in sheet.iter_rows():
                    if any(cell.value is not None for cell in row):
                        row_count += 1
                total_rows += row_count
            
            workbook.close()
            return sheets_count, total_rows
        except Exception as e:
            self.log_message(f"Ошибка при анализе файла {file_path}: {str(e)}")
            return 0, 0

    def extract_cassette_and_measurement_info(self, file_path, sheet_name, file_type, measurement_rank=1):
        """Извлекает номер кассеты и номер замера"""
        if file_type == 'первичный':
            cassette_number = self.extract_cassette_from_filename(file_path)
            measurement_number = 0
        else:
            cassette_number = self.extract_cassette_from_sheetname(sheet_name)
            measurement_number = measurement_rank
        
        return cassette_number, measurement_number, file_type == 'перезамер'

    def extract_cassette_from_filename(self, file_path):
        filename = os.path.basename(file_path)
        match = re.search(r'№\s*(\d+)', filename)
        if match:
            return match.group(1)
        return None

    def extract_cassette_from_sheetname(self, sheet_name):
        match = re.search(r'№\s*(\d+)', sheet_name)
        if match:
            return match.group(1)
        return None

    def process_excel_file(self, file_path, file_type):
        """Обрабатывает один Excel-файл"""
        try:
            visible_sheets = self.get_visible_sheets(file_path)
            if not visible_sheets:
                return pd.DataFrame()
            
            excel_file = pd.ExcelFile(file_path)
            file_dataframes = []
            
            # Для перезамеров
            sheet_stats = []
            if file_type == 'перезамер':
                for sheet_name in visible_sheets:
                    if sheet_name not in excel_file.sheet_names: continue
                        
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')
                    start_row, start_col = self.find_table_start(df)
                    if start_row is None: continue
                    
                    table_df = df.iloc[start_row:, start_col:].copy().reset_index(drop=True)
                    if len(table_df) > 0:
                        original_headers = [str(h) if pd.notna(h) else f"Unnamed_{i}" for i, h in enumerate(table_df.iloc[0])]
                        table_df.columns = original_headers
                        table_df = table_df.drop(0).reset_index(drop=True)
                    
                    if len(table_df) == 0 or len(table_df.columns) == 0: continue
                    
                    if 'Плавка' in table_df.columns:
                        table_df = table_df[table_df['Плавка'].notna()]
                        table_df = table_df[table_df['Плавка'].astype(str).str.strip() != '']
                    
                    sheet_stats.append({'sheet_name': sheet_name, 'row_count': len(table_df), 'table_df': table_df})
            
                sheet_stats.sort(key=lambda x: x['row_count'], reverse=True)
                
                for rank, stat in enumerate(sheet_stats, 1):
                    table_df = stat['table_df']
                    sheet_name = stat['sheet_name']
                    self.process_dataframe(table_df, file_path, sheet_name, file_type, 
                                            rank, file_dataframes)
            else:
                # Первичный файл
                for sheet_name in visible_sheets:
                    if sheet_name not in excel_file.sheet_names: continue
                    
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')
                    start_row, start_col = self.find_table_start(df)
                    if start_row is None: continue
                    
                    table_df = df.iloc[start_row:, start_col:].copy().reset_index(drop=True)
                    if len(table_df) > 0:
                        original_headers = [str(h) if pd.notna(h) else f"Unnamed_{i}" for i, h in enumerate(table_df.iloc[0])]
                        table_df.columns = original_headers
                        table_df = table_df.drop(0).reset_index(drop=True)
                    
                    if len(table_df) == 0 or len(table_df.columns) == 0: continue
                    
                    if 'Плавка' in table_df.columns:
                        table_df = table_df[table_df['Плавка'].notna()]
                        table_df = table_df[table_df['Плавка'].astype(str).str.strip() != '']
                    
                    if len(table_df) == 0: continue
                    
                    self.process_dataframe(table_df, file_path, sheet_name, file_type, 
                                            1, file_dataframes)
            
            if file_dataframes:
                return pd.concat(file_dataframes, ignore_index=True, sort=False)
            else:
                return pd.DataFrame()
            
        except Exception as e:
            self.log_message(f"Ошибка при обработке файла {file_path}: {str(e)}")
            return pd.DataFrame()

    def process_dataframe(self, table_df, file_path, sheet_name, file_type, rank, file_dataframes):
        """Вспомогательная функция для обработки DataFrame"""
        standardized_headers = self.standardize_headers(table_df.columns)
        table_df.columns = standardized_headers
        
        table_df = self.process_coefficient_column(table_df, standardized_headers)
        
        if 'Плавка' in table_df.columns:
            melt_info = table_df['Плавка'].apply(self.extract_melt_info)
            table_df['Год плавки'] = melt_info.apply(lambda x: x[0] if x else None)
            table_df['Номер плавки'] = melt_info.apply(lambda x: x[1] if x else None)
        
        table_df = self.process_atos_rejection(table_df)
        
        cassette_number, measurement_number, is_remeasure = self.extract_cassette_and_measurement_info(
            file_path, sheet_name, file_type, rank
        )
        
        table_df['Имя файла'] = os.path.basename(file_path)
        table_df['Номер замера'] = measurement_number
        table_df['Номер кассеты'] = cassette_number
        table_df['Тип файла'] = 'Перезамер' if is_remeasure else 'Первичный'
        table_df['Полный путь'] = file_path
        
        if 'Плавка' in table_df.columns:
            table_df = table_df.rename(columns={'Плавка': 'Полный номер плавки'})
        
        file_dataframes.append(table_df)

    def process(self):
        try:
            self.log_message("Загрузка кэша...")
            if not self.load_cache():
                self.log_message("Ошибка при загрузке кэша")
                return pd.DataFrame()
            
            result_path = os.path.join(self.directory, 'consolidated_results.xlsx')
            existing_data = pd.DataFrame()
            if os.path.exists(result_path):
                try:
                    existing_data = pd.read_excel(result_path, engine='openpyxl')
                    self.log_message(f"Загружено {len(existing_data)} записей из существующих результатов")
                except Exception as e:
                    self.log_message(f"Ошибка при загрузке существующих результатов: {str(e)}")
            
            new_data = pd.DataFrame()
            modified_files = []
            folders = [f for f in os.listdir(self.directory) 
                      if os.path.isdir(os.path.join(self.directory, f))]
            total_folders = len(folders)
            
            processed_files_count = 0
            skipped_files_count = 0
            modified_files_count = 0
            
            for folder_idx, folder in enumerate(folders):
                folder_path = os.path.join(self.directory, folder)
                
                self.log_message(f"Анализ типов файлов в папке: {folder}")
                file_types = self.determine_file_types_in_folder(folder_path)
                
                xlsx_files = [f for f in os.listdir(folder_path) 
                             if f.lower().endswith('.xlsx')]
                
                if not xlsx_files: continue
                
                xlsx_files_with_paths = [(f, os.path.join(folder_path, f)) for f in xlsx_files]
                xlsx_files_with_paths.sort(key=lambda x: os.path.getmtime(x[1]))
                
                for file_idx, (filename, file_path) in enumerate(xlsx_files_with_paths):
                    file_type = file_types.get(filename, 'первичный')
                    
                    if not self.is_file_modified(file_path):
                        self.log_message(f"Пропуск (не изменился): {folder}/{filename}")
                        skipped_files_count += 1
                        continue
                    
                    self.log_message(f"Обработка ({file_type}): {folder}/{filename}")
                    modified_files.append(file_path)
                    
                    file_data = self.process_excel_file(file_path, file_type)
                    
                    if not file_data.empty:
                        if new_data.empty:
                            new_data = file_data
                        else:
                            new_data = pd.concat([new_data, file_data], ignore_index=True, sort=False)
                        
                        self.update_cache(file_path)
                        processed_files_count += 1
                        
                        if file_path in self.cache_data:
                            modified_files_count += 1
                    else:
                        self.update_cache(file_path)
                
                self.update_progress(int((folder_idx + 1) / total_folders * 100))
            
            if not existing_data.empty and modified_files:
                self.log_message("Удаление старых записей измененных файлов...")
                keep_mask = pd.Series([True] * len(existing_data), index=existing_data.index)
                
                for modified_file in modified_files:
                    modified_filename = os.path.basename(modified_file)
                    if 'Имя файла' in existing_data.columns:
                        file_mask = existing_data['Имя файла'] != modified_filename
                        keep_mask = keep_mask & file_mask
                    if 'Полный путь' in existing_data.columns:
                        path_mask = existing_data['Полный путь'] != modified_file
                        keep_mask = keep_mask & path_mask
                
                existing_data = existing_data[keep_mask].reset_index(drop=True)
            
            if not existing_data.empty and not new_data.empty:
                final_data = pd.concat([existing_data, new_data], ignore_index=True, sort=False)
            elif not existing_data.empty:
                final_data = existing_data
            else:
                final_data = new_data
            
            if not final_data.empty:
                final_data = final_data.drop_duplicates()
                final_data = final_data.reset_index(drop=True)
            
            self.log_message(f"Готово! Обработано: {processed_files_count} файлов ({modified_files_count} измененных), пропущено: {skipped_files_count}")
            return final_data
            
        except Exception as e:
            self.log_message(f"Ошибка: {str(e)}")
            self.log_message(traceback.format_exc())
            return pd.DataFrame()

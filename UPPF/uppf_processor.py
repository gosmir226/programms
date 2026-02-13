import pandas as pd
import os
import numpy as np
import re
import json
from datetime import datetime
import math
import sys

# Добавляем родительскую директорию в путь для импорта общих модулей
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from excel_manager import ExcelManager
    from cache_manager import CacheManager
except ImportError:
    # Фоллбек если пути не настроены (для тестов вне основной структуры)
    ExcelManager = None
    CacheManager = None

def extract_before_date(text):
    # Паттерны для поиска дат в различных форматах
    date_patterns = [
        r'\d{2}\.\d{2}\.\d{2}',
        r'\d{2}\.\d{2}\.\d{4}',
        r'\d{1,2}\.\d{1,2}\.\d{2,4}',
        r'\d{4}-\d{2}-\d{2}',
        r'\d{2}-\d{2}-\d{2}',
        r'\d{2}/\d{2}/\d{2}',
        r'\d{2}/\d{2}/\d{4}'
    ]
    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            date_str = match.group()
            try:
                for fmt in ['%d.%m.%y', '%d.%m.%Y', '%Y-%m-%d', '%d-%m-%y', '%d/%m/%y', '%d/%m/%Y']:
                    try:
                        datetime.strptime(date_str, fmt)
                        date_start = match.start()
                        before_date = text[:date_start].strip()
                        before_date = re.sub(r'[^a-zA-Zа-яА-Я0-9]+$', '', before_date)
                        return before_date.upper()
                    except ValueError:
                        continue
            except Exception:
                continue
    parts = re.split(r'[^a-zA-Zа-яА-Я0-9]+', text)
    if parts:
        return parts[0].upper()
    return ""

def find_fill(data):
    if 'ФОТОПИРОМЕТР' not in data.columns or data['ФОТОПИРОМЕТР'].isna().all():
        return None

    # Убеждаемся, что индексы уникальны
    data = data.reset_index(drop=True)
    max_index = data['ФОТОПИРОМЕТР'].idxmax()
    filtered = data.loc[max_index:].reset_index()

    for i in range(len(filtered) - 1):
        current_value = filtered.loc[i, 'ФОТОПИРОМЕТР']
        next_value = filtered.loc[i + 1, 'ФОТОПИРОМЕТР']

        # Пропускаем итерацию при наличии NaN в значениях
        if pd.isna(current_value) or pd.isna(next_value):
            continue

        # Условие А
        if abs(next_value - current_value) > 35:
            window = filtered.loc[i:i+5, 'ФОТОПИРОМЕТР']
            if any(window < 1200):
                candidate_index = filtered.loc[i, 'index']
                candidate_value = data.loc[candidate_index, 'ФОТОПИРОМЕТР']
                
                # Проверка диапазона значения
                if 1350 <= candidate_value <= 1700:
                    return candidate_index
                else:
                    return None

    return None

def calculate_holding_time_and_pour_time(series, max_value, fill_index):
    if fill_index is None:
        return None, None

    if max_value > 1620:
        lower_bound = 1600
        upper_bound = max_value
    else:
        lower_bound = max_value - 20
        upper_bound = max_value

    holding_indices = series[(series >= lower_bound) & (series <= upper_bound)].index
    if holding_indices.empty:
        return None, None
        
    holding_time = len(holding_indices) * 5
    last_holding_index = holding_indices[-1]
    time_to_pour = abs(fill_index - last_holding_index) * 5
    return holding_time, time_to_pour

def process_files(input_directory, output_file_path, log_callback):
    log_callback("Начало обработки...")
    log_callback(f"Входная директория: {input_directory}")
    log_callback(f"Выходной файл: {output_file_path}")
    
    # Инициализируем менеджеры
    if CacheManager:
        cache = CacheManager("UPPF", output_file_path)
        cache_info = cache.get_cache_info()
        log_callback(f"Загружен кэш, найдено {cache_info['file_count']} обработанных файлов")
    else:
        cache = None
        log_callback("ВНИМАНИЕ: CacheManager не загружен, кэширование отключено")
        
    if ExcelManager:
        excel = ExcelManager(output_file_path)
    else:
        excel = None
        log_callback("ВНИМАНИЕ: ExcelManager не загружен, умное сохранение отключено")
    
    all_rows = []
    
    # Проверяем существование директории
    if not os.path.exists(input_directory):
        log_callback(f"ОШИБКА: Директория {input_directory} не существует!")
        return

    file_count = 0
    processed_count = 0
    
    files_to_process = []
    for root, dirs, files in os.walk(input_directory):
        for filename in files:
            if filename.endswith(('.xlsx', '.xls')):
                files_to_process.append(os.path.join(root, filename))
    
    files_to_process.sort()
    total_files = len(files_to_process)
    log_callback(f"Найдено {total_files} excel файлов для проверки")
    
    for file_path in files_to_process:
        file_count += 1
        filename = os.path.basename(file_path)
        folder_name = os.path.basename(os.path.dirname(file_path))
        
        # Проверяем, был ли файл уже обработан и не изменился ли он
        if cache and not cache.is_file_changed(file_path):
            # log_callback(f"Файл {file_count}/{total_files}: {filename} актуален в кэше, пропускаем")
            continue
        
        log_callback(f"Обработка файла {file_count}/{total_files}: {filename}")
        
        try:
            engine = 'openpyxl' if filename.endswith('.xlsx') else 'xlrd'
            
            # --- ОТЧЕТ ---
            try:
                df_report = pd.read_excel(file_path, sheet_name='ОТЧЕТ', engine=engine)
                df_report = df_report.reset_index(drop=True)
            except Exception as e:
                log_callback(f"Файл {filename}: не удалось прочитать лист 'ОТЧЕТ': {e}")
                continue
            
            if df_report.empty:
                log_callback(f"Файл {filename}: лист 'ОТЧЕТ' пустой")
                continue
            
            # Нормализация названий столбцов
            df_report.columns = [str(col).strip().upper() for col in df_report.columns]
            
            # Проверка наличия столбца ПИРОМЕТР
            if 'ПИРОМЕТР' not in df_report.columns:
                log_callback(f"Файл {filename}: не найден столбец 'ПИРОМЕТР' на листе 'ОТЧЕТ'")
                continue
            
            # Обработка столбца ПИРОМЕТР
            pyro_col = 'ПИРОМЕТР'
            df_report[pyro_col] = (
                df_report[pyro_col]
                .astype(str)
                .str.replace(r'[^ -9.,-]', '', regex=True)
                .str.replace(',', '.')
                .replace('', np.nan)
            )
            df_report[pyro_col] = pd.to_numeric(df_report[pyro_col], errors='coerce')
            
            if df_report[pyro_col].isna().all():
                log_callback(f"Файл {filename}: в 'ПИРОМЕТР' нет числовых данных")
                continue
            
            # Поиск первого значимого изменения температуры (>1°C)
            temp_series = df_report[pyro_col]
            first_change_idx = None
            for i in range(1, len(temp_series)):
                current = temp_series.iloc[i]
                prev = temp_series.iloc[i-1]
                if pd.notna(current) and pd.notna(prev) and abs(current - prev) > 1:
                    first_change_idx = i
                    break
            
            if first_change_idx is None:
                log_callback(f"Файл {filename}: не найдено изменение температуры >1°C")
                continue
            
            # Поиск заливки
            def find_fill_wrapper(data):
                data = data.copy()
                data['ФОТОПИРОМЕТР'] = data[pyro_col]
                return find_fill(data)
            
            fill_index = find_fill_wrapper(df_report)
            
            # Вычисление времени
            full_holding_time = None
            fill_temperature = None
            if fill_index is not None:
                max_temp_index = df_report[pyro_col].idxmax()
                if fill_index >= max_temp_index:
                    full_holding_time = (fill_index - max_temp_index) * 5
                fill_temperature = df_report.loc[fill_index, pyro_col]
            
            # RT6 / PT6
            pt6_col = None
            for col in df_report.columns:
                if str(col).strip() in ['PT6', 'PT6 ']:
                    pt6_col = col
                    break
            
            minPT6, maxPT6 = None, None
            if pt6_col is not None and fill_index is not None:
                interval = df_report.loc[first_change_idx:fill_index]
                minPT6 = interval[pt6_col].min() if not interval[pt6_col].isnull().all() else None
                maxPT6 = interval[pt6_col].max() if not interval[pt6_col].isnull().all() else None

            # --- ПАСПОРТ ---
            passport_dict = {}
            furnace_type = None
            try:
                df_passport = pd.read_excel(file_path, sheet_name='ПАСПОРТ', engine=engine, header=None)
                if df_passport.shape[0] >= 2:
                    df_passport = df_passport.T
                    passport_headers = list(df_passport.iloc[0])
                    passport_values = list(df_passport.iloc[1])
                    temp_passport_dict = {str(k).strip().upper(): v for k, v in zip(passport_headers, passport_values)}
                    
                    volume = temp_passport_dict.get('ОБЪЕМ КАМЕРЫ:', '')
                    furnace_type = 'УППФ-У' if volume == '3700(Liter)' else 'УППФ-50'
                    passport_dict['ПЕЧЬ'] = furnace_type
                    
                    for key in ['НАЧАЛО ЗАПИСИ:', 'ЗАВЕРШЕНИЕ ЗАПИСИ:']:
                        if key in temp_passport_dict:
                            passport_dict[key] = temp_passport_dict[key]
                    
                    if furnace_type == 'УППФ-50':
                        for key in ['ДАТА ПРОВЕРКИ НАТЕКАНИЯ:', 'ВАКУУМ В НАЧАЛЕ ИЗМЕРЕНИЯ:', 'ВАКУУМ В КОНЦЕ ИЗМЕРЕНИЯ:', 'ВРЕМЯ ИЗМЕРЕНИЯ:', 'ОБЪЕМ КАМЕРЫ:', 'УРОВЕНЬ НАТЕКАНИЯ:', 'МИН. ВАКУУМ ВО ВРЕМЯ ПРОВЕРКИ:', 'МИН. ВАКУУМ ПЕРЕД ПРОВЕРКОЙ:']:
                            if key in temp_passport_dict:
                                passport_dict[key] = temp_passport_dict[key]
            except Exception as e:
                # log_callback(f"Файл {filename}: ошибка листа 'ПАСПОРТ': {e}")
                pass

            # --- ТЕМПЕРАТУРА ---
            temp_row_dict = {}
            max_overheat_temp = None
            photo_vals = None
            try:
                df_temp = pd.read_excel(file_path, sheet_name='ТЕМПЕРАТУРА', engine=engine)
                df_temp = df_temp.reset_index(drop=True)
                df_temp.columns = [str(col).strip().upper() for col in df_temp.columns]
                
                if 'ФОТОПИРОМЕТР' in df_temp.columns:
                    photo_vals = pd.to_numeric(df_temp['ФОТОПИРОМЕТР'].astype(str).str.replace(',', '.'), errors='coerce')
                    if not photo_vals.isnull().all():
                        max_overheat_temp = photo_vals.max()
                
                if fill_index is not None and fill_index < len(df_temp):
                    target_columns = ['ТППФ ВЕРХ (Т2)', 'ТППФ НИЗ (Т3)', 'ТППФ СР (Т4)']
                    for col in target_columns:
                        if col in df_temp.columns:
                            temp_row_dict[col] = df_temp.loc[fill_index, col]
                
                # holding_time и time_to_pour
                holding_time, time_to_pour = None, None
                if fill_index is not None and photo_vals is not None:
                    max_val = max_overheat_temp if max_overheat_temp is not None else photo_vals.max()
                    if max_val is not None:
                        holding_time, time_to_pour = calculate_holding_time_and_pour_time(photo_vals, max_val, fill_index)
            except Exception as e:
                # log_callback(f"Файл {filename}: ошибка листа 'ТЕМПЕРАТУРА': {e}")
                pass

            # --- Итоговая строка ---
            heat_name = extract_before_date(folder_name)
            row = {
                'Плавка': heat_name,
                'minPT6': minPT6,
                'maxPT6': maxPT6,
                'Температура перегрева': max_overheat_temp,
                'Температура заливки': fill_temperature,
                'Время выдержки при перегреве': holding_time if 'holding_time' in locals() else None,
                'Время от конца выдержки до заливки': time_to_pour if 'time_to_pour' in locals() else None,
                'Полное время выдержки': full_holding_time,
                'Название файла': filename,
                'Путь до файла excel': file_path
            }
            
            row.update(passport_dict)
            row.update(temp_row_dict)
            
            all_rows.append(row)
            processed_count += 1
            if cache:
                cache.update_file(file_path)
            
        except Exception as e:
            log_callback(f"Ошибка при обработке файла {filename}: {e}")

    if not all_rows and processed_count == 0:
        log_callback("Нет новых данных для обработки.")
        return

    # Формируем DataFrame
    final_result = pd.DataFrame(all_rows)
    
    # Добавляем гиперссылки
    if not final_result.empty and 'Название файла' in final_result.columns:
        final_result['Название файла'] = final_result.apply(
            lambda r: f'=HYPERLINK("{r["Путь до файла excel"]}", "{r["Название файла"]}")', axis=1)

    # Умное сохранение
    log_callback("Слияние новых данных с существующими...")
    
    # Читаем существующие
    if excel:
        existing_df, metadata = excel.read_excel_smart()
    else:
        existing_df = pd.DataFrame()

    if not existing_df.empty:
        combined = pd.concat([existing_df, final_result], ignore_index=True)
        if 'Путь до файла excel' in combined.columns:
            combined['dedup_key'] = combined['Путь до файла excel'].astype(str).str.lower().str.replace(r'[\\/]', r'\\', regex=True).str.strip()
            combined = combined.drop_duplicates(subset=['dedup_key'], keep='last').drop(columns=['dedup_key'])
    else:
        combined = final_result

    # Сортировка и расчеты на всем наборе
    if not combined.empty and 'НАЧАЛО ЗАПИСИ:' in combined.columns and 'Плавка' in combined.columns:
        combined['temp_sort_date'] = pd.to_datetime(combined['НАЧАЛО ЗАПИСИ:'], dayfirst=True, errors='coerce')
        combined = combined.sort_values(by=['Плавка', 'temp_sort_date'])
        combined['Порядок заливки'] = combined.groupby('Плавка').cumcount() + 1
        
        combined['Образец'] = ""
        for heat_name, group in combined.groupby('Плавка'):
            try:
                idx_max_order = group['Порядок заливки'].idxmax()
                combined.at[idx_max_order, 'Образец'] = "Предполагаемый образец"
                
                if 'Температура заливки' in group.columns and group['Температура заливки'].notna().any():
                    min_temp = group['Температура заливки'].min()
                    if min_temp < 1500:
                        min_temp_indices = group[group['Температура заливки'] == min_temp].index
                        for idx in min_temp_indices:
                            current = combined.at[idx, 'Образец']
                            label = "Минимальная температура заливки"
                            combined.at[idx, 'Образец'] = f"{current}, {label}" if current else label
            except Exception: pass
        
        combined = combined.drop(columns=['temp_sort_date'])

    # Записываем
    if excel:
        success = excel.write_excel_smart(
            combined, 
            key_columns=['Путь до файла excel'],
            log_callback=log_callback
        )
    else:
        combined.to_excel(output_file_path, index=False)
        success = True
    
    if success:
        log_callback(f"Успешно обработано: {processed_count} новых файлов")
    else:
        log_callback("Ошибка при сохранении Excel файла")

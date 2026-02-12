import pandas as pd
import os
import numpy as np
import re
import json
from datetime import datetime
import math

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
    if data['ФОТОПИРОМЕТР'].isna().all():
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

def determine_samples_by_temperature(files_data, log_callback):
    """
    Определение образцов по температуре заливки
    Если температура заливки < 1500 - это "Образец", иначе "Нет"
    """
    result = {}
    for item in files_data:
        filename = item['filename']
        fill_temp = item['fill_temperature']
        
        if fill_temp is not None and fill_temp < 1500:
            result[filename] = "Образец"
            log_callback(f"Файл {filename}: ОБРАЗЕЦ (температура заливки={fill_temp} < 1500)")
        else:
            result[filename] = "Нет"
            if fill_temp is None:
                log_callback(f"Файл {filename}: не образец (температура заливки отсутствует)")
            else:
                log_callback(f"Файл {filename}: не образец (температура заливки={fill_temp} >= 1500)")
    
    return result

def process_files(input_directory, output_file_path, log_callback):
    log_callback("Начало обработки...")
    log_callback(f"Входная директория: {input_directory}")
    log_callback(f"Выходной файл: {output_file_path}")
    
    # Определяем путь для кэш-файла
    output_dir = os.path.dirname(output_file_path)
    if not output_dir: # Handle case where output_file_path is just a filename
        output_dir = os.getcwd()
        
    cache_file_path = os.path.join(output_dir, 'casheUPPF.json')
    
    # Загружаем кэш, если он существует
    processed_files_cache = set()
    if os.path.exists(cache_file_path):
        try:
            with open(cache_file_path, 'r', encoding='utf-8') as f:
                cached_data = json.load(f)
                processed_files_cache = set(cached_data.get('processed_files', []))
            log_callback(f"Загружен кэш из {cache_file_path}, найдено {len(processed_files_cache)} обработанных файлов")
        except Exception as e:
            log_callback(f"Ошибка при загрузке кэша: {e}")
    
    all_passport_headers = set()
    all_temp_headers = set()
    all_rows = []
    
    # Проверяем существование директории
    if not os.path.exists(input_directory):
        log_callback(f"ОШИБКА: Директория {input_directory} не существует!")
        return

    file_count = 0
    processed_count = 0
    newly_processed_files = []
    
    for root, dirs, files in os.walk(input_directory):
        folder_name = os.path.basename(root)
        log_callback(f"Обработка папки: {folder_name}")
        
        # Сортируем excel файлы по возрастанию названия
        excel_files = [f for f in files if f.endswith(('.xlsx', '.xls'))]
        excel_files.sort()
        
        log_callback(f"Найдено {len(excel_files)} excel файлов в папке {folder_name}")
        
        for filename in excel_files:
            file_count += 1
            file_path = os.path.join(root, filename)
            
            # Проверяем, был ли файл уже обработан
            if filename in processed_files_cache:
                log_callback(f"Файл {file_count}: {filename} уже был обработан ранее, пропускаем")
                continue
            
            log_callback(f"Обработка файлa {file_count}: {filename}")
            
            try:
                engine = 'openpyxl' if filename.endswith('.xlsx') else 'xlrd'
                
                # --- ОТЧЕТ ---
                try:
                    df_report = pd.read_excel(file_path, sheet_name='ОТЧЕТ', engine=engine)
                    df_report = df_report.reset_index(drop=True)
                    log_callback(f"Файл {filename}: лист 'ОТЧЕТ' успешно прочитан")
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
                    log_callback(f"Файл {filename}: не найден столбец 'ПИРОМЕТР' на листе 'ОТЧЕТ'. Пропуск.")
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
                    log_callback(f"Файл {filename}: в столбце 'ПИРОМЕТР' нет числовых данных")
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
                    log_callback(f"Файл {filename}: не найдено изменение температуры >1°C. Пропуск.")
                    continue
                
                # --- find_fill с использованием столбца ПИРОМЕТР ---
                def find_fill_wrapper(data):
                    data = data.copy()
                    data['ФОТОПИРОМЕТР'] = data[pyro_col]
                    return find_fill(data)
                
                fill_index = find_fill_wrapper(df_report)
                log_callback(f"Файл {filename}: fill_index={fill_index}")
                
                # --- Вычисление полного времени выдержки ---
                full_holding_time = None
                fill_temperature = None
                
                if fill_index is not None:
                    max_temp_index = df_report[pyro_col].idxmax()
                    if fill_index >= max_temp_index:
                        full_holding_time = (fill_index - max_temp_index) * 5
                    fill_temperature = df_report.loc[fill_index, pyro_col]
                
                # --- PT6 ---
                pt6_col = None
                for col in df_report.columns:
                    if str(col).strip() == 'PT6':
                        pt6_col = col
                        break
                    if str(col).strip() == 'PT6 ':
                        pt6_col = col
                        break
                if pt6_col is None:
                    log_callback(f"Файл {filename}: не найден столбец 'PT6'. Пропуск.")
                    continue
                
                minPT6, maxPT6 = None, None
                if fill_index is not None:
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
                        
                        # Создаем словарь с заголовками в верхнем регистре
                        temp_passport_dict = {str(k).strip().upper(): v for k, v in zip(passport_headers, passport_values)}
                        
                        # Определяем тип печи на основе объема камеры
                        volume = temp_passport_dict.get('ОБЪЕМ КАМЕРЫ:', '')
                        if volume == '3700(Liter)':
                            furnace_type = 'УППФ-У'
                        else:
                            furnace_type = 'УППФ-50'
                        
                        # Добавляем тип печи в словарь
                        passport_dict['ПЕЧЬ'] = furnace_type
                        
                        # Всегда добавляем эти столбцы
                        for key in ['НАЧАЛО ЗАПИСИ:', 'ЗАВЕРШЕНИЕ ЗАПИСИ:']:
                            if key in temp_passport_dict:
                                passport_dict[key] = temp_passport_dict[key]
                        
                        # Для УППФ-50 добавляем дополнительные столбцы
                        if furnace_type == 'УППФ-50':
                            additional_keys = [
                                'ДАТА ПРОВЕРКИ НАТЕКАНИЯ:',
                                'ВАКУУМ В НАЧАЛЕ ИЗМЕРЕНИЯ:',
                                'ВАКУУМ В КОНЦЕ ИЗМЕРЕНИЯ:',
                                'ВРЕМЯ ИЗМЕРЕНИЯ:',
                                'ОБЪЕМ КАМЕРЫ:',
                                'УРОВЕНЬ НАТЕКАНИЯ:',
                                'МИН. ВАКУУМ ВО ВРЕМЯ ПРОВЕРКИ:',
                                'МИН. ВАКУУМ ПЕРЕД ПРОВЕРКОЙ:'
                            ]
                            for key in additional_keys:
                                if key in temp_passport_dict:
                                    passport_dict[key] = temp_passport_dict[key]
                        
                        all_passport_headers.update(passport_dict.keys())
                        log_callback(f"Файл {filename}: лист 'ПАСПОРТ' успешно обработан, тип печи: {furnace_type}")
                except Exception as e:
                    log_callback(f"Файл {filename}: не удалось прочитать лист 'ПАСПОРТ': {e}")

                # --- ТЕМПЕРАТУРА ---
                temp_row_dict = {}
                max_overheat_temp = None
                try:
                    df_temp = pd.read_excel(file_path, sheet_name='ТЕМПЕРАТУРА', engine=engine)
                    df_temp = df_temp.reset_index(drop=True)
                    
                    # Нормализуем названия столбцов к верхнему регистру
                    df_temp.columns = [str(col).strip().upper() for col in df_temp.columns]
                    
                    # Обработка ФОТОПИРОМЕТР для расчетов (но не для вывода)
                    if 'ФОТОПИРОМЕТР' in df_temp.columns:
                        df_temp['ФОТОПИРОМЕТР'] = df_temp['ФОТОПИРОМЕТР'].astype(str).replace(r'[^ -9.,-]', '', regex=True)
                        df_temp['ФОТОПИРОМЕТР'] = df_temp['ФОТОПИРОМЕТР'].str.replace(',', '.')
                        df_temp['ФОТОПИРОМЕТР'] = pd.to_numeric(df_temp['ФОТОПИРОМЕТР'], errors='coerce')
                        
                        if not df_temp['ФОТОПИРОМЕТР'].isnull().all():
                            max_overheat_temp = df_temp['ФОТОПИРОМЕТР'].max()
                    
                    # Добавляем индекс для поиска
                    df_temp['Index'] = df_temp.index
                    
                    # Берем только нужные столбцы для вывода
                    target_columns = ['ТППФ ВЕРХ (Т2)', 'ТППФ НИЗ (Т3)', 'ТППФ СР (Т4)']
                    
                    if fill_index in df_temp['Index'].values:
                        temp_row = df_temp.loc[df_temp['Index'] == fill_index].iloc[0]
                        # Берем только указанные столбцы
                        for col in target_columns:
                            if col in df_temp.columns:
                                temp_row_dict[col] = temp_row[col]
                        
                        all_temp_headers.update(target_columns)
                        
                    log_callback(f"Файл {filename}: лист 'ТЕМПЕРАТУРА' успешно обработан, max_overheat_temp={max_overheat_temp}")
                except Exception as e:
                    log_callback(f"Файл {filename}: не удалось прочитать лист 'ТЕМПЕРАТУРА': {e}")

                # --- holding_time и time_to_pour ---
                holding_time, time_to_pour = None, None
                try:
                    if 'df_temp' in locals() and fill_index is not None:
                        if 'max_overheat_temp' in locals() and max_overheat_temp is not None:
                            max_value = max_overheat_temp
                        elif 'ФОТОПИРОМЕТР' in locals() and not df_temp['ФОТОПИРОМЕТР'].isnull().all():
                            max_value = df_temp['ФОТОПИРОМЕТР'].max()
                        else:
                            max_value = None
                            
                        if max_value is not None:
                            holding_time, time_to_pour = calculate_holding_time_and_pour_time(
                                df_temp['ФОТОПИРОМЕТР'],
                                max_value,
                                fill_index
                            )
                except Exception as e:
                    log_callback(f"Ошибка расчета времени для {filename}: {e}")

                # --- Итоговая строка ---
                heat_name = extract_before_date(folder_name)
                row = {
                    'Плавка': heat_name,
                    'minPT6': minPT6,
                    'maxPT6': maxPT6,
                    'Температура перегрева': max_overheat_temp,
                    'Температура заливки': fill_temperature,
                    'Время выдержки при перегреве': holding_time,
                    'Время от конца выдержки до заливки': time_to_pour,
                    'Полное время выдержки': full_holding_time,
                    'Название файла': filename,
                    'Путь до файла excel': file_path
                }
                
                # Добавляем данные из паспорта и температуры
                for source_dict in [passport_dict, temp_row_dict]:
                    for key, value in source_dict.items():
                        if key not in row:
                            row[key] = value
                
                all_rows.append(row)
                
                processed_count += 1
                newly_processed_files.append(filename)
                log_callback(f"Файл {filename}: успешно обработан")
                

                    
            except Exception as e:
                log_callback(f"Ошибка при обработке файла {filename}: {e}")
                import traceback
                log_callback(f"Подробности ошибки: {traceback.format_exc()}")

    log_callback(f"Обработка завершена. Обработано {processed_count} из {file_count} файлов")
    
    # Обновляем кэш
    if newly_processed_files:
        processed_files_cache.update(newly_processed_files)
        cache_data = {
            'processed_files': list(processed_files_cache),
            'last_update': datetime.now().isoformat(),
            'total_processed': len(processed_files_cache)
        }
        try:
            with open(cache_file_path, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            log_callback(f"Ошибка при сохранении кэша: {e}")
    
    if processed_count == 0:
        log_callback("ВНИМАНИЕ: Ни один файл не был обработан! Проверьте логи для выяснения причин.")
        return

    # Формируем итоговую таблицу
    # Сначала собираем все заголовки из всех строк
    all_headers = set()
    for row in all_rows:
        all_headers.update(row.keys())
    
    # Создаем упорядоченный список колонок, где обязательные поля идут первыми
    mandatory_columns = [
        'Плавка', 'Порядок заливки', 'Образец', 'Температура перегрева', 'Температура заливки',
        'Время выдержки при перегреве', 'Время от конца выдержки до заливки',
        'ТППФ ВЕРХ (Т2)', 'ТППФ НИЗ (Т3)', 'ТППФ СР (Т4)',
        'minPT6', 'maxPT6'
    ]
    
    # Удаляем обязательные колонки из общего набора, чтобы не дублировать
    for col in mandatory_columns:
        if col in all_headers:
            all_headers.remove(col)
    
    # Формируем окончательный список колонок
    columns = mandatory_columns + sorted(all_headers)
    
    # Создаем DataFrame
    final_result = pd.DataFrame(all_rows, columns=columns)
    
    # Добавляем гиперссылки
    if not final_result.empty:
        final_result['Название файла'] = final_result.apply(
            lambda row: f'=HYPERLINK("{row["Путь до файла excel"]}", "{row["Название файла"]}")', axis=1)
    
    try:
        if os.path.exists(output_file_path):
            existing_data = pd.read_excel(output_file_path)
            existing_data = existing_data.reset_index(drop=True)
            log_callback(f"Обнаружен существующий файл, загружено {len(existing_data)} строк")
            combined_data = pd.concat([existing_data, final_result], ignore_index=True)
            before_deduplication = len(combined_data)

            # Формируем ключ для дедупликации (нормализованный путь)
            dedup_col = 'dedup_key'
            if 'Путь до файла excel' in combined_data.columns:
                combined_data[dedup_col] = combined_data['Путь до файла excel'].astype(str).str.lower().str.replace(r'[\\/]', r'\\', regex=True).str.strip()
            else:
                # Если вдруг нет пути - используем название (с гиперссылкой)
                combined_data[dedup_col] = combined_data['Название файла'].astype(str).str.strip()
            
            combined_data = combined_data.drop_duplicates(subset=[dedup_col], keep='last')
            combined_data = combined_data.drop(columns=[dedup_col])
            
            after_deduplication = len(combined_data)
            log_callback(f"После объединения: {before_deduplication} строк, после удаления дубликатов: {after_deduplication} строк")
            
            df_to_process = combined_data
        else:
            df_to_process = final_result
            log_callback(f"Создан новый файл с результатами: {output_file_path}")

        # --- Сортировка и добавление 'Порядок заливки' ---
        if 'НАЧАЛО ЗАПИСИ:' in df_to_process.columns and 'Плавка' in df_to_process.columns:
            # Временный столбец для даты
            df_to_process['temp_sort_date'] = pd.to_datetime(
                df_to_process['НАЧАЛО ЗАПИСИ:'], 
                dayfirst=True, 
                errors='coerce'
            )
            
            # Сортировка
            df_to_process = df_to_process.sort_values(by=['Плавка', 'temp_sort_date'])
            
            # Вычисление порядка
            df_to_process['Порядок заливки'] = df_to_process.groupby('Плавка').cumcount() + 1
            
            # Удаление временного столбца
            df_to_process = df_to_process.drop(columns=['temp_sort_date'])
            log_callback("Выполнена хронологическая сортировка и индексация по плавке")

            # --- Расчет столбца "Образец" ---
            df_to_process['Образец'] = ""
            # Группируем по плавке
            for heat_name, group in df_to_process.groupby('Плавка'):
                try:
                    # 1. Последний по времени (максимальный порядок заливки)
                    if 'Порядок заливки' in group.columns and not group.empty:
                        idx_max_order = group['Порядок заливки'].idxmax()
                        if pd.notna(idx_max_order):
                            df_to_process.at[idx_max_order, 'Образец'] = "Предполагаемый образец"
                    
                    # 2. Минимальная температура заливки
                    if 'Температура заливки' in group.columns and group['Температура заливки'].notna().any():
                        min_temp = group['Температура заливки'].min()
                        # Дополнительная проверка на < 1500
                        if min_temp < 1500:
                            # Находим все индексы с этой минимальной температурой
                            min_temp_indices = group[group['Температура заливки'] == min_temp].index
                            for idx in min_temp_indices:
                                current_val = df_to_process.at[idx, 'Образец']
                                label = "Минимальная температура заливки"
                                if current_val and current_val != label:
                                    df_to_process.at[idx, 'Образец'] = f"{current_val}, {label}"
                                else:
                                    df_to_process.at[idx, 'Образец'] = label
                except Exception as e_sample:
                    log_callback(f"Ошибка при расчете образца для плавки {heat_name}: {e_sample}")
            
            log_callback("Столбец 'Образец' заполнен")

        df_to_process.to_excel(output_file_path, index=False)
        log_callback(f"Данные сохранены в файл: {output_file_path}")
        log_callback(f"Всего строк в результате: {len(df_to_process)}")

    except Exception as e:
        log_callback(f"Ошибка при сохранении результата: {e}")

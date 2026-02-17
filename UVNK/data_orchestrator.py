import os
import sys
from nptdms import TdmsFile
import pandas as pd
import re
import report_builder as rb
import core_algorithms as ca
import logging
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# Добавляем родительскую директорию в путь для импорта общих модулей
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from excel_manager import ExcelManager
    from cache_manager import CacheManager
except ImportError:
    ExcelManager = None
    CacheManager = None

logging.getLogger("nptdms.reader").setLevel(logging.ERROR)

FILENAME_PATTERN = re.compile(r'^\d+\.tdms$')

def deduplicate_columns(df):
    """
    Переименовывает дублирующиеся столбцы, добавляя суффиксы.
    Это КРИТИЧНО для создания Excel Table, которая требует уникальных заголовков.
    """
    cols = list(df.columns)
    seen = {}
    new_cols = []
    for col in cols:
        if col in seen:
            seen[col] += 1
            new_name = f"{col}_{seen[col]}"
            new_cols.append(new_name)
        else:
            seen[col] = 0
            new_cols.append(col)
    df.columns = new_cols
    return df

def read_tdms_file(file_path):
    """Читает .tdms файл и возвращает DataFrame. При ошибке возвращает None."""
    try:
        tdms_file = TdmsFile.read(file_path)
        all_groups = tdms_file.groups()
        if not all_groups:
            return None

        dataframes = []
        for group in all_groups:
            channels = group.channels()
            if not channels:
                continue

            data_dict = {}
            for ch in channels:
                arr = ch[:]
                data_dict[ch.name] = arr
            df_group = pd.DataFrame(data_dict)
            dataframes.append(df_group)

        if dataframes:
            merged = pd.concat(dataframes, axis=1) if len(dataframes) > 1 else dataframes[0]
            return merged
        else:
            return None

    except Exception as e:
        print(f"Ошибка чтения файла {file_path}: {e}")
        return None

def merge_dataframes(df_reports, df_pasport):
    """
    Объединяет данные из Reports и Pasport.
    Переносит все уникальные колонки из Pasport (например, Дату) в Reports.
    Если есть NumberOfM, приоритет отдается данным из Pasport.
    """
    if df_reports is None or len(df_reports) == 0:
        return None
    
    df_reports_copy = df_reports.copy()
    
    if df_pasport is None or df_pasport.empty:
        return df_reports_copy

    # 1. Обрабатываем NumberOfM (если есть в паспорте)
    if 'NumberOfM' in df_pasport.columns and not df_pasport['NumberOfM'].empty:
        df_reports_copy['NumberOfM'] = df_pasport['NumberOfM'].iloc[0]
    
    # 2. Переносим все остальные уникальные колонки из Pasport
    for col in df_pasport.columns:
        # Если колонки нет ИЛИ она состоит из NaN (пустая заготовка)
        if col not in df_reports_copy.columns or df_reports_copy[col].isna().all():
            if not df_pasport[col].empty:
                # Распространяем первое значение на все строки (метаданные)
                df_reports_copy[col] = df_pasport[col].iloc[0]

    return df_reports_copy



def process_session_to_excel(root_folder, output_excel_path, log_callback):
    """
    Сканирует корневую папку на наличие подпапок (установок).
    В каждой подпапке ищет Pasport/Reports, обрабатывает их и собирает общую таблицу.
    Использует CacheManager для отслеживания изменений и ExcelManager для надежного обновления Excel.
    """
    if not os.path.exists(root_folder):
        log_callback(f"Ошибка: Корневая папка не существует: {root_folder}")
        return

    # Инициализируем кэш и Excel manager
    if CacheManager:
        cache = CacheManager('UVNK', output_excel_path)
    else:
        cache = None
        log_callback("ВНИМАНИЕ: CacheManager не загружен")
        
    if ExcelManager:
        excel_manager = ExcelManager(output_excel_path)
    else:
        excel_manager = None
        log_callback("ВНИМАНИЕ: ExcelManager не загружен")
    
    # Ищем все подпапки в корневой директории
    subfolders = [
        f for f in os.listdir(root_folder)
        if os.path.isdir(os.path.join(root_folder, f))
    ]

    if not subfolders:
        log_callback("В указанной папке нет вложенных папок.")
        return

    FILENAME_PATTERN = re.compile(r'^(\d+)\.tdms$')
    
    # Собираем список всех tdms файлов
    all_tdms_files = []
    for installation_name in subfolders:
        installation_folder = os.path.join(root_folder, installation_name)
        reports_folder = os.path.join(installation_folder, "Reports")
        
        if os.path.isdir(reports_folder):
            for filename in os.listdir(reports_folder):
                if FILENAME_PATTERN.match(filename):
                    file_path = os.path.join(reports_folder, filename)
                    all_tdms_files.append((file_path, installation_name, filename))
    
    # Определяем, какие файлы нужно перечитать
    if cache:
        stale_file_paths = []
        for f_path, _, _ in all_tdms_files:
            if cache.is_file_changed(f_path):
                stale_file_paths.append(os.path.abspath(f_path))
    else:
        stale_file_paths = [os.path.abspath(f[0]) for f in all_tdms_files]
    
    # Собираем источники (Установки), которые требуют обновления
    modified_installations = set()
    for file_path, installation_name, filename in all_tdms_files:
        if os.path.abspath(file_path) in stale_file_paths:
            modified_installations.add(installation_name)
    
    if not modified_installations:
        log_callback("Все файлы уже обработаны и не изменились. Обработка не требуется.")
        return

    all_result_rows = []
    execution_logs = []
    total_new_rows_count = 0

    log_callback(f"Найдено установок (папок): {len(subfolders)}")
    log_callback(f"Установок требующих обновления: {len(modified_installations)}")

    for installation_name in subfolders:
        # Обрабатываем только установки с измененными файлами
        if installation_name not in modified_installations:
            log_callback(f"\n--- Установка {installation_name}: все файлы актуальны, пропуск ---")
            continue
            
        installation_folder = os.path.join(root_folder, installation_name)
        log_callback(f"\n--- Обработка установки: {installation_name} ---")

        pasport_folder = os.path.join(installation_folder, "Pasport")
        reports_folder = os.path.join(installation_folder, "Reports")

        if not (os.path.isdir(pasport_folder) and os.path.isdir(reports_folder)):
            log_callback(f"Пропуск {installation_name}: нет папок Pasport/Reports")
            execution_logs.append({
                'Дата': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'Установка': installation_name,
                'Файл': '-',
                'Статус': 'Пропущено',
                'Причина': 'Нет папок Pasport/Reports'
            })
            continue

        def get_valid_files(folder):
            return [f for f in os.listdir(folder) if FILENAME_PATTERN.match(f)]

        pasport_files = set(get_valid_files(pasport_folder))
        reports_files = set(get_valid_files(reports_folder))
        
        common_files = sorted(list(pasport_files & reports_files), key=lambda x: os.path.getmtime(os.path.join(reports_folder, x)))

        if not common_files:
            log_callback(f"В {installation_name} нет общих файлов.")
            continue

        current_leakage_info = None
        
        for filename in common_files:
            pasport_path = os.path.join(pasport_folder, filename)
            reports_path = os.path.join(reports_folder, filename)
            abs_reports_path = os.path.abspath(reports_path)
            
            # Проверяем кэш
            if cache and not cache.is_file_changed(abs_reports_path):
                # Но нам нужно натекание!
                df_reports = read_tdms_file(reports_path)
                if df_reports is not None and not df_reports.empty:
                    df_pasport = read_tdms_file(pasport_path)
                    merged_df = None
                    if df_pasport is not None and not df_pasport.empty:
                        try: 
                            merged_df = merge_dataframes(df_reports, df_pasport)
                        except Exception as e:
                            log_callback(f"Предупреждение: Не удалось объединить {filename} с паспортом: {e}")
                            merged_df = None
                    if merged_df is None:
                        merged_df = df_reports.copy()
                        merged_df['NumberOfM'] = os.path.splitext(filename)[0]
                    
                    leakage = ca.get_last_valid_leakage(merged_df)
                    if leakage:
                        current_leakage_info = leakage
                        current_leakage_info['source'] = filename
                
                execution_logs.append({
                    'Дата': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'Установка': installation_name,
                    'Файл': filename,
                    'Статус': 'Пропущено',
                    'Причина': 'Уже в кэше'
                })
                continue

            # Обработка нового/измененного файла
            df_reports = read_tdms_file(reports_path)
            if df_reports is None or df_reports.empty:
                continue

            df_pasport = read_tdms_file(pasport_path)
            if df_pasport is not None and not df_pasport.empty:
                try: 
                    merged_df = merge_dataframes(df_reports, df_pasport)
                except Exception as e:
                    log_callback(f"Ошибка объединения {filename} с паспортом: {e}")
                    merged_df = None
            else:
                log_callback(f"Предупреждение: Файл паспорта для {filename} пуст или не прочитан.")
            
            if merged_df is None:
                log_callback(f"Используются только данные из Reports для {filename}")
                merged_df = df_reports.copy()
                merged_df['NumberOfM'] = os.path.splitext(filename)[0]

            try:
                leakage = ca.get_last_valid_leakage(merged_df)
                if leakage:
                    current_leakage_info = leakage
                    current_leakage_info['source'] = filename

                merged_df = deduplicate_columns(merged_df)
                rows = rb.process_dataframe_segments(merged_df, current_leakage_info)
                if rows:
                    for row in rows: row['Установка'] = installation_name
                    all_result_rows.extend(rows)
                    total_new_rows_count += len(rows)
                    if cache: cache.update_file(abs_reports_path)
                    execution_logs.append({
                        'Дата': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'Установка': installation_name,
                        'Файл': filename,
                        'Статус': 'Успешно',
                        'Причина': f'Обработано {len(rows)} сегментов'
                    })
                else:
                    if cache: cache.update_file(abs_reports_path)
                    execution_logs.append({
                        'Дата': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'Установка': installation_name,
                        'Файл': filename,
                        'Статус': 'Пропущено',
                        'Причина': 'Нет валидных сегментов'
                    })
            except Exception as e:
                log_callback(f"Ошибка {installation_name}/{filename}: {e}")

    # Сохранение результатов
    if all_result_rows or execution_logs:
        if excel_manager:
            result_df = pd.DataFrame(all_result_rows)
            if not result_df.empty:
                # Установка первая
                cols = ['Установка'] + [c for c in result_df.columns if c != 'Установка']
                result_df = result_df[cols]
                result_df = deduplicate_columns(result_df)
                
                excel_manager.write_excel_smart(
                    result_df, 
                    key_columns=['Установка', 'NumberOfM'], # Уникально идентифицирует замер в рамках установки
                    sheet_name='Data',
                    log_callback=log_callback
                )
            
            logs_df = pd.DataFrame(execution_logs)
            if not logs_df.empty:
                logs_df = deduplicate_columns(logs_df)
                excel_manager.write_excel_smart(
                    logs_df,
                    key_columns=['Дата', 'Установка', 'Файл'],
                    sheet_name='Logs',
                    log_callback=log_callback,
                    mode='append' # Логи всегда добавляем
                )
        else:
            # Fallback
            pd.DataFrame(all_result_rows).to_excel(output_excel_path, index=False)

        log_callback(f"Обработка завершена. Добавлено {total_new_rows_count} строк.")
    else:
        log_callback("Нет новых данных для сохранения.")

def clear_cache_for_output(output_excel_path):
    if CacheManager:
        cache = CacheManager('UVNK', output_excel_path)
        cache.clear_cache()

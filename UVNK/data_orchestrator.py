import os
from nptdms import TdmsFile
import pandas as pd
import re
import report_builder as rb
import core_algorithms as ca
import json
import logging
from datetime import datetime
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

logging.getLogger("nptdms.reader").setLevel(logging.ERROR)

FILENAME_PATTERN = re.compile(r'^\d+\.tdms$')
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(BASE_DIR, "cache.json")

def load_cache():
    """Загружает кэш из файла или возвращает пустой словарь"""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            print(f"[КЭШ] Файл {CACHE_FILE} повреждён или пуст. Создаётся новый.")
    return {}

def save_cache(cache):
    """Сохраняет кэш в файл"""
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"[КЭШ] Ошибка при сохранении кэша: {e}")

def deduplicate_columns(df):
    """
    Renames duplicate columns (or empty ones) to ensure every column has a unique string name.
    Useful for creating Excel Tables which require unique non-empty headers.
    """
    new_columns = []
    seen = set()
    for col in df.columns:
        col_str = str(col).strip()
        if not col_str:
            col_str = "Unnamed"
        
        if col_str in seen:
            i = 1
            while f"{col_str}_{i}" in seen:
                i += 1
            col_str = f"{col_str}_{i}"
        
        seen.add(col_str)
        new_columns.append(col_str)
    
    df.columns = new_columns
    return df

def read_tdms_file(path):
    try:
        tdms_file = TdmsFile.read(path)
        data = {}
        for group in tdms_file.groups():
            for channel in group.channels():
                data[channel.name] = channel[:]
        
        # Try creating DataFrame directly (fastest, requires equal lengths)
        try:
            return pd.DataFrame(data)
        except ValueError:
            # Fallback: create using Series to support channels of different lengths
            return pd.DataFrame({k: pd.Series(v) for k, v in data.items()})
            
    except Exception as e:
        print(f"Ошибка при чтении файла {path}: {e}")
        return None

def merge_dataframes(df1, df2):
    if df2.empty:
        raise ValueError("❌ Второй DataFrame пустой.")

    first_row_df2 = df2.iloc[0].to_dict()

    repeated_values = {col: [first_row_df2[col]] * len(df1) for col in df2.columns}

    df2_repeated = pd.DataFrame(repeated_values)

    merged_df = pd.concat([df1.reset_index(drop=True), df2_repeated.reset_index(drop=True)], axis=1)

    return merged_df

def process_session_to_excel(root_folder, output_excel_path, log_callback):
    """
    Сканирует корневую папку на наличие подпапок (установок).
    В каждой подпапке ищет Pasport/Reports, обрабатывает их и собирает общую таблицу.
    """
    if not os.path.exists(root_folder):
        log_callback(f"Ошибка: Корневая папка не существует: {root_folder}")
        return

    # Ищем все подпапки в корневой директории
    subfolders = [
        f for f in os.listdir(root_folder)
        if os.path.isdir(os.path.join(root_folder, f))
    ]

    if not subfolders:
        log_callback("В указанной папке нет вложенных папок.")
        return

    FILENAME_PATTERN = re.compile(r'^(\d+)\.tdms$')
    
    # Загружаем кэш (он один на весь выходной файл)
    cache = load_cache()
    output_excel_key = os.path.abspath(output_excel_path)
    processed_files_set = set(cache.get(output_excel_key, []))

    all_result_rows = []
    newly_processed_files_batch = []
    execution_logs = []  # Список для логов выполнения
    
    total_new_rows_count = 0

    log_callback(f"Найдено установок (папок): {len(subfolders)}")

    for installation_name in subfolders:
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
            return [
                f for f in os.listdir(folder)
                if FILENAME_PATTERN.match(f)
            ]

        pasport_files = set(get_valid_files(pasport_folder))
        reports_files = set(get_valid_files(reports_folder))
        
        # Файлы, которые есть в Reports, но нет в Pasport
        missing_in_pasport = reports_files - pasport_files
        for missing_file in missing_in_pasport:
            execution_logs.append({
                'Дата': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'Установка': installation_name,
                'Файл': missing_file,
                'Статус': 'Пропущено',
                'Причина': 'Отсутствует файл в папке Pasport'
            })
            
        # Файлы, которые есть в Pasport, но нет в Reports
        missing_in_reports = pasport_files - reports_files
        for missing_file in missing_in_reports:
             execution_logs.append({
                'Дата': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'Установка': installation_name,
                'Файл': missing_file,
                'Статус': 'Пропущено',
                'Причина': 'Отсутствует файл в папке Reports'
            })

        def get_mod_time(filename):
            path = os.path.join(reports_folder, filename)
            if os.path.exists(path):
                return os.path.getmtime(path)
            return 0

        common_files = sorted(list(pasport_files & reports_files), key=get_mod_time)

        if not common_files:
            log_callback(f"В {installation_name} нет общих файлов.")
            continue

        # Локальное состояние натекания для этой установки
        current_leakage_info = None
        
        for filename in common_files:
            # Ключ для кэша теперь включает имя установки: UVNK1/123.tdms
            cache_key = f"{installation_name}/{filename}"
            is_new = cache_key not in processed_files_set

            pasport_path = os.path.join(pasport_folder, filename)
            reports_path = os.path.join(reports_folder, filename)

            df_reports = read_tdms_file(reports_path)
            if df_reports is None or df_reports.empty:
                log_callback(f"[{installation_name}] Файл {filename}: Reports не прочитан или пуст. Пропуск.")
                execution_logs.append({
                    'Дата': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'Установка': installation_name,
                    'Файл': filename,
                    'Статус': 'Ошибка',
                    'Причина': 'Reports не прочитан или пуст'
                })
                continue

            df_pasport = read_tdms_file(pasport_path)
            merged_df = None
            
            # Пытаемся объединить, если паспорт прочитался
            if df_pasport is not None and not df_pasport.empty:
                try:
                    merged_df = merge_dataframes(df_reports, df_pasport)
                except Exception as e:
                    log_callback(f"[{installation_name}] {filename}: Ошибка объединения с паспортом ({e}). Обрабатываем только Reports.")
                    merged_df = None
            else:
                log_callback(f"[{installation_name}] {filename}: Паспорт пуст или не прочитан. Обрабатываем только Reports.")

            # Если объединение не удалось — работаем только с Reports
            if merged_df is None:
                merged_df = df_reports.copy()
                # Добавляем имя файла как NumberOfM (без расширения)
                file_name_no_ext = os.path.splitext(filename)[0]
                merged_df['NumberOfM'] = file_name_no_ext

            # Далее обработка (общая для обоих случаев)
            try:
                # Обновляем инфо о натекании из текущего файла
                leakage = ca.get_last_valid_leakage(merged_df)
                if leakage:
                    current_leakage_info = leakage
                    current_leakage_info['source'] = filename
                    log_callback(f"[{installation_name}] {filename}: найдено новое натекание (source={filename}).")

                if is_new:
                    # Перед обработкой нужно убедиться, что имена столбцов уникальны, чтобы не было конфликтов
                    merged_df = deduplicate_columns(merged_df)
                    
                    rows = rb.process_dataframe_segments(merged_df, current_leakage_info)
                    if rows:
                        # Добавляем столбец Установка
                        for row in rows:
                            row['Установка'] = installation_name
                        
                        all_result_rows.extend(rows)
                        total_new_rows_count += len(rows)
                        log_callback(f"[{installation_name}] {filename}: обработано {len(rows)} сегментов.")
                        execution_logs.append({
                            'Дата': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            'Установка': installation_name,
                            'Файл': filename,
                            'Статус': 'Успешно',
                            'Причина': f'Обработано {len(rows)} сегментов'
                        })
                    else:
                        log_callback(f"[{installation_name}] {filename}: не найдено валидных сегментов.")
                        execution_logs.append({
                            'Дата': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            'Установка': installation_name,
                            'Файл': filename,
                            'Статус': 'Пропущено',
                            'Причина': 'Не найдено валидных сегментов'
                        })
                    
                    newly_processed_files_batch.append(cache_key)
                else:
                    # Файл уже был обработан
                    execution_logs.append({
                        'Дата': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'Установка': installation_name,
                        'Файл': filename,
                        'Статус': 'Пропущено',
                        'Причина': 'Уже обработан ранее (в кэше)'
                    })
                    pass

            except Exception as e:
                log_callback(f"Ошибка при обработке {installation_name}/{filename}: {e}")
                execution_logs.append({
                    'Дата': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'Установка': installation_name,
                    'Файл': filename,
                    'Статус': 'Ошибка',
                    'Причина': f'Исключение: {str(e)}'
                })

    # Сохранение результатов
    if all_result_rows or execution_logs:
        result_df = pd.DataFrame(all_result_rows) 
        logs_df = pd.DataFrame(execution_logs)

        # Вывод result_df - Упорядочиваем столбцы - Установка первая
        if not result_df.empty:
            cols = ['Установка'] + [c for c in result_df.columns if c != 'Установка']
            result_df = result_df[cols]

        combined_data_df = pd.DataFrame()
        combined_logs_df = pd.DataFrame()

        if os.path.exists(output_excel_path):
            try:
                # Читаем все листы
                xls = pd.ExcelFile(output_excel_path)
                
                # Sheet с данными (обычно первый или 'Data')
                sheet_names = xls.sheet_names
                data_sheet = 'Data' if 'Data' in sheet_names else sheet_names[0]
                logs_sheet = 'Logs' if 'Logs' in sheet_names else None
                
                existing_data = pd.read_excel(output_excel_path, sheet_name=data_sheet)
                
                if not result_df.empty:
                    combined_data_df = pd.concat([existing_data, result_df], ignore_index=True)
                else:
                    combined_data_df = existing_data
                
                if logs_sheet:
                    existing_logs = pd.read_excel(output_excel_path, sheet_name=logs_sheet)
                    combined_logs_df = pd.concat([existing_logs, logs_df], ignore_index=True)
                else:
                    combined_logs_df = logs_df

            except Exception as e:
                log_callback(f"Ошибка при чтении существующего файла Excel: {e}")
                return
        else:
            combined_data_df = result_df
            combined_logs_df = logs_df
            
        # -- ВАЖНО: Дедупликация колонок перед записью, иначе Excel-таблица сломается --
        if not combined_data_df.empty:
            combined_data_df = deduplicate_columns(combined_data_df)
        if not combined_logs_df.empty:
            combined_logs_df = deduplicate_columns(combined_logs_df)

        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                combined_data_df.to_excel(writer, sheet_name='Data', index=False)
                combined_logs_df.to_excel(writer, sheet_name='Logs', index=False)
                
                # Форматирование как "Умная таблица" (Excel Table)
                # Требует уникальных заголовков (обеспечено deduplicate_columns)
                workbook = writer.book
                
                # --- Лист Data ---
                if not combined_data_df.empty:
                    ws_data = writer.sheets['Data']
                    last_row = len(combined_data_df) + 1  # +1 для заголовка
                    last_col = len(combined_data_df.columns)
                    if last_row >= 2 and last_col >= 1: # Таблица должна иметь хотя бы 1 строку данных
                        ref = f"A1:{get_column_letter(last_col)}{last_row}"
                        tab = Table(displayName="Table_Data", ref=ref)
                        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                        tab.tableStyleInfo = style
                        ws_data.add_table(tab)
                        
                        # Автоширина колонок (примерно)
                        for column in ws_data.columns:
                            max_length = 0
                            column_letter = get_column_letter(column[0].column)
                            for cell in column:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            adjusted_width = (max_length + 2)
                            ws_data.column_dimensions[column_letter].width = min(adjusted_width, 50) 

                # --- Лист Logs ---
                if not combined_logs_df.empty:
                    ws_logs = writer.sheets['Logs']
                    last_row = len(combined_logs_df) + 1 # +1 для заголовка
                    last_col = len(combined_logs_df.columns)
                    if last_row >= 2 and last_col >= 1:
                        ref = f"A1:{get_column_letter(last_col)}{last_row}"
                        tab = Table(displayName="Table_Logs", ref=ref)
                        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                        tab.tableStyleInfo = style
                        ws_logs.add_table(tab)
                         
                         # Автоширина колонок
                        for column in ws_logs.columns:
                            max_length = 0
                            column_letter = get_column_letter(column[0].column)
                            for cell in column:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            adjusted_width = (max_length + 2)
                            ws_logs.column_dimensions[column_letter].width = min(adjusted_width, 70)

            log_callback(f"Всего добавлено {total_new_rows_count} строк данных. Файл сохранён: {output_excel_path}")
            
            if newly_processed_files_batch:    
                # Обновляем кэш
                updated_cache_list = list(processed_files_set.union(newly_processed_files_batch))
                cache[output_excel_key] = updated_cache_list
                save_cache(cache)
                log_callback("Кэш обновлён.")
                
        except Exception as e:
            log_callback(f"Ошибка при сохранении файла: {e}")
    
    else:
        log_callback("Нет новых данных и логов для сохранения.")

import pdfplumber
import pandas as pd
import os
import json
import shutil
import time
import re
import chardet
from concurrent.futures import ProcessPoolExecutor, as_completed

# Константы
number_of_list_for_analysis = 1
latencyBadClusters = -0.055

# ============================================================================
# ФУНКЦИИ ДЛЯ ОБРАБОТКИ HTML
# ============================================================================

def detect_encoding(file_path):
    """Определяет кодировку файла"""
    try:
        with open(file_path, 'rb') as file:
            raw_data = file.read(10000)
            result = chardet.detect(raw_data)
            encoding = result['encoding']
            return encoding if encoding else 'utf-8'
    except Exception:
        return 'utf-8'

def read_file_with_encoding(file_path):
    """Читает файл с правильной кодировкой"""
    encoding = detect_encoding(file_path)
    
    encodings_to_try = [
        encoding,
        'windows-1251',
        'cp1251',
        'utf-8',
        'iso-8859-1',
        'cp866',
        'koi8-r'
    ]
    
    for enc in encodings_to_try:
        try:
            with open(file_path, 'r', encoding=enc) as file:
                content = file.read()
            return content
        except (UnicodeDecodeError, LookupError):
            continue
    
    # Если ни одна кодировка не подошла
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            content = file.read()
        return content
    except Exception:
        return None

def extract_table_section(content):
    """Извлекает секцию с таблицами из HTML файла"""
    # Ищем секцию с таблицами (после "Параметры сечений")
    pattern = r'Параметры сечений[\s\S]*?конец протокола'
    match = re.search(pattern, content)
    
    if match:
        return match.group(0)
    return None

def parse_table_data(table_text, side_of_feather):
    """Парсит данные из таблицы (Спинка или Корыто)"""
    data = []
    
    if side_of_feather not in ['С', 'К']:
        return data
    
    # Разбиваем текст на строки
    lines = [line.rstrip() for line in table_text.split('\n')]
    
    # Ищем строку с заголовками (А2-А2, А3-А3 и т.д.)
    section_headers = []
    start_line_idx = 0
    
    for i, line in enumerate(lines):
        # Ищем заголовки в строке (могут быть кириллические или латинские A)
        if re.search(r'[АA]\d-[АA]\d', line.replace(' ', '')):
            # Извлекаем все заголовки
            headers = re.findall(r'[АA]\d-[АA]\d', line)
            if headers:
                section_headers = headers
                start_line_idx = i + 1
                break
    
    if not section_headers:
        return data
    
    # Парсим строки данных
    for i in range(start_line_idx, len(lines)):
        line = lines[i].strip()
        if not line:
            continue
        
        # Проверяем, является ли строка строкой данных (начинается с цифры и точки)
        if re.match(r'^\d+\.', line):
            # Убираем номер точки и точку
            parts = line.split('.', 1)
            if len(parts) < 2:
                continue
                
            point_num = int(parts[0].strip())
            
            # Пропускаем все точки больше 27 (28, 29, 30 и т.д.)
            if point_num > 27:
                continue
            
            values_str = parts[1].strip()
            
            # Разбиваем значения, учитывая возможные пробелы
            values = re.split(r'\s+', values_str)
            
            # Обрабатываем каждое значение для каждого сечения
            for j, value in enumerate(values):
                if j >= len(section_headers):
                    break
                
                # Извлекаем номер сечения из заголовка
                section_match = re.search(r'[АA](\d)-[АA]\d', section_headers[j])
                if section_match:
                    section = int(section_match.group(1))
                    
                    # Обрабатываем значение отклонения
                    deviation = None
                    if value and value != '--------':
                        # Убираем звездочку если есть
                        clean_value = value.replace('*', '')
                        try:
                            deviation = float(clean_value)
                        except ValueError:
                            deviation = None
                    
                    data.append({
                        'section': section,
                        'side_of_feather': side_of_feather,
                        'point': point_num,
                        'deviation': deviation
                    })
    
    return data

def extract_tables_from_section(table_text):
    """Извлекает данные из таблиц Спинка и Корыто"""
    all_data = []
    
    # Найдем позиции ключевых слов
    spinka_pos = table_text.find('Спинка:')
    korito_pos = table_text.find('Корыто:')
    ramka_pos = table_text.find('Рамка')
    
    if spinka_pos == -1 or korito_pos == -1:
        return all_data
    
    # Извлекаем таблицу Спинка (от 'Спинка:' до 'Корыто:')
    if korito_pos > spinka_pos:
        spinka_text = table_text[spinka_pos:korito_pos]
        spinka_data = parse_table_data(spinka_text, 'С')
        all_data.extend(spinka_data)
    
    # Извлекаем таблицу Корыто (от 'Корыто:' до 'Рамка' или конца)
    if ramka_pos != -1:
        korito_text = table_text[korito_pos:ramka_pos]
    else:
        korito_text = table_text[korito_pos:]
        
    korito_data = parse_table_data(korito_text, 'К')
    all_data.extend(korito_data)
    
    return all_data

def process_html_file(file_path):
    """Обрабатывает HTML файл и создает DataFrame"""
    try:
        # Читаем файл с правильной кодировкой
        content = read_file_with_encoding(file_path)
        if content is None:
            return None
        
        # Извлекаем секцию с таблицами
        table_section = extract_table_section(content)
        if not table_section:
            return None
        
        # Извлекаем данные из таблиц
        all_data = extract_tables_from_section(table_section)
        if not all_data:
            return None
        
        # Создаем DataFrame
        df = pd.DataFrame(all_data)
        
        if df.empty:
            return None
        
        # Убираем строки с пустыми отклонениями
        df = df.dropna(subset=['deviation'])
        
        # Сортируем для удобства просмотра
        df = df.sort_values(['side_of_feather', 'section', 'point']).reset_index(drop=True)
        
        return df
        
    except Exception as e:
        # print(f"Ошибка при обработке HTML файла: {e}")
        return None

# ============================================================================
# ФУНКЦИИ ДЛЯ ОБРАБОТКИ PDF
# ============================================================================

def parse_interval(key):
    low_str, high_str = key.strip("()").split(";")
    low = float(low_str)
    high = float(high_str) if high_str != ')' else float('inf')
    return low, high

def extract_page(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        if number_of_list_for_analysis < len(pdf.pages):
            page = pdf.pages[number_of_list_for_analysis]
            text = page.extract_text()
            return text
    return ""

def process_pdf_text(text):
    if not text:
        return pd.DataFrame()
        
    lines = text.split("\n")
    data = []
    for line in lines:
        if line.strip().startswith("А2_С"):
            cleaned_line = line.replace("(", "").replace(")", "").replace("_", " ")
            parts = cleaned_line.split()
            for i in range(0, len(parts), 4):
                if i + 3 < len(parts):  # Проверяем, что есть все 4 элемента
                    section = parts[i].strip()
                    side = parts[i+1].strip()
                    point = parts[i+2].strip()
                    deviation = parts[i+3].strip()
                    if len(section) > 0:
                        section = section[1:]
                    try:
                        section_num = int(section)
                        point_num = int(point)
                        if deviation == '':
                            continue
                        deviation_num = float(deviation)
                    except ValueError as ve:
                        continue

                    # Пропускаем все точки больше 27 (28, 29, 30 и т.д.)
                    if point_num > 27:
                        continue

                    data.append({
                        "section": section_num,
                        "side_of_feather": side,
                        "point": point_num,
                        "deviation": deviation_num
                    })
    df = pd.DataFrame(data, columns=["section", "side_of_feather", "point", "deviation"])
    return df

# ============================================================================
# УНИВЕРСАЛЬНАЯ ФУНКЦИЯ ОБРАБОТКИ ФАЙЛОВ
# ============================================================================

def process_any_file(file_path):
    """Универсальная функция для обработки файлов разных форматов"""
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.pdf':
        text = extract_page(file_path)
        df = process_pdf_text(text)
    elif file_ext in ['.html', '.htm']:
        df = process_html_file(file_path)
    else:
        return None
    
    return df

# ============================================================================
# ЛОГИКА АНАЛИЗА И КОПИРОВАНИЯ
# ============================================================================

def load_decision_table(path=None):
    if path is None:
        # Default to file in current directory
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "decision_table_T.json")
        
    # print(f"Загружается таблица решений из: {path}")
    if not os.path.exists(path):
        raise FileNotFoundError(f"Файл не найден: {path}")
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data
    except Exception as e:
        print(f"❌ Ошибка при чтении файла таблицы решений: {e}")
        raise

def move_or_copy_file_with_structure(src_path, source_folder, dest_folder, relative_path, status):
    """Копирует файл с сохранением структуры папок"""
    try:
        # Создаем путь для папки с соответствующим статусом
        status_dest_folder = os.path.join(dest_folder, relative_path, status)
        os.makedirs(status_dest_folder, exist_ok=True)
        
        # Формируем конечный путь
        filename = os.path.basename(src_path)
        dest_path = os.path.join(status_dest_folder, filename)
        
        shutil.copy2(src_path, dest_path)
        
        return True, dest_path
    except Exception as e:
        print(f"Ошибка при копировании файла: {e}")
        return False, None

def find_bad_clusters_info_fast(df):
    if df is None or df.empty:
        return []
        
    bad_df = df[df['deviation'] < latencyBadClusters].copy()
    if bad_df.empty:
        return []

    clusters = []
    used = set()

    coords = list(bad_df[['section', 'point']].itertuples(index=True))

    for idx1, sec1, pt1 in coords:
        if idx1 in used:
            continue
        cluster = [(idx1, sec1, pt1)]
        used.add(idx1)

        for idx2, sec2, pt2 in coords:
            if idx2 in used or idx1 == idx2:
                continue
            if abs(sec1 - sec2) <= 1 and abs(pt1 - pt2) <= 1:
                cluster.append((idx2, sec2, pt2))
                used.add(idx2)

        clusters.append(cluster)

    result = []
    for cluster_id, cluster in enumerate(clusters):
        size = len(cluster)
        for idx, sec, pt in cluster:
            result.append({
                'index': idx,
                'cluster_id': cluster_id,
                'value_bad_cluster': size
            })

    return result

def process_single_file(args):
    file_path, decision_table, deviation_ranges, source_folder = args
    file = os.path.basename(file_path)
    
    try:
        df1_final = process_any_file(file_path)
        
        if df1_final is None:
            return {
                'filename': file, 
                'deviation_counts': None, 
                'file_path': file_path,
                'relative_path': os.path.relpath(os.path.dirname(file_path), source_folder)
            }
        
        bad_clusters_info = find_bad_clusters_info_fast(df1_final)
        
        if bad_clusters_info:
            cluster_df = pd.DataFrame(bad_clusters_info)
            df1_final = df1_final.merge(
                cluster_df[['index', 'cluster_id', 'value_bad_cluster']],
                left_index=True,
                right_on='index',
                how='left'
            )
            df1_final.rename(columns={'index': 'original_index'}, inplace=True)
            df1_final_bad_only = df1_final.dropna(subset=['cluster_id'])
        else:
            df1_final_bad_only = pd.DataFrame()

        if not df1_final_bad_only.empty:
            deviation_counts = {}
            for dev_range_name, (low, high) in deviation_ranges.items():
                mask = df1_final_bad_only['deviation'].apply(
                    lambda x: low < abs(x) <= high if pd.notna(x) else False
                )
                count = mask.sum()
                deviation_counts[dev_range_name] = count
            return {
                'filename': file,
                'deviation_counts': deviation_counts,
                'file_path': file_path,
                'relative_path': os.path.relpath(os.path.dirname(file_path), source_folder)
            }
        else:
            return {
                'filename': file,
                'deviation_counts': None,
                'file_path': file_path,
                'relative_path': os.path.relpath(os.path.dirname(file_path), source_folder)
            }

    except Exception as e:
        return {
            'filename': file, 
            'deviation_counts': None, 
            'file_path': file_path,
            'relative_path': os.path.relpath(os.path.dirname(file_path), source_folder),
            'error': str(e)
        }

def process_folder_parallel(folder_path, decision_table):
    all_results = []
    deviation_ranges = {
        dev_range: parse_interval(dev_range)
        for dev_range in decision_table.keys()
    }

    files = []
    for root, _, fs in os.walk(folder_path):
        for file in fs:
            file_lower = file.lower()
            if file_lower.endswith(".pdf"):
                files.append((os.path.join(root, file), root))
            elif file_lower.endswith(".html") or file_lower.endswith(".htm"):
                if "откл" in file_lower:
                    files.append((os.path.join(root, file), root))
    
    args_list = [(f[0], decision_table, deviation_ranges, folder_path) for f in files]

    # Using ProcessPoolExecutor directly might handle the loop but we want to know progress
    # But since this function returns a DF, we will just blocking call it.
    
    with ProcessPoolExecutor() as executor:
        futures = [executor.submit(process_single_file, args) for args in args_list]
        for future in as_completed(futures):
            result = future.result()
            if result:
                all_results.append(result)

    return pd.DataFrame(all_results)

def apply_decision_table(df_results, decision_table):
    grouped_files = {}
    if df_results.empty:
        return grouped_files
        
    for _, row in df_results.iterrows():
        filename = row['filename']
        file_path = row.get('file_path', '')
        relative_path = row.get('relative_path', '')
        deviation_counts = row.get('deviation_counts')
        
        if not deviation_counts:
            grouped_files[filename] = {
                "details": [], 
                "final_status": "Годные", 
                "file_path": file_path,
                "relative_path": relative_path
            }
            continue

        details = []
        for dev_range, count in deviation_counts.items():
            cluster_group = None
            for group_key in decision_table.get(dev_range, {}).keys():
                low, high = parse_interval(group_key)
                if low <= count < high:
                    cluster_group = group_key
                    break
            if cluster_group is None:
                status = "Неизвестно"
            else:
                status = decision_table[dev_range].get(cluster_group, "Неизвестно")
            details.append(f"{dev_range}: {status}")

        final_status = "Годные"
        statuses = [d.split(": ")[1] for d in details]
        if "Брак" in statuses:
            final_status = "Брак"
        elif "Группировка" in statuses:
            final_status = "Группировка"

        grouped_files[filename] = {
            "details": details,
            "final_status": final_status,
            "file_path": file_path,
            "relative_path": relative_path
        }

    return grouped_files

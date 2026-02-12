import pandas as pd
import numpy as np
from collections import Counter
from scipy.signal import medfilt

def add_Index(df):
    df['Index'] = range(1, len(df) + 1)
    return df

def find_valid_segments_by_form(df):
    valid_segments = []
    start_index = None

    for i in range(len(df)):
        current_index = df.iloc[i]['Index']
        form_value = df.iloc[i]['Form']

        if form_value != 0:
            if start_index is None:
                start_index = current_index
        else:
            if start_index is not None:
                valid_segments.append((start_index, df.iloc[i - 1]['Index']))
                start_index = None

    if start_index is not None:
        valid_segments.append((start_index, df.iloc[-1]['Index']))

    return valid_segments

def find_fill(segment):
    if segment['Piro'].isna().all():
        return None

    max_row = segment.loc[segment['Piro'].idxmax()]
    max_index = max_row['Index']

    filtered = segment[segment['Index'] >= max_index].copy()

    indices = filtered['Index'].tolist()
    piro_values = filtered['Piro'].tolist()

    for i in range(len(piro_values) - 1):
        current_value = piro_values[i]
        next_value = piro_values[i + 1]

        if abs(next_value - current_value) > 35:
            window = piro_values[i:i+6]
            if any(val < 1200 for val in window):
                return indices[i]

    return None

def heating_to_fill(df, valid_segments):
    heating_fill_segments = []

    for start, end in valid_segments:
        segment = df[(df['Index'] >= start) & (df['Index'] <= end)]

        rows = segment.to_dict(orient='records')
        start_index = None
        for i in range(1, len(rows)):
            if rows[i]['Piro'] != rows[i - 1]['Piro']:
                start_index = rows[i]['Index']
                break

        end_index = find_fill(segment)

        if start_index is not None and end_index is not None:
            heating_fill_segments.append((start_index, end_index))

    return heating_fill_segments

def start_to_fill(df, valid_segments):
    holding_segments = []

    for start, end in valid_segments:
        segment = df[(df['Index'] >= start) & (df['Index'] <= end)]

        start_index = start
        end_index = find_fill(segment)

        if start_index is not None and end_index is not None:
            holding_segments.append((start_index, end_index))

    return holding_segments

def find_max_bp2(df, heating_segments):
    max_bp2_values = []
    for start, end in heating_segments:
        segment = df[(df['Index'] >= start) & (df['Index'] <= end)]
        if 'BP2' not in segment.columns:
            max_bp2_values.append(None)
            continue
        max_bp2 = segment['BP2'].max()
        max_bp2_values.append(max_bp2 if pd.notna(max_bp2) else None)
    return max_bp2_values


def find_min_bp2(df, heating_segments):
    min_bp2_values = []
    for start, end in heating_segments:
        segment = df[(df['Index'] >= start) & (df['Index'] <= end)]
        if 'BP2' not in segment.columns:
            min_bp2_values.append(None)
            continue
        min_bp2 = segment['BP2'].min()
        min_bp2_values.append(min_bp2 if pd.notna(min_bp2) else None)
    return min_bp2_values

def holding_time(df, holding_segments):
    holding_values = []
    for start, end in holding_segments:
        segment = df[(df['Index'] >= start) & (df['Index'] <= end)]
        holding = 0
        for _, row in segment.iterrows():
            if ((row['DL'] > 1500 and row['TL'] > 1500) or
                (row['DL'] > 1500 and row['TR'] > 1500) or
                (row['DR'] > 1500 and row['TL'] > 1500) or
                (row['DR'] > 1500 and row['TR'] > 1500)):
                holding += 5 / 60
        holding_values.append(holding)
    return holding_values

def heating_time(heating_segments):
    heating_times = []

    for start, end in heating_segments:
        heating_time = (end - start + 1) * 5 / 60
        heating_times.append(heating_time)

    return heating_times

def returning_ppf_temperature(df, valid_segments):
    results = []

    for start, end in valid_segments:
        segment = df[(df['Index'] >= start) & (df['Index'] <= end)]
        return_index = None

        # Перебираем по индексу DataFrame, а не по .iloc, чтобы работать с Index
        for i in range(1, len(segment)):
            current_index = segment.iloc[i]['Index']
            prev_value = segment.iloc[i - 1]['Form']
            current_value = segment.iloc[i]['Form']

            if current_value < prev_value:
                # Проверяем окно из 5 значений, включая текущий и до +4
                window_indices = [current_index + j for j in range(5)]
                
                # Проверяем, существуют ли все строки в пределах сегмента
                if window_indices[-1] > end:
                    continue  # Выход за границы сегмента

                try:
                    window_values = [df[df['Index'] == idx]['Form'].values[0] for idx in window_indices]
                except IndexError:
                    continue  # Одна из строк вне датасета или отсутствует

                # Проверяем, что все значения в окне уменьшаются
                is_decreasing = all(window_values[j] > window_values[j + 1] for j in range(len(window_values) - 1))

                if is_decreasing:
                    return_index = current_index
                    break  # Нашли первое подходящее окно — выходим из цикла

        if return_index is not None:
            return_row = df[df['Index'] == return_index].iloc[0]
            results.append((
                return_row['DL'],
                return_row['DR'],
                return_row['TL'],
                return_row['TR']
            ))
        else:
            results.append((None, None, None, None))

    return results

def holding_time_max_piro(df, valid_segments): 
    results = []
    for start, end in valid_segments:
        segment = df[(df['Index'] >= start) & (df['Index'] <= end)]
        max_temp = segment['Piro'].max()
        if max_temp > 1600:
            count = segment[segment['Piro'] >= 1600].shape[0]
        else:
            lower_bound = max_temp - 20
            upper_bound = max_temp
            count = segment[(segment['Piro'] >= lower_bound) & (segment['Piro'] <= upper_bound)].shape[0]
        results.append(count * 5)
    return results



def speed_form(df, valid_segments, tolerance=0.1):
    speed_form1_list = []
    speed_form2_list = []
    form_status_list = []

    for start, end in valid_segments:
        segment = df[(df['Index'] >= start) & (df['Index'] <= end)]
        if len(segment) < 2:
            speed_form1_list.append(None)
            speed_form2_list.append(None)
            form_status_list.append('Form_OK')
            continue

        rows = segment.to_dict(orient='records')

        # Расчёт сырой скорости (в час) для соседних точек в сегменте
        raw_speeds = []
        for i in range(1, len(rows)):
            diff = rows[i]['Form'] - rows[i - 1]['Form']
            raw_speeds.append(diff * 12)

        if len(raw_speeds) == 0:
            speed_form1_list.append(None)
            speed_form2_list.append(None)
            form_status_list.append('Form_OK')
            continue

        raw_speeds_arr = np.array(raw_speeds, dtype=float)

        # Медианный фильтр с окном 5 (первый/последний два значения заменим на NaN)
        filtered = medfilt(raw_speeds_arr, kernel_size=5)
        filtered = filtered.astype(float)
        if filtered.size >= 2:
            filtered[:2] = np.nan
            filtered[-2:] = np.nan

        # Положительные значения для подсчёта мод (по фильтрованным данным)
        mask_pos = (~np.isnan(filtered)) & (filtered > 0)
        filtered_pos = filtered[mask_pos]

        if filtered_pos.size == 0:
            speed_form1_list.append(None)
            speed_form2_list.append(None)
            form_status_list.append('Form_OK')
            continue

        # Биннинг по tolerance и определение мод по фильтрованным скоростям
        grouped_filtered = [round(val / tolerance) * tolerance for val in filtered_pos]
        speed_counts = Counter(grouped_filtered)
        most_common = speed_counts.most_common(2)

        speed_form1 = most_common[0][0]
        speed_form2 = most_common[1][0] if len(most_common) > 1 else None

        speed_form1_list.append(speed_form1)
        speed_form2_list.append(speed_form2)

        # Подготовка бинов для каждой точки фильтрованных скоростей
        binned_filtered_all = np.full_like(filtered, np.nan, dtype=float)
        valid_idx = (~np.isnan(filtered))
        binned_filtered_all[valid_idx] = np.array(
            [round(val / tolerance) * tolerance for val in filtered[valid_idx]]
        )

        # Функция расчёта доли совпадений для группы по моде
        def calc_match_ratio(target_mode_bin):
            if target_mode_bin is None:
                return None
            group_mask = valid_idx & (binned_filtered_all == target_mode_bin)
            idxs = np.where(group_mask)[0]
            if idxs.size == 0:
                return None
            diffs = np.abs(raw_speeds_arr[idxs] - filtered[idxs])
            matches = np.sum(diffs < 0.2)
            return matches / idxs.size if idxs.size > 0 else None

        ratio1 = calc_match_ratio(speed_form1)
        ratio2 = calc_match_ratio(speed_form2)

        # Логика статуса: если хотя бы в одной группе совпадений < 80% -> Form_Error
        status_error = False
        if ratio1 is not None and ratio1 < 0.8:
            status_error = True
        if ratio2 is not None and ratio2 < 0.8:
            status_error = True

        form_status_list.append('Form_Error' if status_error else 'Form_OK')

    return speed_form1_list, speed_form2_list, form_status_list

def find_stable_value(bp2_values):
    if bp2_values.size == 0:
        return 0
    data_range = (bp2_values.min(), bp2_values.max())
    hist, bin_edges = np.histogram(bp2_values, bins=100, range=data_range)
    stable_bin_index = np.argmax(hist)
    return (bin_edges[stable_bin_index] + bin_edges[stable_bin_index + 1]) / 2

def is_valid_growth(speeds):
    if not speeds:
        return False
    
    main_range = [0.045, 0.2]
    range_width = main_range[1] - main_range[0]
    
    min_val, max_val = min(speeds), max(speeds)
    
    subranges = []
    curr = main_range[0]
    while curr <= max_val + range_width:
        subranges.append((curr, curr + range_width))
        curr += range_width
    curr = main_range[0]
    while curr >= min_val - range_width:
        subranges.append((curr - range_width, curr))
        curr -= range_width
    
    subranges = sorted(set(subranges))
    counts = {sr: 0 for sr in subranges}
    for s in speeds:
        for sr in subranges:
            if sr[0] <= s < sr[1]:
                counts[sr] += 1
                break
    
    if not counts: return False
    max_sr = max(counts, key=counts.get)
    return max_sr == tuple(main_range)

def process_leakage_segment_math(segment_df):
    """Выполняет математический расчет скорости для одного сегмента"""
    target_offset = 120 #120*5/60=10 (минут)
    stabilization_limit = 12 #12*5/60=1 (минута)
    volume_coeff = 820 #4100/5=820 (литров/5 секунд)

    if len(segment_df) < 5:
        return None

    bp2_series = segment_df['BP2'].reset_index(drop=True)
    
    half_idx = len(bp2_series) // 2
    start_idx = bp2_series[:half_idx].idxmin()
    end_idx = bp2_series.idxmax()
    
    calc_start_idx = start_idx + stabilization_limit
    if calc_start_idx >= end_idx:
        calc_start_idx = start_idx
        
    bp2_start = bp2_series[calc_start_idx]
    
    if calc_start_idx + target_offset <= end_idx:
        calc_end_idx = calc_start_idx + target_offset
        actual_rows = target_offset
    else:
        calc_end_idx = end_idx
        actual_rows = end_idx - calc_start_idx
        
    if actual_rows <= 0:
        return None

    bp2_end = bp2_series[calc_end_idx]
    result = (bp2_end - bp2_start) * volume_coeff / actual_rows
    
    return {
        'offset': actual_rows,
        'result': result
    }

def get_last_valid_leakage(df):
    """Находит последнее валидное натекание в датафрейме"""
    if 'BP2' not in df.columns or 'Form' not in df.columns:
        return None

    temp_df = df.copy()
    temp_df['BP2'] = pd.to_numeric(temp_df['BP2'], errors='coerce') * 1000
    temp_df = temp_df[(temp_df['BP2'] <= 60) & (temp_df['Form'] == 0)].copy()

    if temp_df.empty:
        return None

    bp2_values = temp_df['BP2'].values
    stable_val = find_stable_value(bp2_values)
    stable_range = (stable_val - 2, stable_val + 2)
    
    stable_indices = [i for i, v in enumerate(bp2_values) if stable_range[0] <= v <= stable_range[1]]
    
    intervals = []
    if stable_indices:
        start = stable_indices[0]
        for i in range(1, len(stable_indices)):
            if stable_indices[i] != stable_indices[i-1] + 1:
                intervals.append((start, stable_indices[i-1]))
                start = stable_indices[i]
        intervals.append((start, stable_indices[-1]))

    segments = []
    for j in range(len(intervals) - 1):
        s_end = intervals[j][1]
        next_s_start = intervals[j+1][0]
        candidate = temp_df.iloc[s_end+1 : next_s_start]
        if candidate.empty or candidate['BP2'].max() > 60:
            continue
        speeds = candidate['BP2'].diff().dropna().tolist()
        if is_valid_growth(speeds):
            full_seg = temp_df.iloc[max(0, s_end-5) : min(len(temp_df), next_s_start+5)]
            segments.append(full_seg)
    
    if intervals:
        start_cand = temp_df.iloc[0 : intervals[0][0]]
        if not start_cand.empty and is_valid_growth(start_cand['BP2'].diff().dropna().tolist()):
            segments.append(temp_df.iloc[0 : min(len(temp_df), intervals[0][0]+5)])
        
        end_cand = temp_df.iloc[intervals[-1][1]+1 :]
        if not end_cand.empty and is_valid_growth(end_cand['BP2'].diff().dropna().tolist()):
            segments.append(temp_df.iloc[max(0, intervals[-1][1]-5) : len(temp_df)])

    valid_leakage = None
    for seg_df in segments:
        res = process_leakage_segment_math(seg_df)
        if res and res['offset'] >= 100:
            valid_leakage = res

    return valid_leakage

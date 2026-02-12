import pandas as pd
import core_algorithms as ca


def process_dataframe_segments(df, leakage_info=None):
    df = ca.add_Index(df)
    required_columns = ['Piro', 'BP2', 'Form', 'DL', 'TL', 'DR', 'TR']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        return []

    valid_segments = ca.find_valid_segments_by_form(df)
    if not valid_segments:
        return []

    returning_values = ca.returning_ppf_temperature(df, valid_segments)
    holding_time_max_piro_values = ca.holding_time_max_piro(df, valid_segments)
    speed_form1_values, speed_form2_values, form_status_list = ca.speed_form(df, valid_segments)

    result_rows = []
    for i, (start, end) in enumerate(valid_segments):
        segment = df[(df['Index'] >= start) & (df['Index'] <= end)]
        fill_value = ca.find_fill(segment)
        if fill_value is None:
            continue

        heating_segments = ca.heating_to_fill(df, [(start, end)])
        if not heating_segments:
            continue

        holding_segments = ca.start_to_fill(df, [(start, end)])
        if not holding_segments:
            continue

        max_bp2_values = ca.find_max_bp2(df, heating_segments)
        max_bp2 = max_bp2_values[0] if max_bp2_values else None

        min_bp2_values = ca.find_min_bp2(df, heating_segments)
        min_bp2 = min_bp2_values[0] if min_bp2_values else None

        holding_values = ca.holding_time(df, holding_segments)
        holding = holding_values[0] if holding_values else 0

        heating_times = ca.heating_time(heating_segments)
        heating = heating_times[0] if heating_times else 0

        fill_row = df[df['Index'] == fill_value].iloc[0]

        result_row = fill_row.copy()
        result_row['Давление макс'] = max_bp2
        result_row['Давление мин'] = min_bp2
        result_row['Выдержка ППФ'] = holding
        result_row['Время нагрева'] = heating
        result_row['ReturnLL'], result_row['ReturnLR'], result_row['ReturnTL'], result_row['ReturnTR'] = returning_values[i]
        result_row['Время удержания макс. темп.'] = holding_time_max_piro_values[i]
        result_row['max_temp'] = segment['Piro'].max()
        result_row['Термопара_макс'] = segment[segment['TP'] <= 1700]['TP'].max() if 'TP' in segment.columns else None
        result_row['speed_form_1'] = speed_form1_values[i]
        result_row['speed_form_2'] = speed_form2_values[i]
        result_row['Статус формы'] = form_status_list[i]
        result_row['Максимальное положение блока'] = segment['Form'].max()

        # Добавляем данные о натекании
        if leakage_info:
            result_row['offset_leakage'] = leakage_info.get('offset')
            result_row['leakage'] = leakage_info.get('result')
            result_row['leakage_source'] = leakage_info.get('source')
        else:
            result_row['offset_leakage'] = None
            result_row['leakage'] = None
            result_row['leakage_source'] = None

        result_rows.append(result_row)
    return result_rows


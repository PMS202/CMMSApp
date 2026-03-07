import pandas as pd
import re
import calendar as cal
def filter_data_FG(data_path: str, sheet_name: str, date_col: int, line_col: int, month: int, year: int) -> pd.DataFrame:

    # data_file = pd.read_excel(data_path, sheet_name=sheet_name)
    try:
        data_file = pd.read_excel(data_path, sheet_name=sheet_name, engine= "xlrd")
    except Exception as e:
        data_file = pd.read_excel(data_path, sheet_name=sheet_name, engine= "openpyxl")
    date_series = pd.to_datetime(data_file.iloc[:, date_col], errors='coerce')
    valid_date_indexes = date_series[
        (date_series.notna()) &
        (date_series.dt.month == month) &
        (date_series.dt.year == year)
    ].index
    # line_series = data_file.iloc[:, line_col]

    total_lines = [f"F0{i}" if i < 10 else f"F{i}" for i in range(1, 25)]

    result_df = pd.DataFrame(index=[i for i in range(1,cal.monthrange(year,month)[1]+1)],columns=total_lines)
    working_time = result_df.copy()
    for i, date_idx in enumerate(valid_date_indexes, start=1):
        block = data_file.iloc[date_idx + 1: date_idx + 25]
        row_result = {}
        row_result_time = {}
        for line_name in total_lines:
            mask = block.iloc[:, line_col] == line_name
            if mask.any():
                row = block[mask].iloc[0]
                try:
                    row.iloc[line_col + 1: line_col + 25] = pd.to_numeric(row.iloc[line_col + 1: line_col + 25], errors='coerce').fillna(0).astype(int)
                    sum_val = row.iloc[line_col + 1: line_col + 25].sum()
                    count_val = (row.iloc[line_col + 1: line_col + 25] > 0).sum()
                except Exception as e:
                    print(f"Error processing line {line_name} on date {date_series[date_idx]}: {e}")
                    sum_val = 0
                    count_val = 0
                row_result[line_name] = sum_val
                row_result_time[line_name] = count_val
            else:
                row_result[line_name] = 0
                row_result_time[line_name] = 0
        result_df.loc[date_series.dt.day[date_idx],:] = row_result
        working_time.loc[date_series.dt.day[date_idx],:] = row_result_time
        working_time = working_time.apply(lambda row: row.map(lambda x: 0 if x < 3 else ( 8 if 3 < x < 10 else (12 if 10 <= x < 13 else (16 if 13 <= x < 17 else 24)))), axis=1)
    result_df = result_df.apply(lambda col: pd.to_numeric(col, errors='coerce')).fillna(0).astype(int)
    return result_df,working_time

# result_df,working_time = filter_data_FG(r"C:\Users\2173452100291\Documents\Excel_data\Molding daily Mar-25.xlsx", "Molding Mar-25 ",0, 1, 3, 2025)
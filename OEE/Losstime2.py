import pandas as pd
import calendar as cal

def Losstime(data_path: str, sheet_name: str, date_col: int, month: int, year: int) -> pd.DataFrame:

    try:
        data_file = pd.read_excel(data_path, sheet_name=sheet_name, header=2, engine= "xlrd")
    except Exception as e:
        data_file = pd.read_excel(data_path, sheet_name=sheet_name, header=2, engine= "openpyxl")
    date_format: str = "%d/%m/%Y"
    data_file['ParsedDate'] = pd.to_datetime(data_file.iloc[:, date_col], format=date_format,errors='coerce')
    mask = data_file['ParsedDate'].notna() & (data_file['ParsedDate'].dt.month == month) & (data_file['ParsedDate'].dt.year == year)
    df_valid = data_file[mask].copy()

    df_valid['Day'] = df_valid['ParsedDate'].dt.day

    total_lines = [f"F0{i}" if i < 10 else f"F{i}" for i in range(1, 25)]

    common_lines = [col for col in total_lines if col in df_valid.columns]
    df_valid[common_lines] = df_valid[common_lines].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
    result_df = df_valid.groupby('Day')[common_lines].sum()

    result_df = result_df.reindex(columns=total_lines, fill_value=0)
    result_df = result_df.apply(lambda col: pd.to_numeric(col, errors='coerce')).fillna(0).astype(int)
    return result_df

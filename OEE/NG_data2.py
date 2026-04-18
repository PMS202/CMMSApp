import pandas as pd
import re
import calendar as cal

def filter_data_NG(data_path: str, sheet_name: str, date_col: int, begin: int, end: int,month: int, year: int) -> pd.DataFrame:

    try:
        data_file = pd.read_excel(data_path, sheet_name=sheet_name, header=2, engine= "xlrd")
    except Exception as e:
        data_file = pd.read_excel(data_path, sheet_name=sheet_name, header=2, engine= "openpyxl")
    col1 = data_file.iloc[:, date_col]
    date_series = pd.to_datetime(col1, errors='coerce')

    valid_date_indexes = date_series[(date_series.notna()) & (date_series.dt.month == month) & (date_series.dt.year == year)].index

    pattern = re.compile(r"^F(0[1-9]|1[0-9]|2[0-9])$")
    pattern2 = re.compile(r"^FA(0[1-9]|1[0-9]|2[0-9])$")
    valid_columns = [col for col in data_file.columns if isinstance(col, str) and (pattern.match(col) or pattern2.match(col))]
    data_file.fillna(0, inplace=True)
    NG_data = pd.DataFrame(index=[i for i in range(1,cal.monthrange(year,month)[1]+1)],columns=valid_columns)
    FG_data = pd.DataFrame(index=[i for i in range(1,cal.monthrange(year,month)[1]+1)],columns=valid_columns)

    for i, valid_idx in enumerate(valid_date_indexes, start=1):
        data_file.loc[valid_idx + begin: valid_idx + end, valid_columns] = data_file.loc[valid_idx + begin: valid_idx + end, valid_columns].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
        NG_data.loc[i] = data_file.loc[valid_idx + begin: valid_idx + end, valid_columns].sum()
        FG_data.loc[i] = data_file.loc[valid_idx + 22, valid_columns]

    if "molding" in sheet_name.lower() or "final" in sheet_name.lower():
        NG_data = NG_data.apply(lambda col: pd.to_numeric(col, errors='coerce')).fillna(0).astype(int)
        return NG_data
    elif "coil" in sheet_name.lower():
        NG_data = NG_data.apply(lambda col: pd.to_numeric(col, errors='coerce')).fillna(0).astype(int)
        FG_data = FG_data.apply(lambda col: pd.to_numeric(col, errors='coerce')).fillna(0).astype(int)
        return NG_data, FG_data
    else:
        return None


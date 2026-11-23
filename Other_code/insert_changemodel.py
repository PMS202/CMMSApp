import pandas as pd
import datetime as dt
import re
import numpy as np
import os
import sys
from pathlib import Path

# Add project root: c:\Users\2173452100291\Documents\program
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process

excel_file_path = r"C:\Users\2173452100291\Downloads\HieuMPX.xlsx"

def read_excel_file(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name="Sheet1")
        # df = df.replace({np.nan: 0})
        time_columns = ["Start_Time", "Start_Repair_Time", "End_Time"]
        for col in time_columns:
            df[col] = df[col].apply(normalize_time)
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

def normalize_time(value):
    if pd.isna(value):
        return pd.NA
    if isinstance(value, (pd.Timestamp, dt.datetime, dt.time)):
        return value.strftime("%H:%M")
    if isinstance(value, (pd.Timedelta, dt.timedelta)):
        total_seconds = int(value.total_seconds()) % 86400
        hours, remainder = divmod(total_seconds, 3600)
        minutes = remainder // 60
        return f"{hours:02d}:{minutes:02d}"
    if isinstance(value, (int, float)):
        if 0 <= value < 1:
            total_minutes = round(value * 1440) % 1440
            hours, minutes = divmod(total_minutes, 60)
            return f"{hours:02d}:{minutes:02d}"
    value = str(value).strip()
    if value.lower() in ("<na>", "nan", "nat", ""):
        return pd.NA
    value = re.sub(
        r"^24:(\d{2})(:\d{2})?$",
        r"00:\1\2",
        value
    )
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return dt.datetime.strptime(value, fmt).strftime("%H:%M")
        except ValueError:
            pass
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.notna(parsed):
        return parsed.strftime("%H:%M")
    return value

def insert(df):
    if df.empty:
        print("No data to insert.")
        return

    DB = Database_process()

    sql =   """
            UPDATE downtime_records as dr
            SET dr.machine_id = :machine_code
            WHERE dr.downtime_record_id = :downtime_record_id;
            """
    param_list = []
    try:
        df_filtered = df
        for index, row in df_filtered.iterrows():
            param_list.append({
                "downtime_record_id": row["Downtime_ID"],
                "downtime_start_time": row["Start_Time"],
                "downtime_start_repair_time": row["Start_Repair_Time"],
                "downtime_end_time": row["End_Time"],
                "staff_name": row["Staff_Name"],
                "error_code": row["Error_Code"],
                "machine_code": row["Machine_Code"]
            })
        DB.executemany(sql = sql, params_list= param_list)
    except Exception as e:
        print(f"Error inserting data into database: {e}")

def update(df):
    if df.empty:
        print("No data to update.")
        return

    DB = Database_process()

    sql = """
        UPDATE machine_cycle_times as mct
        JOIN product_models_oee as pmo ON mct.model_id = pmo.model_id
        SET mct.cycle_time_seconds = :cycle_time, mct.create_at = "2025-05-01 00:00:00"
        WHERE pmo.model_name = :model_name;
    """
    param_list = []
    try:
        for index, row in df.iterrows():
            param_list.append({
                "model_name": row["model"],
                # "line_name": row["Line Name"],
                "cycle_time": row["cycletime"]
            })
        result = DB.executemany(sql=sql, params_list=param_list)
        print(f"Data updated successfully. Rows affected: {result}")
    except Exception as e:
        print(f"Error updating data into database: {e}")


if __name__ == "__main__":
    df = read_excel_file(excel_file_path)
    # cycle_time_df = df[["model", "cycletime"]]
    # cycle_time_df = cycle_time_df.drop_duplicates(
    #                                     subset=["model", "cycletime"],
    #                                     keep="first"
    #                                 ).reset_index(drop=True)
    # print(cycle_time_df)
    # update(cycle_time_df)
    insert(df)
    # print(df)



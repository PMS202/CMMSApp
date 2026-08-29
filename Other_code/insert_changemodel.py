import pandas as pd
from datetime import datetime
import os
import sys
from pathlib import Path

# Add project root: c:\Users\2173452100291\Documents\program
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process

excel_file_path = r"C:\Users\2173452100291\Downloads\dt_SCA_Kha.xlsx"

def read_excel_file(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name="Sheet1")
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()



def insert(df):
    if df.empty:
        print("No data to insert.")
        return

    DB = Database_process()

    sql = """
            UPDATE downtime_records
            SET error_code = :new_error_code
            WHERE downtime_record_id = :downtime_record_id;
    """
    param_list = []
    try:
        for index, row in df.iterrows():
            param_list.append({
                "new_error_code": row["error_code"],
                "downtime_record_id": row["downtime_record_id"]
            })
        DB.executemany(sql = sql, params_list= param_list)
        # print(f"Data inserted successfully. Rows affected: {len(param_list)}")
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



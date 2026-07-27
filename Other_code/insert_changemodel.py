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

excel_file_path = r"F:\CMMS prepare data\Kha_ML26_group2 (Thach) - Jun.26 - v2.xlsx"

def read_excel_file(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name="Output")
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()



def insert(df):
    if df.empty:
        print("No data to insert.")
        return

    DB = Database_process()

    sql = """UPDATE line_operation_times
                JOIN production_lines AS pl ON line_operation_times.line_id = pl.line_id
                SET operation_hours = :operation_hours
                WHERE pl.line_name = :line_name AND operation_date = :operation_date AND model_running = :model_running;
    """
    param_list = []
    try:
        for index, row in df.iterrows():
            param_list.append({
                "line_name": row["Line Name"],
                "operation_date": row["Date"],
                "operation_hours": row["WT"],
                "model_running": row["Model Name"]
            })
        DB.executemany(sql = sql, params_list= param_list)
        print("Data inserted successfully.")
    except Exception as e:
        print(f"Error inserting data into database: {e}")

def update(df):
    if df.empty:
        print("No data to update.")
        return

    DB = Database_process()

    sql = """
        UPDATE machine_cycle_times as mct
        JOIN machine_oee_register as mor ON mct.machine_id = mor.machine_id AND mct.model_id = mor.model_id
        JOIN production_lines as pl ON mor.line_id = pl.line_id
        JOIN product_models_oee as pmo ON mor.model_id = pmo.model_id
        SET mct.cycle_time_seconds = :cycle_time, mct.create_at = "2025-05-01 00:00:00"
        WHERE pl.line_name = :line_name
        AND pmo.model_name = :model_name;
    """
    param_list = []
    try:
        for index, row in df.iterrows():
            param_list.append({
                "model_name": row["Model Name"],
                "line_name": row["Line Name"],
                "cycle_time": row["Cycle time"]
            })
        result = DB.executemany(sql=sql, params_list=param_list)
        print(f"Data updated successfully. Rows affected: {result}")
    except Exception as e:
        print(f"Error updating data into database: {e}")


if __name__ == "__main__":
    df = read_excel_file(excel_file_path)
    # cycle_time_df = df[["Line Name", "Model Name", "Cycle time"]]
    # cycle_time_df = cycle_time_df.drop_duplicates(
    #                                     subset=["Line Name", "Model Name"],
    #                                     keep="first"
    #                                 ).reset_index(drop=True)
    # print(cycle_time_df)
    # update(cycle_time_df)
    insert(df)



import pandas as pd
import os
import sys
import re
from pathlib import Path
import numpy as np
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process



file_path = r"C:\Users\2173452100291\Downloads\Downtime CMMS MTPI (1).xlsx"
sheet_names_list = ["WT-Jan-26","WT-Feb-26","WT-Mar-26","WT-Apr-26","WT-May-26","WT-Jun-26"]
excel_file = pd.ExcelFile(file_path)
final_working_time = pd.DataFrame()
for sheet_name in sheet_names_list:
    working_time = pd.read_excel(excel_file, sheet_name=sheet_name, header=0)
    working_time["Date"] = pd.to_datetime(working_time["Date"], errors='coerce')
    working_time = working_time[working_time["Date"].notna()]
    working_time = working_time.fillna(0)
    working_time["Date"] = working_time["Date"].dt.strftime("%Y-%m-%d")
    sum_columns = ["OK", "NG", "WT"]
    working_time[sum_columns] = working_time[sum_columns].apply(
        pd.to_numeric, errors="coerce"
    ).fillna(0)
    working_time = (
        working_time
        .groupby(["Line", "Model", "Date"], as_index=False, dropna=False)
        .agg({"Shift": "first", "OK": "sum", "NG": "sum", "Cycle time": "first", "WT": "sum"})
    )
    final_working_time = pd.concat([final_working_time, working_time], ignore_index=True)

# print(final_working_time[final_working_time.isin(["A01"]).any(axis=1)])
# params_list = [
#     {
#         "operation_hours": row["WT"],
#         "line_name": row["Line"],
#         "operation_date": row["Date"],
#         "model_running": row["Model"]
#     }
#     for _, row in final_working_time.iterrows()
# ]
# print(params_list)
database = Database_process()

database.executemany(sql = f''' UPDATE line_operation_times AS lot
JOIN production_lines AS pl ON lot.line_id = pl.line_id
SET lot.operation_hours = :operation_hours
WHERE pl.line_name = :line_name AND lot.operation_date = :operation_date AND lot.model_running = :model_running;''', params_list=[
    {
        "operation_hours": row["WT"],
        "line_name": str(row["Line"]),
        "operation_date": str(row["Date"]),
        "model_running": str(row["Model"])
    }
    for _, row in final_working_time.iterrows()
])

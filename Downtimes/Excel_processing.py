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

class Downtime_Excel_Processor:
    def __init__(self, file_path,sheet_name=None, database = None,area_name=None):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.data = None
        self.database = database
        self.area_name = area_name
        self.error_log = pd.DataFrame(columns=["row_index", "error_type", "error_message"])
        self.error_frame = pd.DataFrame(columns=["date", "line", "start_time","technical_start_time","finish_time","technical_name", "error_code", "machine_code", "column_error", "error_message"])

    def read_filter_excel(self):
        change_model_code_list = ["100","199"]
        def add_error_info(data, column_error, error_message):
            data = data.copy()
            data["column_error"] = column_error
            data["error_message"] = error_message
            return data
        pattern = r'^\d{2}:\d{2}(:\d{2})?$'
        try:
            temp = self.database.query('''SELECT pl.line_name 
                                            FROM production_lines pl 
                                            JOIN downtime_areas_production_lines dapl 
                                                ON pl.line_id = dapl.line_id
                                            JOIN downtime_areas da 
                                                ON dapl.downtime_area_id = da.downtime_area_id
                                            WHERE da.downtime_area_name = :area_name
                                            ORDER BY pl.line_name ASC;''', params = {"area_name": self.area_name})
            line_name_list = [line[0] for line in temp]
            temp = self.database.query('''SELECT m.machine_code
                                            FROM machines as m
                                            JOIN production_lines pl 
                                                ON m.line_id = pl.line_id
                                            JOIN downtime_areas_production_lines dapl 
                                                ON pl.line_id = dapl.line_id
                                            JOIN downtime_areas da 
                                                ON dapl.downtime_area_id = da.downtime_area_id
                                            WHERE da.downtime_area_name = :area_name AND m.machine_status = "GOOD";''', params = {"area_name": self.area_name})
            machine_code_list = [code[0] for code in temp]
            self.data = pd.read_excel(self.file_path, sheet_name=self.sheet_name, skiprows=4, 
                                      header=None, usecols=[0,1,2,3,4,9,10,17], dtype={10: str, 17: str})
            self.data = self.data.rename(columns={0: "date", 1: "line", 2: "start_time", 3: "technical_start_time", 4: "finish_time", 9: "technical_name", 10: "error_code", 17: "machine_code"})
            self.data = self.data.replace({"nan": pd.NA})
            time_columns = ["start_time", "technical_start_time", "finish_time"]
            self.data = self.data[self.data["date"].notna()]
            for col in time_columns:
                self.data[col] = self.data[col].replace({"24:00": "00:00"})
                self.data[col] = self.data[col].apply(
                    lambda x: str(x).strftime("%H:%M") if hasattr(str(x), 'strftime') and pd.notna(x) 
                    else (pd.to_datetime(str(x), format="%H:%M", errors='coerce').strftime("%H:%M") if pd.notna(x) else str(x))
                )
            normalized = self.sheet_name.replace("-", " ").replace(".", " ")
            dt = pd.to_datetime(normalized, format="%b %y")
            self.month_year = dt.strftime("%Y-%m")
            max_date = pd.Period(self.month_year).days_in_month
            self.error_frame = pd.concat([self.error_frame, 
                                          add_error_info(self.data[(self.data["date"].isna()) | (self.data["date"]>max_date) | (self.data["date"]<1)],
                                                                "date", "Invalid date")])
            self.error_frame = pd.concat([self.error_frame, 
                                          add_error_info(self.data[(self.data["technical_name"].isna())],
                                                               "technical_name", "Missing technical name")])
            self.error_frame = pd.concat([self.error_frame, 
                                          add_error_info(self.data[~self.data["start_time"].astype(str).str.match(pattern)], 
                                                              "start_time", "Invalid start time")])
            self.error_frame = pd.concat([self.error_frame, 
                                          add_error_info(self.data[~self.data["technical_start_time"].astype(str).str.match(pattern)], 
                                                              "technical_start_time", "Invalid technical start time")])
            self.error_frame = pd.concat([self.error_frame, 
                                          add_error_info(self.data[~self.data["finish_time"].astype(str).str.match(pattern)], 
                                                              "finish_time", "Invalid finish time")])
            self.error_frame = pd.concat([self.error_frame, 
                                          add_error_info(self.data[(~self.data["error_code"].isin(change_model_code_list)) & (self.data["machine_code"].isna())], 
                                                              "machine_code", "Machine code missing")])
            self.error_frame = pd.concat([self.error_frame, 
                                          add_error_info(self.data[~self.data["line"].isin(line_name_list)], 
                                                              "line", "Invalid line")])
            self.error_frame = pd.concat([self.error_frame, 
                                          add_error_info(self.data[~self.data["machine_code"].isin(machine_code_list)], 
                                                              "machine_code", "Invalid machine code")])
            self.error_frame = self.error_frame.drop_duplicates()
            self.data = self.data[~self.data.index.isin(self.error_frame.index)]
            
            total_loss = (pd.to_datetime(self.data["finish_time"], format="%H:%M", errors='coerce') - 
                          pd.to_datetime(self.data["start_time"], format="%H:%M", errors='coerce')).dt.total_seconds() / 60
            total_loss = total_loss.where(total_loss >= 0, total_loss + 1440)
            self.data.insert(5, "total_loss_time", total_loss)
            
            wait_technical = (pd.to_datetime(self.data["technical_start_time"], format="%H:%M", errors='coerce') - 
                              pd.to_datetime(self.data["start_time"], format="%H:%M", errors='coerce')).dt.total_seconds() / 60
            wait_technical = wait_technical.where(wait_technical >= 0, wait_technical + 1440)

            self.data.insert(6, "wait_technical_time", wait_technical)
            return self.data,self.error_frame
        except Exception as e: 
            raise Exception(f"Error reading Excel file: {e}")
            return None, None

if __name__ == "__main__":
    file_path = r"C:\Users\2173452100291\Documents\program\Copy of Downtime  SC-A-Tu.xlsx"
    sheet_name = "Nov-25"
    # file_path = r"\\172.30.73.156\ushare\2173452300015\7. AUDITS\1. Monthly data analysis\1. 2025\1. Downtime_IATF\4. Other\Downtime HOKKO-2025 Rev.xlsx"
    # sheet_name = "Dec.25"
    month_year = "2025-11"
    database = Database_process()
    processor = Downtime_Excel_Processor(file_path, sheet_name, database,area_name="SC-A")
    processor.read_filter_excel()
    print("Filtered Data:")
    print(processor.error_frame)
    
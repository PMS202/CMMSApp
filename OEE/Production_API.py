import requests
import time
import sys
from pathlib import Path
from datetime import datetime as dt
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process

class Production_API:
    def __init__(self,params_dict, API_URL, database=None):
        self.params_dict = params_dict
        self.API_URL = API_URL
        self.database = database
 
    def get_inspection_output(self,API=None,params: tuple=None):
        last_day_of_month = lambda year, month: (dt(year, month % 12 + 1, 1) - dt(year, month, 1)).days
        if API is None:
            API = self.API_URL
        if params is None:
            params = self.params_dict
        if not isinstance(params, dict) or not isinstance(API, str):
            raise ValueError("params must be a dict and API must be a string")
        date_tuple = list(params.values())[0]
        if not isinstance(date_tuple, tuple):
            raise ValueError("params values must be a tuple")
        begin_date, end_date = date_tuple[0].split("-"), date_tuple[-1].split("-")
        year = begin_date[0]
        key = list(params.keys())[0]
        result = []
        for month in range(int(begin_date[1]), int(end_date[1])+1):
            for date in range(int(begin_date[2]), last_day_of_month(int(year), month)+1):
                fdate = f"{year}-{month:02d}-{date:02d}"
                resp = requests.get(API, params={key: fdate}, timeout=10)
                resp.raise_for_status()
                time.sleep(0.2)
                with_date = resp.json()
                with_date["productionDate"] = fdate
                result.append(with_date)
        return result

if __name__ == "__main__":
    params_dict = {"productionDate": ("2026-05-01", "2026-05-31")}
    API_URL = "http://172.30.73.149:1810/ScaMonitor/GetInspectionOkNg_All?"
    production_api = Production_API(params_dict, API_URL)
    data = production_api.get_inspection_output()
    try:
        database = Database_process()
    except Exception as e:
        print(f"Error connecting to database: {e}")
        sys.exit(1)
    try:
        lines = database.query(sql = ''' SELECT pl.line_name 
                                    FROM production_lines as pl
                                    JOIN downtime_areas_production_lines as dapl ON pl.line_id = dapl.line_id
                                    JOIN downtime_areas as da ON dapl.downtime_area_id = da.downtime_area_id
                                    WHERE da.downtime_area_name = "SC-A"; '''
                                            "")
        lines = [line[0] for line in lines]
        history_model_ofLines = database.query(sql = ''' SELECT pl.line_name,lot.operation_date,lot.operation_hours,lot.change_model, lot.change_from
                                                        FROM line_operation_times AS lot
                                                        JOIN production_lines AS pl ON lot.line_id = pl.line_id
                                                        WHERE lot.change_model IS NOT NULL AND MONTH(lot.operation_date) = 12 AND YEAR(lot.operation_date) = 2025
                                                        ORDER BY lot.line_id;''') #đổi change_model qua bảng downtime_record để lấy model
        history_model_dict = {}
        for item in history_model_ofLines:
            if item[0] not in history_model_dict:
                history_model_dict[item[0]] = {}
            history_model_dict[item[0]][item[1].strftime("%Y-%m-%d")] = {"operation_hours": item[2], "model": item[3], "change_from": str(item[4])}
    except Exception as e:
        print(f"Error inserting data into database: {e}")
        sys.exit(1)
    params_list  = []
    
    SELECT_CURRENT_MODEL_SQL = ''' SELECT pl.line_name, po.model_name
                                            FROM `production_lines` as pl
                                            JOIN `production_output` as po ON pl.line_id = po.line_id
                                            WHERE po.production_date = (SELECT MAX(production_date) FROM production_output WHERE line_id = pl.line_id);'''
    current_model_flag = {}
    try:
        current_models = database.query(sql=SELECT_CURRENT_MODEL_SQL)
        for line, model in current_models:
            current_model_flag[line] = model
    except Exception as e:
        print(f"Error fetching current models from database: {e}")
        sys.exit(1)
    start_date = dt.strptime("06:00:00", "%H:%M:%S")
    
    for item in data:
        production_date = item.pop("productionDate")
        for line in lines:
            try:
                item[line]["OkQty"]
            except KeyError:
                continue
            if line in item:
                try:
                    model_info = history_model_dict[line].get(production_date, {"model": None, "change_from": None})
                except KeyError:
                    model_info = {"model": None, "change_from": None}
                if model_info["model"] is None or model_info["model"] == current_model_flag[line]:
                    model_info["model"] = current_model_flag.get(line, None)
                else:
                    change_time = dt.strptime(history_model_dict[line][production_date]["change_from"], "%H:%M:%S")
                    total_duration_SEC =  float(history_model_dict[line][production_date]["operation_hours"]*3600)
                    duration_model_old = change_time - start_date
                    duration_model_old_SEC = (duration_model_old.total_seconds() / 3600 + 24)*3600 if duration_model_old.total_seconds() < 0 else duration_model_old.total_seconds()
                    if duration_model_old_SEC > float(total_duration_SEC):
                        raise ValueError(f"Duration of old model for line {line} on {production_date} exceeds total operation hours.")
                    percentage_old_model = duration_model_old_SEC / float(total_duration_SEC)
                    old_model_output = int(item[line]["OkQty"]*percentage_old_model)
                    new_model_output = item[line]["OkQty"] - old_model_output
                    param = {"line_name": line, "production_date": production_date, "model_name": current_model_flag[line] , "OK_qty": old_model_output, "NG_qty": 0}
                    params_list.append(param)
                    param = {"line_name": line, "production_date": production_date, "model_name": model_info["model"], "OK_qty": new_model_output, "NG_qty": item[line]["NgQty"]}
                    params_list.append(param)
                    current_model_flag[line] = model_info["model"]
                    continue
                param = {"line_name": line, "production_date": production_date, "model_name": model_info["model"] if model_info["model"] is not None else "Unknown", "OK_qty": item[line]["OkQty"], "NG_qty": item[line]["NgQty"]}
                params_list.append(param)
    try:
        database.executemany(sql = ''' INSERT IGNORE INTO `production_output` (line_id, production_date, model_name, OK_qty, NG_qty)
                                        VALUES ((SELECT line_id FROM production_lines WHERE line_name = :line_name), 
                                                :production_date, :model_name, :OK_qty, :NG_qty)
                                                ON DUPLICATE KEY UPDATE 
                                                OK_qty     = VALUES(OK_qty),
                                                NG_qty     = VALUES(NG_qty);''', params_list = params_list)
    except Exception as e:
        print(f"Error inserting data into database: {e}")
        sys.exit(1)
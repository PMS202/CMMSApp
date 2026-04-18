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
        month = begin_date[1]
        year = begin_date[0]
        key = list(params.keys())[0]
        result = []
        for date in range(int(begin_date[2]), int(end_date[2])+1):
            fdate = f"{year}-{month}-{date:02d}"
            resp = requests.get(API, params={key: fdate}, timeout=10)
            resp.raise_for_status()
            time.sleep(0.2)
            with_date = resp.json()
            with_date["productionDate"] = fdate
            result.append(with_date)
        return result

if __name__ == "__main__":
    params_dict = {"productionDate": ("2025-12-01", "2025-12-31")}
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
                                                        ORDER BY lot.line_id;''')
        history_model_dict = {}
        for item in history_model_ofLines:
            if item[0] not in history_model_dict:
                history_model_dict[item[0]] = {}
            history_model_dict[item[0]][item[1].strftime("%Y-%m-%d")] = {"operation_hours": item[2], "model": item[3], "change_from": str(item[4])}
    except Exception as e:
        print(f"Error inserting data into database: {e}")
        sys.exit(1)
    params_list  = []
    current_model_flag = {  "A02": "SCFN3323XV-450-1R5A052H-T",
                            "A03": "SCF29-300-1R8A018JV",
                            "A04": "SC14-250-1R4A55UH",
                            "A05": "SCF29-300-1R8A018JV",
                            "A06": "Unknown",
                            "A07": "SCF25XV-280-2R1A005JH",
                            "A08": "SCF29XV-210-1R9A012JH-CG(SA)",
                            "A09": "SCFN3021-300-1R2A051H",
                            "A10": "SCF25-000-2R1B002JV-VT",
                            "A11": "SCF14XV-1250-1R6A94UJH",
                            "A12": "SCN46-320-2R2AJH-BW",
                            "A13": "SCF29-300-1R8A018JV",
                            "A14": "SCN3222-300-1R3A008H",
                            "A15": "SCN3222-300-1R3A008H"
                          } # cần lấy model của tháng trước
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
        database.executemany(sql = ''' INSERT INTO `production_output` (line_id, production_date, model_name, OK_qty, NG_qty)
                                        VALUES ((SELECT line_id FROM production_lines WHERE line_name = :line_name), 
                                                :production_date, :model_name, :OK_qty, :NG_qty)
                                                ON DUPLICATE KEY UPDATE output_id = output_id;''', params_list = params_list)
    except Exception as e:
        print(f"Error inserting data into database: {e}")
        sys.exit(1)
import pandas as pd
import sys
import os
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Maintenance.scan_qrcode import Scan_record_process


scanner = Scan_record_process()
list_dict = []

dict = [
{'machine_code':'MCG152305', 'record_name':'MCG152305_Dispenser_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 0, 'end_page': 1},
{'machine_code':'MCG152018', 'record_name':'MCG152018_Dispenser_TAD-200_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 2, 'end_page': 2},
{'machine_code':'MCG152423', 'record_name':'MCG152423_Winding_Machine_GORMAN_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 3, 'end_page': 4},
{'machine_code':'ZAC-021', 'record_name':'ZAC-021_BO_DEM_SO_VONG_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 5, 'end_page': 5},
{'machine_code':'MCG152428', 'record_name':'MCG152428_Winding_Machine_GORMAN_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 6, 'end_page': 7},
{'machine_code':'ZAC-024', 'record_name':'ZAC-024_BO_DEM_VONG_QUAN_DAY_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 8, 'end_page': 8},
{'machine_code':'688731', 'record_name':'688731_WIRE_WINDING_MACHINE_PE2_Z02_VINH_2026-08-24.pdf', 'start_page': 9, 'end_page': 10},
{'machine_code':'ZAC-014', 'record_name':'ZAC-014_BO_DEM_VONG_QUAN_DAY_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 11, 'end_page': 11},
{'machine_code':'MCG152518', 'record_name':'MCG152518_Winding_Machine_GORMAN_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 12, 'end_page': 13},
{'machine_code':'MCG151986', 'record_name':'MCG151986_Oven_DNF64-2_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 14, 'end_page': 15},
{'machine_code':'ACS-039', 'record_name':'ACS-039_DONG_HO_TU_SAY_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 16, 'end_page': 16},
{'machine_code':'MCG152119', 'record_name':'MCG152119_Oven_DKM-400_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 17, 'end_page': 18},
{'machine_code':'ACS-041', 'record_name':'ACS-041_DONG_HO_TU_SAY_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 19, 'end_page': 19},
{'machine_code':'MCG151926', 'record_name':'MCG151926_Transformer_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 20, 'end_page': 20},
{'machine_code':'MCG153345', 'record_name':'MCG153345_Oven_DKM-400_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 21, 'end_page': 22},
{'machine_code':'MCG152352', 'record_name':'MCG152352_Transformer_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 23, 'end_page': 23},
{'machine_code':'ACS-038', 'record_name':'ACS-038_DONG_HO_TU_SAY_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 24, 'end_page': 24},
{'machine_code':'1604210', 'record_name':'1604210_DRYING_MACHINE_(OVEN)-BINDER_FD-115_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 25, 'end_page': 26},
{'machine_code':'688753', 'record_name':'688753_ZCT_AUTOMATIC_TESTING_MC_PE2_Z02_VINH_2026-08-24.pdf', 'start_page': 27, 'end_page': 29},
{'machine_code':'MCG152855', 'record_name':'MCG152855_Auto_Matic_Voltage_Stabilizer_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 30, 'end_page': 30},
{'machine_code':'MCG152318', 'record_name':'MCG152318_Variable_Transformer_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 31, 'end_page': 31},
{'machine_code':'MCG152331', 'record_name':'MCG152331_Transformer_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 32, 'end_page': 32},
{'machine_code':'MCG152307', 'record_name':'MCG152307_Variable_Transformer_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 33, 'end_page': 33},
{'machine_code':'MCG152330', 'record_name':'MCG152330_Variable_Transformer_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 34, 'end_page': 34},
{'machine_code':'MCG152866', 'record_name':'MCG152866_Transformer_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 35, 'end_page': 35},
{'machine_code':'MCG152286', 'record_name':'MCG152286_Transformer_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 36, 'end_page': 36},
{'machine_code':'ZAJ-023', 'record_name':'ZAJ-023_AUTO_JIG_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 37, 'end_page': 37},
{'machine_code':'ZAJ-024', 'record_name':'ZAJ-024_AUTO_JIG_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 38, 'end_page': 38},
{'machine_code':'MCG152361', 'record_name':'MCG152361_Conveyor_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 39, 'end_page': 39},
{'machine_code':'LHP-006', 'record_name':'LHP-006_THIET_BI_CHAO_SAY_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 40, 'end_page': 40},
{'machine_code':'MCG153673', 'record_name':'MCG153673_Electric_testing_jig_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 41, 'end_page': 41},
{'machine_code':'MCG152867', 'record_name':'MCG152867_Electric_testing_jig_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 42, 'end_page': 42},
{'machine_code':'ZJ-059', 'record_name':'ZJ-059_JIG_KIEM_DIEN_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 43, 'end_page': 43},
{'machine_code':'ZJ-058', 'record_name':'ZJ-058_JIG_KIEM_DIEN_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 44, 'end_page': 44},
{'machine_code':'ZJ-061', 'record_name':'ZJ-061_JIG_KIEM_DIEN_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 45, 'end_page': 45},
{'machine_code':'MCG210087', 'record_name':'MCG210087_Cutting_machine_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 46, 'end_page': 47},
{'machine_code':'ZAC-019', 'record_name':'ZAC-019_BO_DEM_SO_VONG_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 48, 'end_page': 48},
{'machine_code':'ZAC-023', 'record_name':'ZAC-023_BO_DEM_VONG_QUAN_DAY_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 49, 'end_page': 49},
{'machine_code':'ZAC-004', 'record_name':'ZAC-004_BO_DEM_VONG_QUAN_DAY_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 50, 'end_page': 50},
{'machine_code':'ZAC-013', 'record_name':'ZAC-013_BO_DEM_VONG_QUAN_DAY_PE2_Z02_VINH_2026-08-28.pdf', 'start_page': 51, 'end_page': 51}
]

for item in dict:
    machine_code = item["machine_code"]
    record_name = item["record_name"]
    start_page = item["start_page"]
    end_page = item["end_page"]
    if start_page is not None:
        scanner.split_pdf(input_file=r"//172.30.73.156/share/28082026132318.pdf",
                          start=start_page, end=end_page, output_file=f"X:\\Scan_Result_X\\{record_name}")
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
{"machine_code":"2130288", "record_name":"2130288_ROUNDING_+_DISPENSING_PE1_NFC4_THANH_2026-01-05.pdf", "start_page": 34, "end_page": 36},
{"machine_code":"2130291", "record_name":"2130291_ASSEMBLING_+JET_DISPENSER_PE1_NFC4_THANH_2026-01-05.pdf", "start_page": 26, "end_page": 27},
{"machine_code":"2142312-1/13", "record_name":"2142312-113_WINDING_MACHINE_for_2142312_PE1_NFC5_THANH_2026-01-06.pdf", "start_page": 28, "end_page": 29},
{"machine_code":"2142312-6/13", "record_name":"2142312-613_ASSEMBLING_AND_FORMING__for_2142312_PE1_NFC5_THANH_2026-01-06.pdf", "start_page": 40, "end_page": 42},
{"machine_code":"2142312-7/13", "record_name":"2142312-713_ASSEMBLY_T_CORE_AND_JET_DISPENSER_for_2142312_PE1_NFC5_THANH_2026-01-06.pdf", "start_page": 2, "end_page": 3},
{"machine_code":"2166321-1/4", "record_name":"2166321-14_WINDING_MACHINE_for_2166321_PE1_NFC4_THANH_2026-01-05.pdf", "start_page": 4, "end_page": 5},
{"machine_code":"2224560", "record_name":"2224560_FORMING_MACHINE_PE1_NFC6_THANH_2026-01-07.pdf", "start_page": 37, "end_page": 39},
{"machine_code":"2224561", "record_name":"2224561_WINDING_MACHINE_PE1_NFC6_THANH_2026-01-07.pdf", "start_page": 6, "end_page": 7},
{"machine_code":"2224562", "record_name":"2224562_GLUE,T-CORE_ASS'Y__MACHINE_PE1_NFC6_THANH_2026-01-07.pdf", "start_page": 8, "end_page": 9},
{"machine_code":"MCG150005", "record_name":"MCG150005_Forming_PE1_WSPE1_THANH_2026-01-19.pdf" , "start_page": 0, "end_page": 1},
{"machine_code":"MCG150011", "record_name":"MCG150011_Forming_PE1_WSPE1_THANH_2026-01-19.pdf" , "start_page": 10, "end_page": 11},
{"machine_code":"MCG150032", "record_name":"MCG150032_Forming_PE1_WSPE1_THANH_2026-01-19.pdf" , "start_page": 12, "end_page": 13},
{"machine_code":"MCG150048", "record_name":"MCG150048_Forming_PE1_WSPE1_THANH_2026-01-19.pdf" , "start_page": 14, "end_page": 15},
{"machine_code":"MCG150273", "record_name":"MCG150273_Forming_PE1_WSPE1_THANH_2026-01-19.pdf" , "start_page": 16, "end_page": 17},
{"machine_code":"MCG150297", "record_name":"MCG150297_Forming_PE1_WSPE1_THANH_2026-01-17.pdf" , "start_page": 18, "end_page": 19},
{"machine_code":"MCG153626", "record_name":"MCG153626_Forming_PE1_WSPE1_THANH_2026-01-19.pdf" , "start_page": 32, "end_page": 33},
{"machine_code":"MCG153647", "record_name":"MCG153647_Forming_PE1_WSPE1_THANH_2026-01-19.pdf" , "start_page": 20, "end_page": 21},
{"machine_code":"MCG180040", "record_name":"MCG180040_Forming_PE1_WSPE1_THANH_2026-01-19.pdf" , "start_page": 22, "end_page": 23},
{"machine_code":"MCG190210", "record_name":"MCG190210_Forming_PE1_WSPE1_THANH_2026-01-19.pdf" , "start_page": 24, "end_page": 25},
{"machine_code":"MCG230122", "record_name":"MCG230122_Forming_PE1_WSPE1_THANH_2026-01-19.pdf" , "start_page": 30, "end_page": 31}]

for item in dict:
    machine_code = item["machine_code"]
    record_name = item["record_name"]
    start_page = item["start_page"]
    end_page = item["end_page"]
    if start_page is not None:
        scanner.split_pdf(input_file=r"X:\Scan\10082026142124.pdf",
                          start=start_page, end=end_page, output_file=f"X:\\Scan_Result_X\\{record_name}")
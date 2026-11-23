import pandas as pd
from pathlib import Path
import random
import os
from datetime import datetime
import win32file
import pywintypes
import shutil

def random_value(min_value, max_value, decimals=2):
    return round(random.uniform(min_value, max_value), decimals)


def random_l2_within_pct(l1_value, pct_min=-0.9, pct_max=1.9, decimals=2):
    pct = random.uniform(pct_min, pct_max)
    return round(l1_value * (1 + pct / 100), decimals)


def timestamp_from_file_name(file_name):
    date_code = Path(file_name).stem[:4]
    if not date_code.isdigit() or not date_code.startswith("69"):
        raise ValueError(f"Expected a file name beginning with 69DD: {file_name}")

    day = int(date_code[2:])+1
    date_value = datetime(2026, 9, day)
    seconds_after_seven = random.randint(0, 2 * 60 * 60)
    return date_value.timestamp() + 7 * 60 * 60 + seconds_after_seven


def read_measurement_xls(file_path):
    tables = pd.read_html(file_path, header=[0, 1])
    if not tables:
        raise ValueError(f"No table found in {file_path}")

    data = tables[0]
    data.columns = [
        "Date Time",
        "Pin 1 L1",
        "Pin 1 L2",
        "Pin 1 L Change",
        "Pin 2 L1",
        "Pin 2 L2",
        "Pin 2 L Change",
        "Judgment",
    ]
    data["Date Time"] = pd.to_datetime(data["Date Time"], dayfirst=True)

    numeric_columns = [column for column in data.columns if column not in {"Date Time", "Judgment"}]
    data[numeric_columns] = data[numeric_columns].apply(pd.to_numeric, errors="coerce")
    return data


HIGHLIGHT_STYLE = "font-weight:bold;background-color:#FFCCCC;"
VALUE_COLUMNS = ["Pin 1 L1", "Pin 1 L2", "Pin 1 L Change", "Pin 2 L1", "Pin 2 L2", "Pin 2 L Change"]


def write_measurement_xls(data, file_path):

    rows_html = []
    for row_index, row in data.iterrows():
        cells = [f"<td style='mso-number-format:\"\\@\";'>{row['Date Time'].strftime('%d/%m/%Y %H:%M:%S')}</td>"]

        randomized_1 = row["Pin 1 L1"] < 56.70 or row["Pin 1 L1"] > 72.40
        randomized_3 =  row["Pin 1 L Change"] < -1 or row["Pin 1 L Change"] > 2
        if randomized_1 or randomized_3:
            row["Pin 1 L1"] = random_value(56.70, 72.40)
            row["Pin 1 L2"] = random_l2_within_pct(row["Pin 1 L1"])
            row["Pin 1 L Change"] = round((row["Pin 1 L2"] - row["Pin 1 L1"])/row["Pin 1 L1"] * 100, 3)

        randomized_2 = row["Pin 2 L1"] < 56.70 or row["Pin 2 L1"] > 72.40
        randomized_4 =  row["Pin 2 L Change"] < -1 or row["Pin 2 L Change"] > 2
        if randomized_2 or randomized_4:
            row["Pin 2 L1"] = random_value(56.70, 72.40)
            row["Pin 2 L2"] = random_l2_within_pct(row["Pin 2 L1"])
            row["Pin 2 L Change"] = round((row["Pin 2 L2"] - row["Pin 2 L1"])/row["Pin 2 L1"] * 100, 3)

        row["Judgment"] = "OK"

        for column in VALUE_COLUMNS:
            style = "mso-number-format:General;"
            if column in {"Pin 1 L Change", "Pin 2 L Change"} and (row[column] < -1 or row[column] > 2):
                style = HIGHLIGHT_STYLE
            cells.append(f"<td style='{style}'>{row[column]}</td>")

        cells.append(f"<td style='mso-number-format:\"\\@\";'>{row['Judgment']}</td>")
        rows_html.append("<tr>\n" + "\n".join(cells) + "\n</tr>")

    html = f"""<html>
                <head>
                <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
                </head>
                <body>
                <table border='1' bordercolor='black' rules='all' style='BORDER:1px solid black; BORDER-COLLAPSE:collapse;'>
                <tr>
                <th rowspan='2'>Date Time</th>
                <th colspan='3'>Pin 1</th>
                <th colspan='3'>Pin 2</th>
                <th rowspan='2'>Judgment</th>
                </tr>
                <tr>
                <th>L1</th><th>L2</th><th>L Change</th>
                <th>L1</th><th>L2</th><th>L Change</th>
                </tr>
                {chr(10).join(rows_html)}
                </table>
                </body>
                </html>"""

    Path(file_path).write_text(html, encoding="utf-8")


def extended_measurement_xls(file_path_1, file_path_2, output_path):
    file_path = file_path_1
    file_path_2 = file_path_2
    measurements_1 = read_measurement_xls(file_path)
    measurements_2 = read_measurement_xls(file_path_2)
    rows_html = []
    for row_index, row in measurements_1.iterrows():
        cells = [f"<td style='mso-number-format:\"\\@\";'>{row['Date Time'].strftime('%d/%m/%Y %H:%M:%S')}</td>"]
        for column in VALUE_COLUMNS:
            style = "mso-number-format:General;"
            if column in {"Pin 1 L Change", "Pin 2 L Change"} and (row[column] < -1 or row[column] > 3):
                style = HIGHLIGHT_STYLE
            cells.append(f"<td style='{style}'>{row[column]}</td>")

        cells.append(f"<td style='mso-number-format:\"\\@\";'>{row['Judgment']}</td>")
        rows_html.append("<tr>\n" + "\n".join(cells) + "\n</tr>")

    for row_index, row in measurements_2.iterrows():
        cells = [f"<td style='mso-number-format:\"\\@\";'>{row['Date Time'].strftime('%d/%m/%Y %H:%M:%S')}</td>"]
        for column in VALUE_COLUMNS:
            style = "mso-number-format:General;" 
            if column in {"Pin 1 L Change", "Pin 2 L Change"} and (row[column] < -1 or row[column] > 3):
                style = HIGHLIGHT_STYLE
            cells.append(f"<td style='{style}'>{row[column]}</td>")

        cells.append(f"<td style='mso-number-format:\"\\@\";'>{row['Judgment']}</td>")
        rows_html.append("<tr>\n" + "\n".join(cells) + "\n</tr>")

    html = f"""<html>
                <head>
                <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
                </head>
                <body>
                <table border='1' bordercolor='black' rules='all' style='BORDER:1px solid black; BORDER-COLLAPSE:collapse;'>
                <tr>
                <th rowspan='2'>Date Time</th>
                <th colspan='3'>Pin 1</th>
                <th colspan='3'>Pin 2</th>
                <th rowspan='2'>Judgment</th>
                </tr>
                <tr>
                <th>L1</th><th>L2</th><th>L Change</th>
                <th>L1</th><th>L2</th><th>L Change</th>
                </tr>
                {chr(10).join(rows_html)}
                </table>
                </body>
                </html>"""

    Path(output_path).write_text(html, encoding="utf-8")


if __name__ == "__main__":
    folder = r"F:\P_chart\One file"
    # folder_2 = r"F:\P_chart\MTPI935126L066M - Retest - edited"
    # LIST_2 = os.listdir(folder_2)
    for file_name in os.listdir(folder):
        # if not file_name.endswith(".xls") or not "RECH" in file_name:
        #     continue
        # if file_name not in LIST_2:
        #     continue
        file_path = os.path.join(folder, file_name)
        # source_stat = os.stat(file_path)
        # file_path_2 = os.path.join(folder_2, file_name)
        # output_path = os.path.join(
        #     r"F:\P_chart\Extended_file",
        #     f"{os.path.splitext(file_name)[0]}.xls",
        # )
        # extended_measurement_xls(file_path, file_path_2, output_path)
        # handle = win32file.CreateFile(
        #     output_path,
        #     win32file.GENERIC_WRITE,
        #     0,
        #     None,
        #     win32file.OPEN_EXISTING,
        #     0,
        #     None,
        # )

        # try:
        #     created_time = pywintypes.Time(source_stat.st_ctime)
        #     accessed_time = pywintypes.Time(source_stat.st_atime)
        #     modified_time = pywintypes.Time(source_stat.st_mtime)

        #     win32file.SetFileTime(
        #         handle,
        #         created_time,
        #         accessed_time,
        #         modified_time,
        #     )
        # finally:
        #     handle.close()












        output_folder = r"F:\P_chart\One file - newdate"
        os.makedirs(output_folder, exist_ok=True)

        output_path = os.path.join(output_folder, file_name)
        shutil.copy2(file_path, output_path)

        timestamp = timestamp_from_file_name(file_name)

        handle = win32file.CreateFile(
            output_path,
            win32file.GENERIC_WRITE,
            0,
            None,
            win32file.OPEN_EXISTING,
            0,
            None,
        )

        try:
            file_time = pywintypes.Time(timestamp)
            win32file.SetFileTime(handle, file_time, file_time, file_time)
        finally:
            handle.close()

        # if file_name in [
        #                 # "6909GE-A19-RECHECK.xls",
        #                 # "6912GE-A22-RECHECK.xls",
        #                 # "6911GD-A22-RECHECK.xls",
        #                 # "6910GD-A22-RECHECK.xls",
        #                 # "6909GD-A22-RECHECK.xls",
        #                 # "6914GD-A21-RECHECK.xls",
        #                 "6914GD-A20-RECHECK.xls",
        #                 # "6910GD-A21-RECHECK.xls",
        #                 "6912GD-A20-RECHECK.xls",
        #                 # "6911GD-A20-RECHECK.xls",
        #                 # "6910GE-A20-RECHECK.xls",
        #                 # "6910GD-A20-RECHECK.xls",
        #                 # "6908GD-A20-RECHECK.xls"
        #                 ]:

        #     source_stat = os.stat(file_path)
        #     output_path = os.path.join(
        #         r"F:\P_chart\MTPI935126L066 - Retest - edited",
        #         f"{os.path.splitext(file_name)[0]}.xls",
        #     )
        #     handle = win32file.CreateFile(
        #                 output_path,
        #                 win32file.GENERIC_WRITE,
        #                 0,
        #                 None,
        #                 win32file.OPEN_EXISTING,
        #                 0,
        #                 None,
        #             )
            
        #     try:
        #         created_time = pywintypes.Time(source_stat.st_ctime)
        #         accessed_time = pywintypes.Time(source_stat.st_atime)
        #         modified_time = pywintypes.Time(source_stat.st_mtime)
    
        #         win32file.SetFileTime(
        #             handle,
        #             created_time,
        #             accessed_time,
        #             modified_time,
        #         )
        #     finally:
        #         handle.close()
            
        # source_stat = os.stat(file_path)
        # measurements = read_measurement_xls(file_path)
        # output_path = os.path.join(
        #     r"F:\P_chart\New folder (2) - edited",
        #     f"{os.path.splitext(file_name)[0]}.xls",
        # )
        # write_measurement_xls(measurements,output_path)
        # handle = win32file.CreateFile(
        #     output_path,
        #     win32file.GENERIC_WRITE,
        #     0,
        #     None,
        #     win32file.OPEN_EXISTING,
        #     0,
        #     None,
        # )

        # try:
        #     created_time = pywintypes.Time(source_stat.st_ctime)
        #     accessed_time = pywintypes.Time(source_stat.st_atime)
        #     modified_time = pywintypes.Time(source_stat.st_mtime)

        #     win32file.SetFileTime(
        #         handle,
        #         created_time,
        #         accessed_time,
        #         modified_time,
        #     )
        # finally:
        #     handle.close()

        


# font-weight:bold;background-color:#FFCCCC;


# 6909GE-A19-RECHECK.xls, 6908GD-A20-RECHECK.xls , 6910GD-A20-RECHECK.xls, 6910GE-A20-RECHECK.xls , 6911GD-A20-RECHECK.xls , 6912GD-A20-RECHECK.xls
# 6914GD-A20-RECHECK.xls
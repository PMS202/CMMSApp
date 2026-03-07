
import calendar
import datetime as dt
import pandas as pd

# import os
# import sys
# import NG_data2 as NG
# import FG_data2 as FG
# import Losstime2 as LT

# sys.path.append(r'C:\Users\2173452100291\Documents\OEE_project')

# from Database.MariaDB import Database_process
import Calculation.Losstime2 as LT
import Calculation.FG_data2 as FG
import Calculation.NG_data2 as NG


Net_available_runtime = {8: 425, 9.5: 515, 12: 635, 16: 850, 24: 580}


class OEE_result():
    def __init__(self, cycle_time=None):
        super().__init__()
        self.NG_file = None
        self.NG_Coil_Sheetname = None
        self.date_col_NG_Coil = None
        self.begin_NG_coil = None
        self.end_NG_coil = None
        self.NG_Molding_Sheetname = None
        self.date_col_NG_Molding = None
        self.begin_NG_Molding = None
        self.end_NG_Molding = None

        self.month = None
        self.year = None

        self.FG_file = None
        self.FG_sheet_name = None
        self.FG_date_col = None
        self.FG_line_col = None

        self.Molding_lt_sheet_name = None
        self.Coil_lt_sheet_name = None
        self.lt_date_col = None
        self.cycle_time_list = cycle_time
        self.cycle_time_list.columns = [
            "line_name", "machine_name", "machine_id", "cycletime"]
        self.df_dic = {
            "Day": [],
            "Working_shift": [],
            "Net_avaiable_runtime": [],
            "Downtime": [],
            "Other_stop": [],
            "Runtime": [],
            "FGs": [],
            "Defect": [],
            "A": [],
            "P": [],
            "Q": [],
            "OEE": []
        }
        self.df_dic_month = {
            "Working_shift": [0],
            "Net_avaiable_runtime": [0],
            "Downtime": [0],
            "Other_stop": [0],
            "Runtime": [0],
            "FGs": [0],
            "Defect": [0],
            "A": [0],
            "P": [0],
            "Q": [0],
            "OEE": [0]
        }

    def default_setting(self):
        self.default_NG_Coil_Sheetname = "3-Summary Coil"
        self.default_date_col_NG_Coil = 0
        self.default_begin_NG_coil = 2
        self.default_end_NG_coil = 21
        self.default_NG_Molding_Sheetname = "4-Summary Final"
        self.default_date_col_NG_Molding = 0
        self.default_begin_NG_Molding = 2
        self.default_end_NG_Molding = 20

        self.default_month = [12 if dt.datetime.now(
        ).month == 1 else dt.datetime.now().month - 1][0]
        self.default_year = [dt.datetime.now(
        ).year - 1 if dt.datetime.now().month == 1 else dt.datetime.now().year][0]

        self.default_FG_sheet_name = f"Molding {calendar.month_abbr[self.default_month]}-{self.default_year-2000} "
        # self.default_FG_sheet_name = "Molding Mar-25 "
        self.default_FG_date_col = 0
        self.default_FG_line_col = 1

        self.default_Molding_lt_sheet_name = "2.MOLDING"
        self.default_Coil_lt_sheet_name = "1.COIL"
        self.default_lt_date_col = 0

    def OEE_cal_result(self, NG_file, FG_file):

        self.default_setting()
        if NG_file is None and FG_file is None:
            return None
        self.NG_file = NG_file
        self.FG_file = FG_file

        if self.NG_Coil_Sheetname is None:
            self.NG_Coil_Sheetname = self.default_NG_Coil_Sheetname

        if self.date_col_NG_Coil is None:
            self.date_col_NG_Coil = self.default_date_col_NG_Coil

        if self.begin_NG_coil is None:
            self.begin_NG_coil = self.default_begin_NG_coil

        if self.end_NG_coil is None:
            self.end_NG_coil = self.default_end_NG_coil

        if self.NG_Molding_Sheetname is None:
            self.NG_Molding_Sheetname = self.default_NG_Molding_Sheetname

        if self.date_col_NG_Molding is None:
            self.date_col_NG_Molding = self.default_date_col_NG_Molding

        if self.begin_NG_Molding is None:
            self.begin_NG_Molding = self.default_begin_NG_Molding

        if self.end_NG_Molding is None:
            self.end_NG_Molding = self.default_end_NG_Molding

        if self.month is None:
            self.month = self.default_month

        if self.year is None:
            self.year = self.default_year

        if self.FG_sheet_name is None:
            self.FG_sheet_name = self.default_FG_sheet_name

        if self.FG_date_col is None:
            self.FG_date_col = self.default_FG_date_col

        if self.FG_line_col is None:
            self.FG_line_col = self.default_FG_line_col

        if self.Molding_lt_sheet_name is None:
            self.Molding_lt_sheet_name = self.default_Molding_lt_sheet_name

        if self.Coil_lt_sheet_name is None:
            self.Coil_lt_sheet_name = self.default_Coil_lt_sheet_name

        if self.lt_date_col is None:
            self.lt_date_col = self.default_lt_date_col

        NG_Coil, FG_Coil = NG.filter_data_NG(self.NG_file, self.NG_Coil_Sheetname, self.date_col_NG_Coil,
                                             self.begin_NG_coil, self.end_NG_coil, self.month, self.year)
        NG_Molding = NG.filter_data_NG(self.NG_file, self.NG_Molding_Sheetname, self.date_col_NG_Molding,
                                       self.begin_NG_Molding, self.end_NG_Molding, self.month, self.year)
        FG_Molding, working_time = FG.filter_data_FG(self.FG_file, self.FG_sheet_name, self.FG_date_col,
                                                     self.FG_line_col, self.month, self.year)
        Losstime_Mold = LT.Losstime(self.FG_file, self.Molding_lt_sheet_name, self.lt_date_col,
                                    self.month, self.year)
        Losstime_Coil = LT.Losstime(self.FG_file, self.Coil_lt_sheet_name, self.lt_date_col,
                                    self.month, self.year)
        total_lines = [f"F0{i}" if i < 10 else f"F{i}" for i in range(1, 25)]
        list_df_molding = []
        list_df_coil = []
        list_df_molding_monthly = []
        list_df_coil_monthly = []
        for line in total_lines:
            result_df_modling = pd.DataFrame(self.df_dic)
            result_df_coil = result_df_modling.copy()
            result_df_modling_monthly = pd.DataFrame(self.df_dic_month)
            result_df_coil_monthly = result_df_modling_monthly.copy()
            NG_Coil_line = NG_Coil[line]
            NG_Molding_line = NG_Molding[line]
            FG_Molding_line = FG_Molding[line]
            FG_Coil_line = FG_Coil[line]
            Losstime_Mold_line = Losstime_Mold[line]
            Losstime_Coil_line = Losstime_Coil[line]
            NG_Coil_line.fillna(0, inplace=True)
            NG_Molding_line.fillna(0, inplace=True)
            FG_Molding_line.fillna(0, inplace=True)
            FG_Coil_line.fillna(0, inplace=True)
            Losstime_Mold_line.fillna(0, inplace=True)
            Losstime_Coil_line.fillna(0, inplace=True)
            cycle_time_molding = self.cycle_time_list[(self.cycle_time_list["line_name"] == line) & (
                self.cycle_time_list["machine_name"] == "Molding")]["cycletime"].values[0]
            cycle_time_coil = self.cycle_time_list[(self.cycle_time_list["line_name"] == line) & (
                self.cycle_time_list["machine_name"] == "Coil")]["cycletime"].values[0]

            result_df_modling["Day"] = FG_Molding_line.index
            result_df_modling.set_index("Day", inplace=True)

            result_df_modling["Working_shift"] = working_time[line]
            result_df_modling["Net_avaiable_runtime"] = result_df_modling["Working_shift"].map(
                Net_available_runtime)
            result_df_modling["Downtime"] = Losstime_Mold_line
            result_df_modling["FGs"] = FG_Molding_line
            result_df_modling["Defect"] = NG_Molding_line
            result_df_modling.fillna(0, inplace=True)
            result_df_modling["Runtime"] = result_df_modling["Net_avaiable_runtime"] - \
                result_df_modling["Downtime"]
            result_df_modling["A"] = result_df_modling["Runtime"] / \
                (result_df_modling["Net_avaiable_runtime"]+0.000001)
            result_df_modling["P"] = result_df_modling["FGs"] / \
                ((60/cycle_time_molding) *
                (result_df_modling["Runtime"]+0.000001))
            result_df_modling["Q"] = result_df_modling["FGs"] / \
                (result_df_modling["FGs"] +
                result_df_modling["Defect"] + 0.000001)
            result_df_modling["OEE"] = result_df_modling["A"] * \
                result_df_modling["P"] * result_df_modling["Q"]
            result_df_modling[result_df_modling < 0] = 0
            result_df_modling["temp"] = result_df_modling["Working_shift"]
            result_df_modling["temp"] = (result_df_modling["temp"] / result_df_modling["temp"]).fillna(0)
            result_df_modling = result_df_modling.mul(result_df_modling.iloc[:, -1], axis=0)
            result_df_modling.drop(columns=["temp"], inplace=True)

            result_df_modling_monthly["Working_shift"] = result_df_modling["Working_shift"].sum()
            result_df_modling_monthly["Net_avaiable_runtime"] = result_df_modling["Net_avaiable_runtime"].sum()
            result_df_modling_monthly["Downtime"] = result_df_modling["Downtime"].sum()
            result_df_modling_monthly["FGs"] = result_df_modling["FGs"].sum()
            result_df_modling_monthly["Defect"] = result_df_modling["Defect"].sum()
            result_df_modling_monthly["Runtime"] = result_df_modling["Runtime"].sum()
            result_df_modling_monthly["A"] = result_df_modling_monthly["Runtime"] / \
                (result_df_modling_monthly["Net_avaiable_runtime"]+0.000001)
            result_df_modling_monthly["P"] = result_df_modling_monthly["FGs"] / \
                ((60/cycle_time_molding) *
                 (result_df_modling_monthly["Runtime"]+0.000001))
            result_df_modling_monthly["Q"] = result_df_modling_monthly["FGs"] / \
                (result_df_modling_monthly["FGs"] +
                 result_df_modling_monthly["Defect"] + 0.000001)
            result_df_modling_monthly["OEE"] = result_df_modling_monthly["A"] * \
                result_df_modling_monthly["P"] * result_df_modling_monthly["Q"]
            result_df_modling = result_df_modling.round(2)
            result_df_modling_monthly = result_df_modling_monthly.round(2)
            
            list_df_molding.append((f"{line}", result_df_modling))
            list_df_molding_monthly.append(
                (f"{line}", result_df_modling_monthly))
            result_df_modling = None
            result_df_modling_monthly = None

            result_df_coil["Day"] = FG_Coil_line.index
            result_df_coil.set_index("Day", inplace=True)

            result_df_coil["Working_shift"] = working_time[line]
            result_df_coil["Net_avaiable_runtime"] = result_df_coil["Working_shift"].map(
                Net_available_runtime)
            result_df_coil["Downtime"] = Losstime_Coil_line
            result_df_coil["FGs"] = FG_Coil_line
            result_df_coil["Defect"] = NG_Coil_line
            result_df_coil.fillna(0, inplace=True)
            result_df_coil["Runtime"] = result_df_coil["Net_avaiable_runtime"] - \
                result_df_coil["Downtime"]
            result_df_coil["A"] = result_df_coil["Runtime"] / \
                (result_df_coil["Net_avaiable_runtime"]+0.000001)
            result_df_coil["P"] = result_df_coil["FGs"] / \
                ((60/cycle_time_coil)*(result_df_coil["Runtime"]+0.000001))
            result_df_coil["Q"] = result_df_coil["FGs"] / \
                (result_df_coil["FGs"] + result_df_coil["Defect"] + 0.000001)
            result_df_coil["OEE"] = result_df_coil["A"] * \
                result_df_coil["P"] * result_df_coil["Q"]
            
            result_df_coil["temp"] = result_df_coil["Working_shift"]
            result_df_coil["temp"] = ((result_df_coil["temp"] / result_df_coil["temp"])*(result_df_coil["FGs"] / result_df_coil["FGs"])).fillna(0)
            result_df_coil = result_df_coil.mul(result_df_coil.iloc[:, -1], axis=0)
            result_df_coil.drop(columns=["temp"], inplace=True)
            
            result_df_coil_monthly["Working_shift"] = int(
                result_df_coil["Working_shift"].sum())
            result_df_coil_monthly["Net_avaiable_runtime"] = result_df_coil["Net_avaiable_runtime"].sum(
            )
            result_df_coil_monthly["Downtime"] = result_df_coil["Downtime"].sum(
            )
            result_df_coil_monthly["FGs"] = result_df_coil["FGs"].sum()
            result_df_coil_monthly["Defect"] = result_df_coil["Defect"].sum()
            result_df_coil_monthly["Runtime"] = result_df_coil["Runtime"].sum()
            result_df_coil_monthly["A"] = result_df_coil_monthly["Runtime"] / \
                (result_df_coil_monthly["Net_avaiable_runtime"]+0.000001)
            result_df_coil_monthly["P"] = result_df_coil_monthly["FGs"] / \
                ((60/cycle_time_coil) *
                 (result_df_coil_monthly["Runtime"]+0.000001))
            result_df_coil_monthly["Q"] = result_df_coil_monthly["FGs"] / \
                (result_df_coil_monthly["FGs"] +
                 result_df_coil_monthly["Defect"] + 0.000001)
            result_df_coil_monthly["OEE"] = result_df_coil_monthly["A"] * \
                result_df_coil_monthly["P"] * result_df_coil_monthly["Q"]
            result_df_coil = result_df_coil.round(2)
            result_df_coil_monthly = result_df_coil_monthly.round(2)
            list_df_coil.append((f"{line}", result_df_coil))
            list_df_coil_monthly.append((f"{line}", result_df_coil_monthly))
            result_df_coil = None
            result_df_coil_monthly = None
        return list_df_molding, list_df_molding_monthly, list_df_coil, list_df_coil_monthly


# db = Database_process()
# info = db.query(sql='''SELECT `Production_Lines`.line_name,`Machines`.machine_name,`Machine_CycleTime`.machine_id,`Machine_CycleTime`.cycletime
#                                     FROM `Machine_CycleTime`
#                                     JOIN `Machines` ON `Machines`.machine_id = `Machine_CycleTime`.machine_id
#                                     JOIN `Production_Lines` ON `Production_Lines`.line_id = `Machines`.line_id;''', params=None)

# oe = OEE_result(info)
# oe.default_setting()
# oe.month = 1
# oe.FG_sheet_name = "Molding Jan-25 "
# list_df_molding, list_df_molding_monthly, list_df_coil, list_df_coil_monthly = oe.OEE_cal_result(NG_file=r"Excel_data\Tong hop so luong NG OK 1-25.xls",
#                                                                                                  FG_file=r"Excel_data\Molding daily Jan-25.xlsx")
# print(list_df_coil, end="\n")
# 

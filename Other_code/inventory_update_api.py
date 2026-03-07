from fastapi import FastAPI
from fastapi.responses import JSONResponse
import pandas as pd
from datetime import datetime
import os
import sys
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process

app = FastAPI(
    title="Inventory Update API"
)

link_dict = {
    "XUAT IE1": r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat IE1(HT) .xlsx",
    "XUAT IE2": r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat IE2(HT).xlsx",
    "XUAT IE3": r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat (HT).xlsx",
    "XUAT IE4": r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat IE4(HT).xlsx",
    "XUAT PEM": r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat PEM (HT).xlsx",
    "XUAT PI":  r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat PI(HT).xlsx",
    "TEV suggestion - Format(PE3)": r"\\tev-1\TEV_ushare\2173451623029\1. Spare Part Controlling (PE3).xlsm",
    "TEV suggestion - Format": r"\\tev-1\TEV_ushare\dohung\Bảo trì CXA\Spare Part Controlling - v1.0.xlsx"
}
department_index_dict = {
    "XUAT IE1": 1,
    "XUAT IE2": 2,
    "XUAT IE3": 3,
    "XUAT IE4": 4,
    "XUAT PEM": 5,
    "XUAT PI":  6,
    "TEV suggestion - Format(PE3)": 3,
    "TEV suggestion - Format": 1
}

def load_inventory_data():
    df_total = pd.DataFrame()
    error_dict = {}

    for sheet_name, file_path in link_dict.items():
        try:
            if sheet_name != "TEV suggestion - Format(PE3)" and sheet_name != "TEV suggestion - Format":
                if sheet_name == "XUAT IE3":
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet_name,
                        skiprows=4,
                        header=None,
                        usecols=[1, 9, 14]
                    )
                    df = df.rename(columns={
                        1: "code",
                        9: "current_stock",
                        14: "waiting_receive"
                    })
                else:
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet_name,
                        skiprows=4,
                        header=None,
                        usecols=[1, 9, 11]
                    )
                    df = df.rename(columns={
                        1: "code",
                        9: "current_stock",
                        11: "waiting_receive"
                    })

                df = df.dropna(subset=["code"])
                df = df[df["code"].astype(str).str.len() > 0]
                df["code"] = df["code"].astype(str).str.strip()
                df["department_id"] = department_index_dict.get(sheet_name, None)
                df_total = pd.concat([df_total, df], ignore_index=True, sort=False)
            else:
                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    skiprows=4,
                    header=None,
                    usecols=[3,16]
                )
                df = df.rename(columns={
                    3: "code",
                    16: "current_stock"
                })
                df = df[df["code"].astype(str).str[0].isin(["8", "9"])]
                df["code"] = df["code"].astype(int).astype(str).str.strip()
                df = df.dropna(subset=["code"])
                df = df[df["code"].astype(str).str.len() > 0]
                df["code"] = df["code"].astype(str).str.strip()
                df["waiting_receive"] = 0
                df["department_id"] = department_index_dict.get(sheet_name, None)
                df_total = pd.concat([df_total, df], ignore_index=True, sort=False)

        except Exception as e:
            error_dict[sheet_name] = str(e)
            print(e)
        

    df_total = df_total[df_total["code"].astype(str).str.startswith(("8", "9"))]
    df_total["current_stock"] = pd.to_numeric(df_total["current_stock"], errors="coerce").fillna(0)
    df_total["waiting_receive"] = pd.to_numeric(df_total["waiting_receive"], errors="coerce").fillna(0)
    df_total = (df_total.groupby(["code", "department_id"], as_index=False)[["current_stock", "waiting_receive"]].sum())

    return df_total, error_dict


def update_inventory_table(df):
    DB = Database_process()

    sql = """
    INSERT INTO inventory (
        part_code,
        part_name,
        part_name_vi,
        unit,
        current_stock,
        waiting_receive,
        life_time,
        department_id,
        update_at
    )
    SELECT
        :part_code AS part_code,
        COALESCE(p.part_name, '') AS part_name,
        COALESCE(p.part_name_vi, '') AS part_name_vi,
        COALESCE(p.po_unit, '') AS unit,
        :current_stock AS current_stock,
        :waiting_receive AS waiting_receive,
        0 AS life_time,
        :department_id AS department_id,
        :update_at AS update_at
    FROM purchase p
    WHERE p.part_code = :part_code_find
    ORDER BY p.part_id ASC
    LIMIT 1
    ON DUPLICATE KEY UPDATE
        current_stock = VALUES(current_stock),
        waiting_receive = VALUES(waiting_receive),
        update_at = VALUES(update_at);
    """

    now = datetime.now()
    params_list = [
        {
            "part_code": row.code,
            "current_stock": row.current_stock,
            "waiting_receive": row.waiting_receive,
            "department_id": row.department_id,
            "update_at": now,
            "part_code_find": row.code,
        }
        for _, row in df.iterrows()
    ]
    try:
        DB.executemany(sql, params_list=params_list )
        DB.close()
        return None
    except Exception as e:
        DB.close()
        return e


@app.get("/inventory/update")
def update_inventory():
    df, excel_errors = load_inventory_data()

    if excel_errors:
        return JSONResponse({"status": "error from excel", "detail": excel_errors})
    db_error = update_inventory_table(df)

    if db_error:
        return JSONResponse({"status": "error from database", "detail": db_error})

    return JSONResponse({"status": "finish"})

    
if __name__ == "__main__":
    # import uvicorn
    # uvicorn.run(
    #     "inventory_update_api:app",
    #     host="0.0.0.0",
    #     port=8000,
    #     workers=1,
    #     reload=False
    # )
    df, excel_errors = load_inventory_data()
    print(df[df["code"] == "9000014135"])
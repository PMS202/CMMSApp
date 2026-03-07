# from fastapi import FastAPI
# from fastapi.responses import JSONResponse
import pandas as pd
from datetime import datetime
import os
import sys
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process



link_dict = {
    "UNITPRICE 2025": r"\\172.30.73.156\purchase\1. PROCUREMENT\UNITPRICE.xlsx"
}
department_index_dict = {
    "XUAT IE1": 1,
    "XUAT IE2": 2,
    "XUAT IE3": 3,
    "XUAT IE4": 4,
    "XUAT PEM": 5,
    "XUAT PI":  6,
    "TOTAL": 3
}

def load_inventory_data():
    df_total = pd.DataFrame()
    error_dict = {}

    for sheet_name, file_path in link_dict.items():
        try:
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                skiprows=3,
                header=None,
                usecols=[0,1,2,3,4,5,6,7,9]
            )           
            df = df.rename(columns={
                0: "part_code",
                1: "part_name",
                2: "part_name_vi",
                3: "vendor_code",
                4: "vendor_name",
                5: "po_unit",
                6: "unit_price",
                7: "currency",
                9: "lead_time",
            })          
            df = df.dropna(subset=["part_code"])
            df = df[df["part_code"].astype(str).str.len() > 0]
            df["part_code"] = df["part_code"].astype(str).str.strip()
            df_total = pd.concat([df_total, df], ignore_index=True, sort=False)

        except Exception as e:
            error_dict[sheet_name] = str(e)
        

    df_total = df_total[df_total["part_code"].astype(str).str.startswith(("8", "9"))]
    df_total["vendor_code"] = pd.to_numeric(df_total["vendor_code"], errors="coerce").fillna(0)
    df_total["part_name_vi"] = pd.to_numeric(df_total["part_name_vi"], errors="coerce").fillna(0)

    return df_total, error_dict


def update_inventory_table(df):
    DB = Database_process()

    sql = """
    INSERT IGNORE INTO purchase (
        part_code,
        part_name,
        part_name_vi,
        vendor_code,
        vendor_name,
        po_unit,
        unit_price,
        currency,
        lead_time
    )
    VALUES (
        :part_code,
        :part_name,
        :part_name_vi,
        :vendor_code,
        :vendor_name,
        :po_unit,
        :unit_price,
        :currency,
        :lead_time
    );
    """

    params_list = [
        {
            "part_code": row.part_code,
            "part_name": row.part_name,
            "part_name_vi": row.part_name_vi,
            "vendor_code": row.vendor_code,
            "vendor_name": row.vendor_name,
            "po_unit": row.po_unit,
            "unit_price": row.unit_price,
            "currency": row.currency,
            "lead_time": row.lead_time,
        }
        for _, row in df.iterrows()
    ]
    print(len(params_list))
    try:
        DB.executemany(sql, params_list=params_list )
        DB.close()
        return None
    except Exception as e:
        DB.close()
        return e


# @app.get("/inventory/update")
# def update_inventory():
#     df, excel_errors = load_inventory_data()

#     if excel_errors:
#         return JSONResponse({"status": "error from excel", "detail": excel_errors})
#     db_error = update_inventory_table(df)

#     if db_error:
#         print(db_error)
#         return JSONResponse({"status": "error from database", "detail": db_error})

#     return JSONResponse({"status": "finish"})

    
if __name__ == "__main__":
    # import uvicorn
    # uvicorn.run(
    #     "inventory_update_api:app",
    #     host="0.0.0.0",
    #     port=8000,
    #     workers=1,
    #     reload=False
    # )
    df, _ = load_inventory_data()
    e = update_inventory_table(df)
    if e:
        print(e)
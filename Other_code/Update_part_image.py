# from fastapi import FastAPI
# from fastapi.responses import JSONResponse
import pandas as pd
from datetime import datetime
import os
import sys
from pathlib import Path

# Add project root: c:\Users\2173452100291\Documents\program
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from Database.MariaDB import Database_process
import shutil
import tempfile


link_dict = {
    "part_image": r"C:\Users\2173452100291\Desktop\Excel_to_add_db\part_image_2.xlsx"
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
                usecols=[0,1]
            )           
            df = df.rename(columns={
                0: "machine_code",
                1: "image_link",
            })          
            df = df.dropna(subset=["part_code"])
            df = df[df["part_code"].astype(str).str.len() > 0]
            df["part_code"] = df["part_code"].astype(str).str.strip()
            df_total = pd.concat([df_total, df], ignore_index=True, sort=False)

        except Exception as e:
            error_dict[sheet_name] = str(e)
        

    df_total = df_total[df_total["part_code"].astype(str).str.startswith(("8", "9"))]

    return df_total, error_dict


def update_inventory_table(df):
    DB = Database_process()

    sql = """
    INSERT IGNORE INTO part_image (
        part_code,
        image_link
    )
    VALUES (
        :part_code,
        :image_link
    );
    """

    params_list = [
        {
            "part_code": row.part_code,
            "image_link": row.image_link
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
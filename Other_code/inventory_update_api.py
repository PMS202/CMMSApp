from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, HTMLResponse,  JSONResponse
import pandas as pd
from datetime import datetime
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
    # "XUAT IE1": r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat IE1(HT) .xlsx",
    # "XUAT IE2": r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat IE2(HT).xlsx",
    "XUAT IE3": r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat (HT).xlsx",
    # "XUAT IE4": r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat IE4(HT).xlsx",
    # "XUAT PEM": r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat PEM (HT).xlsx",
    # "XUAT PI":  r"\\172.30.73.156\mcg\QUAN LY CCDC&PTTT\Du Lieu Xuat CCDC&PTTT\Du lieu xuat PI(HT).xlsx",
    "TOTAL": r"C:\Users\2173452100291\Downloads\PE3_DUNG_CU_STOCK_THANG_082026.xlsx",
    # "TEV suggestion - Format": r"\\tev-1\TEV_ushare\dohung\Bảo trì CXA\Spare Part Controlling - v1.0.xlsx",
}
department_index_dict = {
    "XUAT IE1": 1,
    "XUAT IE2": 2,
    "XUAT IE3": 3,
    "XUAT IE4": 4,
    "XUAT PEM": 5,
    "XUAT PI":  6,
    "TOTAL": 3,
    "TEV suggestion - Format": 1
}

def load_inventory_data():
    df_total = pd.DataFrame()
    error_dict = {}

    for sheet_name, file_path in link_dict.items():
        try:
            if sheet_name != "TOTAL" and sheet_name != "TEV suggestion - Format":
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
                    skiprows=9,
                    header=None,
                    usecols=[1,13]
                )
                df = df.rename(columns={
                    1: "code",
                    13: "workshop_stock"
                })
                df = df[df["code"].astype(str).str[0].isin(["8", "9"])]
                df["code"] = df["code"].astype(str).str.strip()
                df = df.dropna(subset=["code"])
                df = df[df["code"].astype(str).str.len() > 0]
                df["code"] = df["code"].astype(str).str.strip()
                df["waiting_receive"] = 0
                df["department_id"] = department_index_dict.get(sheet_name, None)
                df_total = pd.concat([df_total, df], ignore_index=True, sort=False)
                
        except Exception as e:
            error_dict[sheet_name] = str(e)
            print(f"Error processing sheet {sheet_name}: {e}")
        

    if df_total.empty or "code" not in df_total.columns:
        return df_total, error_dict

    df_total = df_total[df_total["code"].astype(str).str.startswith(("8", "9"))]
    df_total["current_stock"] = pd.to_numeric(df_total["current_stock"], errors="coerce").fillna(0)
    df_total["waiting_receive"] = pd.to_numeric(df_total["waiting_receive"], errors="coerce").fillna(0)
    df_total["workshop_stock"] = pd.to_numeric(df_total.get("workshop_stock", 0), errors="coerce").fillna(0)
    df_total = (df_total.groupby(["code", "department_id"], as_index=False)[["current_stock", "workshop_stock", "waiting_receive"]].sum())

    return df_total, error_dict


def update_inventory_table(df):
    DB = Database_process()

    sql = """
    INSERT INTO inventory (
        part_code,
        part_name,
        part_name_vi,
        unit,
        MCG_stock,
        workshop_stock,
        outstanding_orders,
        department_id,
        update_at
    )
    SELECT
        :part_code AS part_code,
        COALESCE(p.part_name, '') AS part_name,
        COALESCE(p.part_name_vi, '') AS part_name_vi,
        COALESCE(p.po_unit, '') AS unit,
        :MCG_stock AS MCG_stock,
        :workshop_stock AS workshop_stock,
        :outstanding_orders AS outstanding_orders,
        :department_id AS department_id,
        :update_at AS update_at
    FROM (SELECT 1) AS dummy
    LEFT JOIN (
        SELECT part_name, part_name_vi, po_unit
        FROM purchase
        WHERE part_code = :part_code_find
        ORDER BY part_id ASC
        LIMIT 1
    ) p ON 1 = 1
    ON DUPLICATE KEY UPDATE
        MCG_stock = VALUES(MCG_stock),
        outstanding_orders = VALUES(outstanding_orders),
        update_at = VALUES(update_at);
    """

    now = datetime.now()
    params_list = [
        {
            "part_code": row.code,
            "MCG_stock": row.current_stock,
            "workshop_stock": row.workshop_stock,
            "outstanding_orders": row.waiting_receive,
            "department_id": row.department_id,
            "update_at": now,
            "part_code_find": row.code,
        }
        for _, row in df.iterrows()
    ]
    try:
        DB.executemany(sql, params_list=params_list )
        DB.close()
        print("Inventory table updated successfully.")
        return None
    except Exception as e:
        DB.close()
        print(f"Error updating inventory table: {e}")
        return e


@app.get("/inventory/update")
def update_inventory():
    try:
        df, excel_errors = load_inventory_data()

        if excel_errors:
            return JSONResponse({"status": "error from excel", "detail": excel_errors})
        db_error = update_inventory_table(df)

        if db_error:
            return JSONResponse({"status": "error from database", "detail": str(db_error)})

        return JSONResponse({"status": "finish"})
    except Exception as e:
        return JSONResponse({"status": "error", "detail": str(e)}, status_code=500)


@app.get("/hello", response_class=HTMLResponse)
def hello_api():
    return """
    <!DOCTYPE html>
    <html lang="vi">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Hello Page</title>
        <style>
            * {
                box-sizing: border-box;
            }

            body {
                margin: 0;
                min-height: 100vh;
                display: grid;
                place-items: center;
                font-family: Georgia, serif;
                background:
                    linear-gradient(135deg, #102a43, #243b53 55%, #486581);
                color: white;
            }

            .hello-page {
                width: min(90%, 620px);
                padding: 64px 48px;
                text-align: center;
                border: 1px solid rgba(255, 255, 255, 0.25);
                background: rgba(255, 255, 255, 0.12);
                box-shadow: 0 24px 80px rgba(0, 0, 0, 0.28);
                backdrop-filter: blur(12px);
            }

            h1 {
                margin: 0 0 16px;
                font-size: clamp(48px, 10vw, 92px);
                letter-spacing: 2px;
            }

            p {
                margin: 0;
                color: #d9e2ec;
                font-family: Arial, sans-serif;
                font-size: 18px;
            }
        </style>
    </head>
    <body>
        <main class="hello-page">
            <h1>Hello</h1>
            <p>Chào mừng bạn đến với CMMS App</p>
        </main>
    </body>
    </html>
    """


    
if __name__ == "__main__":
    # import uvicorn
    # uvicorn.run(
    #     "inventory_update_api:app",
    #     host="0.0.0.0",
    #     port=8000,
    #     workers=1,
    #     reload=False
    # )
    # update_inventory()
    df, excel_errors = load_inventory_data()
    print(df[df["code"] == "9000011403"])

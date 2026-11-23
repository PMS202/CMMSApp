import fitz  
import cv2
import numpy as np
import zxingcpp
import json
import pandas as pd
import os
import sys
from pathlib import Path
os.environ["FLAGS_use_mkldnn"] = "0"
os.environ["FLAGS_enable_pir_api"] = "0"
from paddleocr import PaddleOCR
import paddle
from sqlalchemy import text
# import paddlex
ocr_engine = PaddleOCR(
                        text_detection_model_name="PP-OCRv5_server_det",
                        text_recognition_model_name="PP-OCRv5_server_rec",
                        use_doc_orientation_classify=False,
                        use_doc_unwarping=False,
                        use_textline_orientation=False,
                        )

class Scan_record_process():
    def __init__(self):
        self.roi_size = {"x":100 , "y":100, "width":1500 , "height" : 1500}
        self.y = self.roi_size["y"]
        self.height = self.y + self.roi_size["height"]
        self.x = self.roi_size["x"]
        self.width = self.x + self.roi_size["width"]
        self.oneFile = None

    def paths(self,link):
        self.link = link
        self.path = None
        if self.link != "" and not self.link.lower().endswith(".pdf"):
            dirs = os.listdir(self.link)
            self.path = [os.path.join(self.link,dir) for dir in dirs]
            self.oneFile = False
        else:
            self.path = [self.link,]
            self.oneFile = True
        return self.path

    def OCR_oneFile(self,path):
        pdf = fitz.open(path)
        readed_code = []
        for page_num in range(pdf.page_count):
            page = pdf[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(8, 8)) 
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            if pix.n == 4:
                img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
            roi = img[self.y:self.height, self.x:self.width]
            # cv2.imshow("ROI", roi)
            # cv2.waitKey(0)
            result = ocr_engine.ocr(roi,cls=False)
            for line in result[0]:
                text = line[1][0]
                if "Code:" in text:
                    code = text.split("Code:")[-1].strip()
                    readed_code.append({
                        "Code": code,
                        "Page": page_num
                    })
                if "Attached equipment:" in text or "Attachedequipment:" in text:
                    code = text.split("Attached equipment:")[-1].strip()
                    code = text.split("Attachedequipment:")[-1].strip()
                    readed_code.append({
                        "Attached equipment": code,
                        "Page": page_num
                    })

        return readed_code

    def OCR_multipleFiles(self, paths):
        all_readed_codes = []
        for path in paths:
            readed_codes = self.OCR_oneFile(path)
            all_readed_codes.extend(readed_codes)
        return all_readed_codes

    def split_pdf(self,input_file, start, end, output_file):
        doc = fitz.open(input_file)
        new_doc = fitz.open()
        for i in range(start, end + 1):
            page = doc[i]
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5),alpha=False)  
            new_page = new_doc.new_page(width=pix.width, height=pix.height)
            new_page.insert_image(new_page.rect, pixmap=pix)
        new_doc.save(output_file,        
                     deflate=True,
                        garbage=4)
        new_doc.close()
        doc.close()

    
if __name__ == "__main__":
    scanner = Scan_record_process()
    PROJECT_ROOT = Path(__file__).resolve().parents[1]
    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    # from Database.MariaDB import Database_process
    # database = Database_process()
    # file_path = r"C:\Users\2173452100291\Downloads\maintenance_MA15_MA22_af5.xlsx"
    # try:
    #     df = pd.read_excel(file_path, sheet_name="Sheet1")
    #     # df = df[df["Line_name"] == "MA07"]
    # except Exception as e:
    #     print(f"Error reading Excel file: {e}")
    # paths = scanner.paths(r"C:\Users\2173452100291\Downloads\Scan\Af5")
    # for path in paths:
    #     df_filtered = df[df['File'] == path]
    #     print(f"Processing path: {path}, filtered DataFrame has {len(df_filtered)} rows")
    #     for idx, row in df_filtered.iterrows():
    #         file_name = f"{row['Machine_code']}_{row['Machine_name']}_PE5_{row['Line_name']}_Dien_{row['Maintenance_date']}.pdf"
    #         file_name = file_name.replace("/", "-")
    #         file_name = file_name.replace(" ", "_")
    #         start_page = int(row['Page'])
    #         end_page = start_page + int(row['Page_num']) - 1
    #         with database.Session() as session:
    #             try:
    #                 session.execute(text('''
    #                                 INSERT INTO `maintenance_records` ( machine_id, maintenance_date, technician, `next_due_date`, line_id, record_link )
    #                                             SELECT m.machine_id,
    #                                                 :maintenance_date,
    #                                                 :technical,
    #                                                 :next_date,
    #                                                 p.line_id,
    #                                                 :link
    #                                             FROM `machines` m
    #                                             JOIN `production_Lines` p ON p.line_name = :line
    #                                             WHERE m.machine_code = :code;
    #                                 '''),
    #                                 {
    #                                     'maintenance_date': row['Maintenance_date'],
    #                                     'technical': "Dien",
    #                                     'next_date': row['Next_date'],
    #                                     'link': os.path.join(r"X:\Scan_Result_X", file_name),
    #                                     'line': row['Line_name'],
    #                                     'code': row['Machine_code']
    #                                 }
    #                                 )
    #                 session.execute(text('''
    #                                         UPDATE maintenance_plan mp
    #                                         JOIN machines m ON m.machine_id = mp.machine_id
    #                                         JOIN production_Lines p ON p.line_id = mp.line_id
    #                                         SET mp.status = "Ontime",
    #                                             mp.maintenance_date = :maintenance_date,
    #                                             mp.week = :week
    #                                         WHERE m.machine_code = :machine_code
    #                                         AND p.line_name = :line_name AND mp.maintenance_date IS NULL AND mp.month_year_id = 21;
    #                 '''), {
    #                     'maintenance_date': row['Maintenance_date'],
    #                     'machine_code': row['Machine_code'],
    #                     'line_name': row['Line_name'],
    #                     'week': row['Week']
    #                 })
    #                 scanner.split_pdf(path, start_page, end_page, os.path.join(r"X:\Scan_Result_X", file_name))
    #                 session.commit()
    #             except Exception as e:
    #                 print(f"Error splitting PDF file {file_name}: {e}")
    #                 session.rollback()
    
    paths = scanner.paths(r"X:\19112026200409.pdf")
    readed_codes = scanner.OCR_multipleFiles(paths)

    df = pd.DataFrame(readed_codes)
    output_path = os.path.join(PROJECT_ROOT, 'exported_files', 'PE3_F18.xlsx')
    df.to_excel(output_path, index=False)



    
    # path = r"C:\Users\2173452100291\Downloads\Scan\Af2\15092026124842.pdf"
    # scanner.split_pdf(input_file=path, start=35, end=36, output_file=os.path.join(r"X:\Scan_Result_X",
    #                                                                                        "688195_DISPENSOR_ROBOT_SM200DS-3A_PE5_MA07_Tan_2026-09-09.pdf"))
    # scanner.split_pdf(input_file=path, start=4, end=7, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                        "JIG_T_MA07_2026-09-18.pdf"))
    # scanner.split_pdf(input_file=path, start=8, end=11, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                            "JIG_T_MA15_2026-09-11.pdf"))
    # scanner.split_pdf(input_file=path, start=12, end=17, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                            "JIG_T_MA16_2026-09-12.pdf"))
    # scanner.split_pdf(input_file=path, start=18, end=21, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                            "JIG_T_MA17_2026-09-14.pdf"))
    # scanner.split_pdf(input_file=path, start=22, end=25, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                                "JIG_T_MA18_2026-09-15.pdf"))
    # scanner.split_pdf(input_file=path, start=26, end=30, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                                "JIG_M_MA19_2026-09-19.pdf"))
    # scanner.split_pdf(input_file=path, start=31, end=35, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                                "JIG_M_MA20_2026-09-17.pdf"))
    # scanner.split_pdf(input_file=path, start=36, end=39, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                                "JIG_M_MA21_2026-09-18.pdf"))
    # scanner.split_pdf(input_file=path, start=40, end=43, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                                "JIG_M_MA22_2026-09-19.pdf"))

    # path = r"C:\Users\2173452100291\Downloads\Scan\Jig2\21092026131535.pdf"
    # scanner.split_pdf(input_file=path, start=0, end=3, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                        "JIG_T_MA06_2026-06-09.pdf"))
    # scanner.split_pdf(input_file=path, start=4, end=7, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                        "JIG_T_MA07_2026-06-18.pdf"))
    # scanner.split_pdf(input_file=path, start=8, end=11, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                            "JIG_T_MA15_2026-06-11.pdf"))
    # scanner.split_pdf(input_file=path, start=12, end=15, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                            "JIG_T_MA16_2026-06-12.pdf"))
    # scanner.split_pdf(input_file=path, start=16, end=19, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                            "JIG_T_MA17_2026-06-13.pdf"))
    # scanner.split_pdf(input_file=path, start=20, end=23, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                                "JIG_T_MA18_2026-06-15.pdf"))
    # scanner.split_pdf(input_file=path, start=24, end=25, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                                "JIG_M_MA22_2026-06-19.pdf"))

    # path = r"C:\Users\2173452100291\Downloads\Scan\Jig2\21092026131501.pdf"
    # scanner.split_pdf(input_file=path, start=0, end=3, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                        "JIG_T_MA06_2026-03-09.pdf"))
    # scanner.split_pdf(input_file=path, start=4, end=7, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                        "JIG_T_MA07_2026-03-18.pdf"))
    # scanner.split_pdf(input_file=path, start=8, end=11, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                            "JIG_T_MA15_2026-03-11.pdf"))
    # scanner.split_pdf(input_file=path, start=12, end=15, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                            "JIG_T_MA16_2026-03-12.pdf"))
    # scanner.split_pdf(input_file=path, start=16, end=19, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                            "JIG_T_MA17_2026-03-13.pdf"))
    # scanner.split_pdf(input_file=path, start=20, end=24, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                                "JIG_T_MA18_2026-03-14.pdf"))
    # scanner.split_pdf(input_file=path, start=25, end=26, output_file=os.path.join(r"C:\Users\2173452100291\Downloads\Scan\Jig2\split",
    #                                                                                                "JIG_M_MA22_2026-03-19.pdf"))